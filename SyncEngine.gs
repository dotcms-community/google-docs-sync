/**
 * Extract the dotCMS identifier from a workflow response entity.
 * If entity.identifier exists, use it.
 * Otherwise, the identifier is the key whose value is a non-array object.
 */
/**
 * Extract the dotCMS identifier from a workflow response entity.
 * Response shapes:
 *   1. { identifier: "..." }                          — identifier is a property
 *   2. { "<identifier>": { ...contentlet... } }       — identifier is the key
 *   3. { results: [{ "<identifier>": { ... } }], summary: { ... } } — nested in results array
 */
function _extractIdentifier(entity) {
  if (!entity) return '';

  // Shape 1: identifier is a direct property
  if (entity.identifier) return entity.identifier;

  // Shape 3: results array — dig into first result
  if (Array.isArray(entity.results) && entity.results.length > 0) {
    var firstResult = entity.results[0];
    if (firstResult.identifier) return firstResult.identifier;
    // Result is { "<identifier>": { ...contentlet... } }
    var resultKeys = Object.keys(firstResult);
    if (resultKeys.length > 0) return resultKeys[0];
  }

  // Shape 2: identifier is the key of a non-array object value
  var keys = Object.keys(entity);
  for (var i = 0; i < keys.length; i++) {
    var val = entity[keys[i]];
    if (val && typeof val === 'object' && !Array.isArray(val) && val.baseType) {
      return keys[i];
    }
  }
  return '';
}

/**
 * Orchestrates the full sync process: metadata extraction, body HTML,
 * image upload, and content push to dotCMS.
 */
var SyncEngine = {

  PROP_SYNC_LOG: 'dotcms_sync_log',

  /**
   * Main sync entry point.
   * @param {Object} options - { contentType, siteId, folderPath, bodyField, languageId, publishMode }
   * @returns {Object} - sync result with status, message, failures
   */
  sync: function (options) {
    var startTime = new Date();
    var s = SettingsService.getAll();

    if (!s.hostUrl || !s.apiToken) {
      return this._fail('dotCMS host URL and API token are required. Configure in Settings tab.');
    }

    var result = {
      status: 'success',
      message: '',
      failures: [],
      contentIdentifier: '',
      imagesUploaded: 0,
      imagesSkipped: 0,
      imagesFailed: 0
    };

    try {
      // Step 1: Extract metadata fields from the table
      var metadataFields = DocParser.extractMetadataFields();

      // Step 2: Get body HTML
      var bodyHtml = DocParser.getBodyHtml();

      // Step 3: Get images from the document
      var images = DocParser.getImages();

      // Step 4: Upload images (with deduplication)
      var imgResult = { uploaded: [], failed: [], imageMap: {} };
      if (images.length > 0) {
        imgResult = ImageHandler.uploadImages(
          images,
          s.hostUrl,
          s.apiToken,
          options.siteId,
          options.folderPath || '',
          options.languageId || 1
        );
      }

      // Step 5: Replace image URLs in body HTML
      if (images.length > 0) {
        bodyHtml = ImageHandler.replaceImageUrls(bodyHtml, images, imgResult.imageMap, s.hostUrl);
      }

      // Step 6: Build the contentlet payload from metadata table
      // The metadata table is the single source of truth for all field values.
      // Sidebar options are fallbacks only.
      var contentlet = {};
      var relationships = {};
      var existingId = '';

      // Get content type fields to identify relationship fields
      var ctFields = DotCMSApi.getContentTypeFields(s.hostUrl, s.apiToken, options.contentType);
      var relFieldMap = {};
      for (var rf = 0; rf < ctFields.length; rf++) {
        if (ctFields[rf].relationshipVelocityVar) {
          relFieldMap[ctFields[rf].variable] = ctFields[rf].relationshipVelocityVar;
        }
      }

      for (var i = 0; i < metadataFields.length; i++) {
        var mf = metadataFields[i];
        if (!mf.variable) continue;

        if (mf.variable === 'identifier') {
          if (mf.value) existingId = mf.value;
          continue;
        }
        if (mf.variable === 'inode') {
          continue;
        }

        // Skip hint values like "[option1 | option2]"
        if (!mf.value || (mf.value.charAt(0) === '[' && mf.value.charAt(mf.value.length - 1) === ']')) {
          continue;
        }

        // Relationship fields go into the relationships object
        if (relFieldMap[mf.variable]) {
          var velVar = relFieldMap[mf.variable];
          var ids = mf.value.split(',').map(function (id) { return id.trim(); }).filter(function (id) { return id; });
          relationships[velVar] = ids;
        } else {
          contentlet[mf.variable] = mf.value;
        }
      }

      // Apply sidebar fallbacks for required fields if not in table
      if (!contentlet.contentType) contentlet.contentType = options.contentType;
      // Host can be under 'host' or a custom HostFolderField variable
      var hasHost = Object.keys(contentlet).some(function (k) {
        return contentlet[k] === options.siteId;
      });
      if (!contentlet.host && !hasHost) contentlet.host = options.siteId;
      if (!contentlet.languageId) contentlet.languageId = options.languageId || 1;

      // Add body content to the designated body field
      contentlet[options.bodyField || 'body'] = bodyHtml;

      if (existingId) {
        contentlet.identifier = existingId;
      }

      // Step 7: Fire workflow action (SAVE or PUBLISH)
      var action = options.publishMode === 'PUBLISH' ? 'PUBLISH' : 'EDIT';

      // Attach relationships to the contentlet payload if any exist
      if (Object.keys(relationships).length > 0) {
        contentlet.relationships = relationships;
      }

      var fireResult = DotCMSApi.fireWorkflow(s.hostUrl, s.apiToken, action, contentlet);

      // Extract identifier from response.
      // PUBLISH returns: { entity: { identifier: "...", ... } }
      // EDIT returns:    { entity: { "<identifier>": { ... } } }
      var returnedId = '';
      var returnedInode = '';

      returnedId = _extractIdentifier(fireResult);
      if (fireResult.identifier) returnedInode = fireResult.inode || '';

      // Write the identifier and inode back to the metadata table so subsequent syncs update
      if (returnedId) {
        DocParser.updateMetadataField('identifier', returnedId);
        result.contentIdentifier = returnedId;
      }
      if (returnedInode) {
        DocParser.updateMetadataField('inode', returnedInode);
      }

      // Compile results
      result.imagesUploaded = imgResult.uploaded.filter(function (u) { return !u.skipped; }).length;
      result.imagesSkipped = imgResult.uploaded.filter(function (u) { return u.skipped; }).length;
      result.imagesFailed = imgResult.failed.length;
      result.failures = imgResult.failed;

      if (imgResult.failed.length > 0) {
        result.status = 'partial';
        result.message = 'Content synced but ' + imgResult.failed.length + ' image(s) failed to upload.';
      } else if (!returnedId) {
        result.status = 'partial';
        result.message = 'Content sent but no identifier returned.';
      } else {
        result.message = existingId ? 'Content updated successfully.' : 'Content created — identifier: ' + returnedId;
      }

    } catch (e) {
      result.status = 'error';
      result.message = 'Sync failed: ' + (e.message || String(e));
    }

    // Save sync log
    var endTime = new Date();
    result.timestamp = endTime.toISOString();
    result.duration = (endTime.getTime() - startTime.getTime()) / 1000;
    this._saveSyncLog(result);

    return result;
  },

  // ── Document property helpers ──

  _saveSyncLog: function (result) {
    var props = PropertiesService.getDocumentProperties();
    var log = [];
    var raw = props.getProperty(this.PROP_SYNC_LOG);
    if (raw) {
      try { log = JSON.parse(raw); } catch (e) { log = []; }
    }
    // Keep last 10 entries
    log.unshift({
      timestamp: result.timestamp,
      status: result.status,
      message: result.message,
      failures: result.failures,
      duration: result.duration,
      contentIdentifier: result.contentIdentifier
    });
    if (log.length > 5) log = log.slice(0, 5);
    props.setProperty(this.PROP_SYNC_LOG, JSON.stringify(log));
  },

  getSyncLog: function () {
    var props = PropertiesService.getDocumentProperties();
    var raw = props.getProperty(this.PROP_SYNC_LOG);
    if (!raw) return [];
    try { return JSON.parse(raw); } catch (e) { return []; }
  }
};
