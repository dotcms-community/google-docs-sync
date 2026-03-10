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
  if (typeof entity.identifier === 'string' && entity.identifier) return entity.identifier;

  // Shape 3: results array — dig into first result
  if (Array.isArray(entity.results) && entity.results.length > 0) {
    var firstResult = entity.results[0];
    if (typeof firstResult.identifier === 'string' && firstResult.identifier) return firstResult.identifier;
    // Result is { "<identifier>": { ...contentlet... } }
    var resultKeys = Object.keys(firstResult);
    for (var ri = 0; ri < resultKeys.length; ri++) {
      var rVal = firstResult[resultKeys[ri]];
      // The value should be a contentlet object with baseType or identifier
      if (rVal && typeof rVal === 'object' && !Array.isArray(rVal)) {
        if (rVal.identifier) return rVal.identifier;
        if (rVal.baseType) return resultKeys[ri];
      }
    }
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

      // Step 5b: Upload images from image/file metadata fields
      var ctFields = DotCMSApi.getContentTypeFields(s.hostUrl, s.apiToken, options.contentType);
      var imageFieldVars = [];
      for (var fi = 0; fi < ctFields.length; fi++) {
        var ftLower = ctFields[fi].fieldType.toLowerCase().replace(/-/g, '');
        if (ftLower.indexOf('image') !== -1 || ftLower.indexOf('file') !== -1 || ftLower.indexOf('binary') !== -1) {
          imageFieldVars.push(ctFields[fi].variable);
        }
      }

      if (imageFieldVars.length > 0) {
        var metaImages = DocParser.getMetadataImageFields(imageFieldVars);
        for (var mi = 0; mi < metaImages.length; mi++) {
          var metaImg = metaImages[mi];

          // If the cell already has an identifier below the image, skip upload
          if (metaImg.existingId && metaImg.existingId.length > 10) {
            result.imagesSkipped++;
            continue;
          }

          try {
            // Check blob hash cache first
            var cachedAsset = imgResult.imageMap[metaImg.blobHash];
            var assetId = '';
            if (cachedAsset && cachedAsset.identifier && cachedAsset.identifier.length > 10) {
              assetId = cachedAsset.identifier;
            } else {
              var tempId = DotCMSApi.uploadTemp(s.hostUrl, s.apiToken, metaImg.blob, metaImg.name);
              var assetEntity = DotCMSApi.createDotAsset(s.hostUrl, s.apiToken, tempId, options.siteId, options.folderPath || '', options.languageId || 1);
              assetId = _extractIdentifier(assetEntity);
              if (assetId) {
                imgResult.imageMap[metaImg.blobHash] = { identifier: assetId };
              }
            }
            if (assetId) {
              // Write identifier below the image, keeping the image visible
              DocParser.setImageFieldIdentifier(metaImg.variable, assetId);
              result.imagesUploaded++;
            } else {
              imgResult.failed.push({ name: metaImg.name, error: 'No identifier returned for image field' });
            }
          } catch (e) {
            imgResult.failed.push({ name: metaImg.name, error: e.message || String(e) });
          }
        }
        // Save updated image map
        ImageHandler.saveImageMap(imgResult.imageMap);
      }

      // Re-read metadata fields after image field updates
      metadataFields = DocParser.extractMetadataFields();

      // Step 6: Build the contentlet payload from metadata table
      // The metadata table is the single source of truth for all field values.
      // Sidebar options are fallbacks only.
      var contentlet = {};
      var existingId = '';

      // Identify relationship fields
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
        if (mf.variable === 'inode' || mf.variable === 'dotBodyField') {
          continue;
        }

        // Skip hint values like "[option1 | option2]"
        if (!mf.value || (mf.value.charAt(0) === '[' && mf.value.charAt(mf.value.length - 1) === ']')) {
          continue;
        }

        // Relationship fields: send as Lucene query for dotCMS
        if (relFieldMap[mf.variable]) {
          var ids = mf.value.split(',').map(function (id) { return id.trim(); }).filter(function (id) { return id; });
          if (ids.length > 0) {
            contentlet[mf.variable] = '+identifier:(' + ids.join(' OR ') + ')';
          }
        } else {
          contentlet[mf.variable] = mf.value;
        }
      }

      // Apply sidebar fallbacks for required fields if not in table
      if (!contentlet.contentType) contentlet.contentType = options.contentType;
      // Detect if this content type has a custom HostFolderField
      var customHostVar = null;
      for (var hf = 0; hf < ctFields.length; hf++) {
        var hfType = ctFields[hf].fieldType.toLowerCase().replace(/-/g, '');
        if (hfType.indexOf('hostfolder') !== -1) {
          customHostVar = ctFields[hf].variable;
          break;
        }
      }
      // Only add host fallback if there's no custom HostFolderField already in the contentlet
      if (customHostVar && contentlet[customHostVar]) {
        // Custom host field is already set — don't add 'host'
      } else if (!contentlet.host) {
        contentlet.host = options.siteId;
      }
      if (!contentlet.languageId) contentlet.languageId = options.languageId || 1;

      // Add body content to the designated body field
      contentlet[options.bodyField || 'body'] = bodyHtml;

      if (existingId) {
        contentlet.identifier = existingId;
      }

      // Step 7: Fire workflow action (SAVE or PUBLISH)
      var action = options.publishMode === 'PUBLISH' ? 'PUBLISH' : 'EDIT';

      var fireResult = DotCMSApi.fireWorkflow(s.hostUrl, s.apiToken, action, contentlet);

      // Extract identifier from response.
      // PUBLISH returns: { entity: { identifier: "...", ... } }
      // EDIT returns:    { entity: { "<identifier>": { ... } } }
      var returnedId = '';
      var returnedInode = '';

      // Debug: store raw response for inspection
      try {
        PropertiesService.getDocumentProperties().setProperty('_debug_fireResult',
          JSON.stringify(fireResult).substring(0, 500));
      } catch(de) {}

      returnedId = _extractIdentifier(fireResult);
      // Safety: ensure returnedId is a string, not an object
      if (returnedId && typeof returnedId === 'object') {
        returnedId = returnedId.identifier || '';
      }
      if (fireResult.identifier && typeof fireResult.identifier === 'string') {
        returnedInode = fireResult.inode || '';
      }

      // Write the identifier and inode back to the metadata table so subsequent syncs update
      if (returnedId && typeof returnedId === 'string') {
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
        var debugResp = '';
        try { debugResp = JSON.stringify(fireResult).substring(0, 300); } catch(de) {}
        result.message = 'Content sent but no identifier returned. Response: ' + debugResp;
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
