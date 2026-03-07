/**
 * Orchestrates the full sync process: metadata extraction, body HTML,
 * image upload, and content push to dotCMS.
 */
var SyncEngine = {

  PROP_DOTCMS_ID: 'dotcms_content_identifier',
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

      // Step 6: Build the contentlet payload
      var contentlet = {
        contentType: options.contentType,
        host: options.siteId,
        folder: options.folderPath || '',
        languageId: options.languageId || 1
      };

      // Add metadata fields
      for (var i = 0; i < metadataFields.length; i++) {
        var mf = metadataFields[i];
        if (mf.variable && mf.value) {
          contentlet[mf.variable] = mf.value;
        }
      }

      // Add body content to the designated body field
      contentlet[options.bodyField || 'body'] = bodyHtml;

      // Check if we're updating an existing content item
      var existingId = this._getStoredIdentifier();
      if (existingId) {
        contentlet.identifier = existingId;
      }

      // Step 7: Fire workflow action (SAVE or PUBLISH)
      var action = options.publishMode === 'PUBLISH' ? 'PUBLISH' : 'SAVE';
      var entity = DotCMSApi.fireWorkflow(s.hostUrl, s.apiToken, action, contentlet);

      // Store the content identifier for future updates
      if (entity.identifier) {
        this._storeIdentifier(entity.identifier);
        result.contentIdentifier = entity.identifier;
      }

      // Compile results
      result.imagesUploaded = imgResult.uploaded.filter(function (u) { return !u.skipped; }).length;
      result.imagesSkipped = imgResult.uploaded.filter(function (u) { return u.skipped; }).length;
      result.imagesFailed = imgResult.failed.length;
      result.failures = imgResult.failed;

      if (imgResult.failed.length > 0) {
        result.status = 'partial';
        result.message = 'Content synced but ' + imgResult.failed.length + ' image(s) failed to upload.';
      } else {
        result.message = existingId ? 'Content updated successfully.' : 'Content created successfully.';
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

  _getStoredIdentifier: function () {
    var props = PropertiesService.getDocumentProperties();
    return props.getProperty(this.PROP_DOTCMS_ID) || '';
  },

  _storeIdentifier: function (identifier) {
    var props = PropertiesService.getDocumentProperties();
    props.setProperty(this.PROP_DOTCMS_ID, identifier);
  },

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
    if (log.length > 10) log = log.slice(0, 10);
    props.setProperty(this.PROP_SYNC_LOG, JSON.stringify(log));
  },

  getSyncLog: function () {
    var props = PropertiesService.getDocumentProperties();
    var raw = props.getProperty(this.PROP_SYNC_LOG);
    if (!raw) return [];
    try { return JSON.parse(raw); } catch (e) { return []; }
  }
};
