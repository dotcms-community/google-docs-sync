/**
 * Handles image upload deduplication and dotAsset creation.
 * Tracks image blob hashes → dotAsset identifiers in document properties.
 */
var ImageHandler = {

  PROP_IMAGE_MAP: 'dotcms_image_map',

  /**
   * Get the stored image hash → dotAsset ID map from document properties.
   */
  getImageMap: function () {
    var props = PropertiesService.getDocumentProperties();
    var raw = props.getProperty(this.PROP_IMAGE_MAP);
    if (!raw) return {};
    try {
      return JSON.parse(raw);
    } catch (e) {
      return {};
    }
  },

  /**
   * Save the image map to document properties.
   */
  saveImageMap: function (map) {
    var props = PropertiesService.getDocumentProperties();
    props.setProperty(this.PROP_IMAGE_MAP, JSON.stringify(map));
  },

  /**
   * Upload images to dotCMS, skipping already-uploaded ones (by hash).
   * Returns { uploaded: [...], failed: [...], imageMap: {...} }
   */
  uploadImages: function (images, host, token, siteId, folderPath, languageId, progressCallback) {
    var imageMap = this.getImageMap();
    var uploaded = [];
    var failed = [];

    for (var i = 0; i < images.length; i++) {
      var img = images[i];

      // Skip if already uploaded (hash match) — but only if the cached identifier looks valid
      var cached = imageMap[img.blobHash];
      if (cached && cached.identifier && cached.identifier.length > 10) {
        uploaded.push({
          name: img.name,
          blobHash: img.blobHash,
          dotAssetId: cached.identifier,
          skipped: true
        });
        continue;
      }

      try {
        // Step 1: Upload to Temp API
        var tempId = DotCMSApi.uploadTemp(host, token, img.blob, img.name);

        // Step 2: Create dotAsset
        var assetEntity = DotCMSApi.createDotAsset(host, token, tempId, siteId, folderPath, languageId);
        var assetId = _extractIdentifier(assetEntity);

        if (!assetId) {
          throw new Error('No identifier in dotAsset response: ' + JSON.stringify(assetEntity).substring(0, 200));
        }

        imageMap[img.blobHash] = {
          identifier: assetId,
          inode: assetEntity.inode || '',
          fileName: assetEntity.fileName || ''
        };

        uploaded.push({
          name: img.name,
          blobHash: img.blobHash,
          dotAssetId: assetId,
          skipped: false
        });
      } catch (e) {
        failed.push({
          name: img.name,
          blobHash: img.blobHash,
          error: e.message || String(e)
        });
      }

      if (progressCallback) {
        progressCallback(i + 1, images.length);
      }
    }

    // Persist updated map
    this.saveImageMap(imageMap);

    return { uploaded: uploaded, failed: failed, imageMap: imageMap };
  },

  /**
   * Replace image references in HTML with dotAsset URLs.
   * Google exported HTML has images as data: URIs or googleusercontent URLs.
   * We replace them in order with the corresponding dotAsset URLs.
   */
  replaceImageUrls: function (html, images, imageMap, host) {
    var imgTagRegex = /<img[^>]+src="([^"]+)"[^>]*>/gi;
    var imgIndex = 0;

    html = html.replace(imgTagRegex, function (match, src) {
      if (imgIndex < images.length) {
        var img = images[imgIndex];
        imgIndex++;
        var assetInfo = imageMap[img.blobHash];
        if (assetInfo) {
          var newSrc = host + '/dA/' + assetInfo.identifier;
          return match.replace(src, newSrc);
        }
      }
      return match;
    });

    return html;
  }
};
