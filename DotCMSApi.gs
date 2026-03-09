/**
 * dotCMS REST API wrapper.
 */
var DotCMSApi = {

  // ── Content Types ──

  getContentTypes: function (host, token) {
    var resp = this._get(host, token, '/api/v1/contenttype?per_page=100&orderby=name');
    var entities = resp.entity || [];
    return entities.map(function (ct) {
      return { name: ct.name, variable: ct.variable, id: ct.id };
    });
  },

  getContentTypeFields: function (host, token, contentTypeVar) {
    var resp = this._get(host, token, '/api/v1/contenttype/' + contentTypeVar + '/fields');
    var allFields = resp.entity || [];
    return allFields.map(function (f) {
      return {
        variable: f.variable,
        name: f.name,
        fieldType: f.fieldType || f.clazz || '',
        required: !!f.required,
        values: f.values || '',
        dataType: f.dataType || '',
        relationType: f.relationType || '',
        cardinality: (typeof f.cardinality !== 'undefined') ? f.cardinality : null
      };
    }).filter(function (f) {
      // Exclude layout fields (rows, columns, tabs, line dividers)
      var skip = [
        'com.dotcms.contenttype.model.field.impl.RowField',
        'com.dotcms.contenttype.model.field.impl.ColumnField',
        'com.dotcms.contenttype.model.field.impl.TabDividerField',
        'com.dotcms.contenttype.model.field.impl.LineDividerField',
        'com.dotcms.contenttype.model.field.impl.PermissionTabField',
        'com.dotcms.contenttype.model.field.impl.RelationshipsTabField',
        'com.dotcms.contenttype.model.field.impl.HostFolderField'
      ];
      return skip.indexOf(f.fieldType) === -1;
    });
  },

  // ── Sites & Folders ──

  getSites: function (host, token) {
    var resp = this._get(host, token, '/api/v1/site?per_page=50');
    var entities = resp.entity || [];
    return entities.map(function (s) {
      return { name: s.hostname, identifier: s.identifier };
    });
  },

  getFolders: function (host, token, siteId) {
    var resp = this._get(host, token, '/api/v1/folder/siteid/' + siteId + '?per_page=200');
    var entities = resp.entity || [];
    return entities.map(function (f) {
      return { name: f.name, path: f.path, inode: f.inode };
    });
  },

  // ── Languages ──

  getLanguages: function (host, token) {
    var resp = this._get(host, token, '/api/v2/languages');
    var entities = resp.entity || [];
    return entities.map(function (l) {
      return { id: l.id, language: l.language, country: l.country, languageCode: l.languageCode };
    });
  },

  // ── Content Search (for relationship lookups) ──

  searchContent: function (host, token, query) {
    var resp = this._get(host, token, '/api/content/_search', {
      query: query,
      limit: 20
    });
    var contentlets = (resp.entity && resp.entity.jsonObjectView && resp.entity.jsonObjectView.contentlets) || resp.contentlets || [];
    return contentlets.map(function (c) {
      return { identifier: c.identifier, title: c.title || c.name || c.identifier, contentType: c.contentType };
    });
  },

  // ── Temp File Upload ──

  uploadTemp: function (host, token, blob, fileName) {
    var boundary = '----FormBoundary' + Utilities.getUuid();
    var payload = Utilities.newBlob('').getBytes();

    var pre = '--' + boundary + '\r\n' +
      'Content-Disposition: form-data; name="file"; filename="' + fileName + '"\r\n' +
      'Content-Type: ' + blob.getContentType() + '\r\n\r\n';
    var post = '\r\n--' + boundary + '--\r\n';

    var preBytes = Utilities.newBlob(pre).getBytes();
    var fileBytes = blob.getBytes();
    var postBytes = Utilities.newBlob(post).getBytes();

    var allBytes = [];
    allBytes = allBytes.concat(preBytes).concat(fileBytes).concat(postBytes);

    var options = {
      method: 'post',
      contentType: 'multipart/form-data; boundary=' + boundary,
      headers: { 'Authorization': 'Bearer ' + token, 'Origin': host },
      payload: allBytes,
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(host + '/api/v1/temp', options);
    var json = JSON.parse(response.getContentText());
    if (json.tempFiles && json.tempFiles.length > 0) {
      return json.tempFiles[0].id;
    }
    throw new Error('Temp upload failed: ' + response.getContentText());
  },

  // ── Create dotAsset ──

  createDotAsset: function (host, token, tempId, siteId, folderPath, languageId) {
    var payload = {
      contentlet: {
        contentType: 'dotAsset',
        asset: tempId,
        hostFolder: siteId,
        folder: folderPath || '',
        languageId: languageId || 1
      }
    };

    var resp = this._post(host, token, '/api/v1/workflow/actions/default/fire/PUBLISH', payload);
    var entity = resp.entity || {};

    return entity;
  },

  // ── Fire Workflow (Save or Publish content) ──

  fireWorkflow: function (host, token, action, contentletData) {
    var endpoint = '/api/v1/workflow/actions/default/fire/' + action;
    var payload = { contentlet: contentletData };
    var resp = this._post(host, token, endpoint, payload);
    var entity = resp.entity || {};

    return entity;
  },

  // ── HTTP helpers ──

  _get: function (host, token, path, params) {
    var url = host + path;
    if (params) {
      var qs = Object.keys(params).map(function (k) {
        return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
      }).join('&');
      url += (url.indexOf('?') === -1 ? '?' : '&') + qs;
    }
    var options = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token, 'Origin': host },
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  },

  _post: function (host, token, path, payload) {
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token, 'Origin': host },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(host + path, options);
    var text = response.getContentText();
    Logger.log('POST ' + path + ' response (first 500): ' + text.substring(0, 500));
    return JSON.parse(text);
  }
};
