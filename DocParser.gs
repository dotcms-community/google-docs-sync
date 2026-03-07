/**
 * Document parsing: metadata table detection, field extraction, body HTML export.
 */
var DocParser = {

  MARKER_FIELD: 'Field',
  MARKER_VALUE: 'Value',

  /**
   * Find the metadata table by scanning for a table whose first row
   * contains 'Field' and 'Value' headers.
   * Returns {tableIndex, table} or null.
   */
  findMetadataTable: function () {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var tables = body.getTables();

    for (var i = 0; i < tables.length; i++) {
      var table = tables[i];
      if (table.getNumRows() < 1 || table.getRow(0).getNumCells() < 2) continue;

      var h1 = table.getRow(0).getCell(0).getText().trim();
      var h2 = table.getRow(0).getCell(1).getText().trim();

      if (h1 === this.MARKER_FIELD && h2 === this.MARKER_VALUE) {
        return { tableIndex: i, table: table };
      }
    }
    return null;
  },

  /**
   * Extract field/value pairs from the metadata table (skipping header row).
   */
  extractMetadataFields: function () {
    var result = this.findMetadataTable();
    if (!result) return [];

    var table = result.table;
    var fields = [];
    for (var r = 1; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      var fieldVar = row.getCell(0).getText().trim();
      var fieldVal = row.getCell(1).getText().trim();
      if (fieldVar) {
        fields.push({ variable: fieldVar, value: fieldVal });
      }
    }
    return fields;
  },

  /**
   * Generate the metadata table at the top of the document.
   * Only includes required fields (and non-layout, non-relationship fields get free text;
   * selects/booleans get dropdown notation).
   */
  generateMetadataTable: function (fields) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    // Filter to required fields only, exclude certain types
    var requiredFields = fields.filter(function (f) {
      return f.required;
    });

    // Build table: header + one row per required field
    var numRows = requiredFields.length + 1;
    var table = body.insertTable(0, this._buildTableCells(requiredFields));

    // Style header row
    var headerRow = table.getRow(0);
    headerRow.getCell(0).setText(this.MARKER_FIELD).setAttributes(this._boldStyle());
    headerRow.getCell(1).setText(this.MARKER_VALUE).setAttributes(this._boldStyle());

    // Add field rows
    for (var i = 0; i < requiredFields.length; i++) {
      var row = table.getRow(i + 1);
      row.getCell(0).setText(requiredFields[i].variable);

      var hint = this._getFieldHint(requiredFields[i]);
      if (hint) {
        row.getCell(1).setText(hint);
      }
    }

    // Insert a paragraph after the table for body content
    body.insertParagraph(1, '');

    return requiredFields;
  },

  /**
   * Update a specific field's value cell in the metadata table.
   * Called from sidebar when user selects a value from a dropdown or lookup.
   */
  updateMetadataField: function (fieldVariable, value) {
    var result = this.findMetadataTable();
    if (!result) return;

    var table = result.table;
    for (var r = 1; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      if (row.getCell(0).getText().trim() === fieldVariable) {
        row.getCell(1).setText(value);
        return;
      }
    }
  },

  _buildTableCells: function (requiredFields) {
    var cells = [['Field', 'Value']];
    for (var i = 0; i < requiredFields.length; i++) {
      cells.push([requiredFields[i].variable, '']);
    }
    return cells;
  },

  _getFieldHint: function (field) {
    // Select, boolean, and relationship fields are handled by the sidebar Field Editor.
    // No hints needed in the table — the sidebar writes values directly.
    return '';
  },

  _boldStyle: function () {
    var style = {};
    style[DocumentApp.Attribute.BOLD] = true;
    return style;
  },

  /**
   * Export the body content (everything outside the metadata table) as clean HTML.
   * Uses the Google Docs export API to get HTML, then strips the metadata table HTML.
   */
  getBodyHtml: function () {
    var doc = DocumentApp.getActiveDocument();
    var docId = doc.getId();

    // Export full doc as HTML
    var url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + docId + '&exportFormat=html';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var html = response.getContentText();

    // Clean the HTML
    html = HtmlCleaner.clean(html);

    // Remove the metadata table from the HTML
    html = this._removeFirstTable(html);

    return html;
  },

  /**
   * Remove the first <table> element from HTML (the metadata table).
   */
  _removeFirstTable: function (html) {
    var tableStart = html.indexOf('<table');
    if (tableStart === -1) return html;

    var tableEnd = html.indexOf('</table>', tableStart);
    if (tableEnd === -1) return html;

    return html.substring(0, tableStart) + html.substring(tableEnd + 8);
  },

  /**
   * Extract all inline images from the document body (excluding the metadata table).
   * Returns array of {blob, name, elementIndex, blobHash}.
   */
  getImages: function () {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var metaTable = this.findMetadataTable();
    var metaTableIndex = metaTable ? metaTable.tableIndex : -1;

    var images = [];
    var numChildren = body.getNumChildren();
    var imgCount = 0;

    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);

      // Skip the metadata table
      if (i === metaTableIndex) continue;

      this._collectImagesFromElement(child, images, imgCount);
    }
    return images;
  },

  _collectImagesFromElement: function (element, images, counter) {
    var type = element.getType();

    if (type === DocumentApp.ElementType.INLINE_IMAGE) {
      var blob = element.getBlob();
      var hash = this._hashBlob(blob);
      var name = element.getAltTitle() || ('image_' + images.length + '.' + this._getExtension(blob.getContentType()));
      images.push({ blob: blob, name: name, blobHash: hash });
    } else if (type === DocumentApp.ElementType.INLINE_DRAWING) {
      // Drawings don't have a direct blob export in Apps Script;
      // we can attempt to get them via the element if possible
      // This is a known limitation — we flag it.
      try {
        var blob = element.getBlob();
        if (blob) {
          var hash = this._hashBlob(blob);
          images.push({ blob: blob, name: 'drawing_' + images.length + '.png', blobHash: hash });
        }
      } catch (e) {
        // Drawings may not support getBlob; skip gracefully
      }
    }

    // Recurse into containers (paragraphs, list items, table cells, etc.)
    if (typeof element.getNumChildren === 'function') {
      for (var c = 0; c < element.getNumChildren(); c++) {
        this._collectImagesFromElement(element.getChild(c), images, counter);
      }
    }
  },

  _hashBlob: function (blob) {
    var bytes = blob.getBytes();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes);
    return digest.map(function (b) {
      return ('0' + (b & 0xFF).toString(16)).slice(-2);
    }).join('');
  },

  _getExtension: function (contentType) {
    var map = {
      'image/png': 'png',
      'image/jpeg': 'jpg',
      'image/gif': 'gif',
      'image/svg+xml': 'svg',
      'image/webp': 'webp'
    };
    return map[contentType] || 'png';
  }
};
