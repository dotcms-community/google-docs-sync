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
    var numChildren = body.getNumChildren();

    for (var i = 0; i < numChildren; i++) {
      var child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.TABLE) continue;

      var table = child.asTable();
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
  IDENTIFIER_FIELD: 'identifier',

  /**
   * Generate or merge the metadata table.
   * If a table already exists, only add missing required fields (preserving existing values).
   * If no table exists, create a new one.
   */
  generateMetadataTable: function (fields) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    var requiredFields = fields.filter(function (f) {
      return f.required;
    });

    var existing = this.findMetadataTable();

    if (existing) {
      // Merge: add rows for fields not already in the table
      var table = existing.table;
      var existingVars = {};
      for (var r = 1; r < table.getNumRows(); r++) {
        var varName = table.getRow(r).getCell(0).getText().trim();
        if (varName) existingVars[varName] = true;
      }

      // Ensure identifier row exists
      if (!existingVars['identifier']) {
        var idRow = table.appendTableRow();
        idRow.appendTableCell('identifier');
        idRow.appendTableCell('');
      }

      // Ensure contentType row exists
      if (!existingVars['contentType']) {
        var ctRow = table.appendTableRow();
        ctRow.appendTableCell('contentType');
        ctRow.appendTableCell('');
      }

      // Ensure default rows exist
      var defaults = ['host', 'folder', 'languageId'];
      for (var d = 0; d < defaults.length; d++) {
        if (!existingVars[defaults[d]]) {
          var defRow = table.appendTableRow();
          defRow.appendTableCell(defaults[d]);
          defRow.appendTableCell('');
        }
      }

      // Add missing required fields
      for (var i = 0; i < requiredFields.length; i++) {
        if (!existingVars[requiredFields[i].variable]) {
          var newRow = table.appendTableRow();
          newRow.appendTableCell(requiredFields[i].variable);
          newRow.appendTableCell('');
        }
      }
    } else {
      // Create new table
      var table = body.insertTable(0, this._buildTableCells(requiredFields));

      // Style header row
      var headerRow = table.getRow(0);
      headerRow.getCell(0).setText(this.MARKER_FIELD).setAttributes(this._boldStyle());
      headerRow.getCell(1).setText(this.MARKER_VALUE).setAttributes(this._boldStyle());

      // Insert a paragraph after the table for body content
      body.insertParagraph(1, '');
    }

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

    // Row not found — append it
    var newRow = table.appendTableRow();
    newRow.appendTableCell(fieldVariable);
    newRow.appendTableCell(value);
  },

  _buildTableCells: function (requiredFields) {
    var cells = [['Field', 'Value'], ['identifier', ''], ['contentType', ''], ['host', ''], ['folder', ''], ['languageId', '']];
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
   * Builds HTML directly from DocumentApp elements — no external API needed.
   */
  getBodyHtml: function () {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var metaTable = this.findMetadataTable();
    var metaTableIndex = metaTable ? metaTable.tableIndex : -1;
    var numChildren = body.getNumChildren();
    var html = '';

    for (var i = 0; i < numChildren; i++) {
      if (i === metaTableIndex) continue;
      html += this._elementToHtml(body.getChild(i));
    }

    return html.trim();
  },

  /**
   * Convert a DocumentApp element to HTML.
   */
  _elementToHtml: function (element) {
    var type = element.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      return this._paragraphToHtml(element.asParagraph());
    }
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      return this._listItemToHtml(element.asListItem());
    }
    if (type === DocumentApp.ElementType.TABLE) {
      return this._tableToHtml(element.asTable());
    }
    if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      return '<hr>';
    }

    // Fallback: extract text
    var text = element.asText ? element.asText().getText() : '';
    if (text) return '<p>' + this._escapeHtml(text) + '</p>';
    return '';
  },

  _paragraphToHtml: function (para) {
    var heading = para.getHeading();
    var content = this._renderInlineContent(para);
    if (!content && heading === DocumentApp.ParagraphHeading.NORMAL) return '';

    var tag = 'p';
    switch (heading) {
      case DocumentApp.ParagraphHeading.HEADING1: tag = 'h1'; break;
      case DocumentApp.ParagraphHeading.HEADING2: tag = 'h2'; break;
      case DocumentApp.ParagraphHeading.HEADING3: tag = 'h3'; break;
      case DocumentApp.ParagraphHeading.HEADING4: tag = 'h4'; break;
      case DocumentApp.ParagraphHeading.HEADING5: tag = 'h5'; break;
      case DocumentApp.ParagraphHeading.HEADING6: tag = 'h6'; break;
    }
    return '<' + tag + '>' + content + '</' + tag + '>';
  },

  _listItemToHtml: function (item) {
    var glyph = item.getGlyphType();
    var content = this._renderInlineContent(item);
    // Simple list rendering — nested lists would need tracking state across calls
    var isOrdered = (glyph === DocumentApp.GlyphType.NUMBER ||
                     glyph === DocumentApp.GlyphType.LATIN_UPPER ||
                     glyph === DocumentApp.GlyphType.LATIN_LOWER ||
                     glyph === DocumentApp.GlyphType.ROMAN_UPPER ||
                     glyph === DocumentApp.GlyphType.ROMAN_LOWER);
    var tag = isOrdered ? 'ol' : 'ul';
    return '<' + tag + '><li>' + content + '</li></' + tag + '>';
  },

  _tableToHtml: function (table) {
    var html = '<table>';
    for (var r = 0; r < table.getNumRows(); r++) {
      html += '<tr>';
      var row = table.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        html += '<td>' + this._escapeHtml(cell.getText()) + '</td>';
      }
      html += '</tr>';
    }
    html += '</table>';
    return html;
  },

  /**
   * Render inline content (text with formatting + inline images) from a paragraph or list item.
   */
  _renderInlineContent: function (element) {
    var html = '';
    var numChildren = element.getNumChildren();

    for (var i = 0; i < numChildren; i++) {
      var child = element.getChild(i);
      var childType = child.getType();

      if (childType === DocumentApp.ElementType.TEXT) {
        html += this._textToHtml(child.asText());
      } else if (childType === DocumentApp.ElementType.INLINE_IMAGE) {
        // Placeholder src — will be replaced by ImageHandler with dotAsset URLs
        html += '<img src="image_placeholder_' + i + '">';
      } else if (childType === DocumentApp.ElementType.INLINE_DRAWING) {
        html += '<img src="drawing_placeholder_' + i + '">';
      }
    }
    return html;
  },

  /**
   * Convert a Text element to HTML, preserving bold, italic, underline, links.
   */
  _textToHtml: function (textEl) {
    var text = textEl.getText();
    if (!text) return '';

    var html = '';
    var i = 0;
    while (i < text.length) {
      // Find runs of same formatting
      var bold = textEl.isBold(i);
      var italic = textEl.isItalic(i);
      var underline = textEl.isUnderline(i);
      var strikethrough = textEl.isStrikethrough(i);
      var link = textEl.getLinkUrl(i);

      var j = i + 1;
      while (j < text.length &&
             textEl.isBold(j) === bold &&
             textEl.isItalic(j) === italic &&
             textEl.isUnderline(j) === underline &&
             textEl.isStrikethrough(j) === strikethrough &&
             textEl.getLinkUrl(j) === link) {
        j++;
      }

      var chunk = this._escapeHtml(text.substring(i, j));

      if (link) chunk = '<a href="' + this._escapeHtml(link) + '">' + chunk + '</a>';
      if (bold) chunk = '<strong>' + chunk + '</strong>';
      if (italic) chunk = '<em>' + chunk + '</em>';
      if (underline) chunk = '<u>' + chunk + '</u>';
      if (strikethrough) chunk = '<s>' + chunk + '</s>';

      html += chunk;
      i = j;
    }
    return html;
  },

  _escapeHtml: function (str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
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
