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
   * If no table exists, create a new one with system + required fields only.
   */
  generateMetadataTable: function (fields) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    // Detect if there's a custom HostFolderField on this content type
    var hostFolderField = null;
    for (var h = 0; h < fields.length; h++) {
      var ftNorm = fields[h].fieldType.toLowerCase().replace(/-/g, '');
      if (ftNorm.indexOf('hostfolderfield') !== -1 || ftNorm.indexOf('hostfolder') !== -1) {
        hostFolderField = fields[h];
        break;
      }
    }

    var hostVar = hostFolderField ? hostFolderField.variable : 'host';

    // Only required non-body, non-system fields go in the table by default
    var requiredFields = fields.filter(function (f) {
      if (!f.required) return false;
      var ft = f.fieldType.toLowerCase().replace(/-/g, '');
      if (ft.indexOf('wysiwyg') !== -1 || ft.indexOf('storyblock') !== -1) return false;
      if (/^fields\d+$/.test(f.variable)) return false;
      if (ft.indexOf('hostfolder') !== -1) return false;
      // Skip if this is the hostFolderField (already added as system field)
      if (hostFolderField && f.variable === hostFolderField.variable) return false;
      return true;
    });

    var existing = this.findMetadataTable();

    if (existing) {
      var table = existing.table;
      var existingVars = {};
      for (var r = 1; r < table.getNumRows(); r++) {
        var varName = table.getRow(r).getCell(0).getText().trim();
        if (varName) existingVars[varName] = true;
      }

      var systemFields = ['identifier', 'contentType', hostVar, 'languageId'];
      for (var s = 0; s < systemFields.length; s++) {
        if (!existingVars[systemFields[s]]) {
          var sRow = table.appendTableRow();
          sRow.appendTableCell(systemFields[s]);
          sRow.appendTableCell('');
          existingVars[systemFields[s]] = true;
        }
      }

      for (var i = 0; i < requiredFields.length; i++) {
        var rf = requiredFields[i];
        if (!existingVars[rf.variable]) {
          var newRow = table.appendTableRow();
          newRow.appendTableCell(rf.variable);
          var valCell = newRow.appendTableCell('');
          this._applyDropdownHint(valCell, rf);
          existingVars[rf.variable] = true;
        }
      }
    } else {
      var cells = this._buildTableCells(requiredFields, hostVar);
      var table = body.insertTable(0, cells);

      var headerRow = table.getRow(0);
      headerRow.getCell(0).setText(this.MARKER_FIELD).setAttributes(this._boldStyle());
      headerRow.getCell(1).setText(this.MARKER_VALUE).setAttributes(this._boldStyle());

      // Apply dropdown hints — system fields are rows 1-5, required fields start at row 6
      var startRow = 6;
      for (var j = 0; j < requiredFields.length; j++) {
        var rowIdx = startRow + j;
        if (rowIdx < table.getNumRows()) {
          this._applyDropdownHint(table.getRow(rowIdx).getCell(1), requiredFields[j]);
        }
      }

      body.insertParagraph(1, '');
    }

    return requiredFields;
  },

  /**
   * Add a single field to the metadata table (called when user selects an optional field).
   */
  addFieldToTable: function (fieldVariable, fieldInfo) {
    var result = this.findMetadataTable();
    if (!result) return;

    var table = result.table;
    // Check if already exists
    for (var r = 1; r < table.getNumRows(); r++) {
      if (table.getRow(r).getCell(0).getText().trim() === fieldVariable) {
        return; // already in table
      }
    }

    var newRow = table.appendTableRow();
    newRow.appendTableCell(fieldVariable);
    var valCell = newRow.appendTableCell('');
    if (fieldInfo) {
      this._applyDropdownHint(valCell, fieldInfo);
    }
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
        var cell = row.getCell(1);
        cell.setText(value);
        // Reset hint styling when a real value is set
        cell.setAttributes(this._normalStyle());
        return;
      }
    }

    // Row not found — append it
    var newRow = table.appendTableRow();
    newRow.appendTableCell(fieldVariable);
    newRow.appendTableCell(value);
  },

  _normalStyle: function () {
    var style = {};
    style[DocumentApp.Attribute.ITALIC] = false;
    style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    return style;
  },

  _buildTableCells: function (editableFields, hostVar) {
    var cells = [['Field', 'Value'], ['identifier', ''], ['contentType', ''], [hostVar || 'host', ''], ['languageId', '']];
    for (var i = 0; i < editableFields.length; i++) {
      cells.push([editableFields[i].variable, '']);
    }
    return cells;
  },

  /**
   * For select and boolean fields, set the value cell text to show available options
   * as a hint. The sidebar Field Editor or manual editing can set the actual value.
   */
  _applyDropdownHint: function (cell, field) {
    var ft = field.fieldType.toLowerCase();

    if (ft.indexOf('select') !== -1 || ft.indexOf('radio') !== -1) {
      var vals = (field.values || '').split(/\r?\n/).filter(function (v) { return v.trim(); });
      if (vals.length > 0) {
        var options = vals.map(function (v) {
          var parts = v.trim().split('|');
          return (parts[1] || parts[0]).trim();
        });
        cell.setText('[' + options.join(' | ') + ']');
        cell.setAttributes(this._hintStyle());
      }
    } else if (ft.indexOf('boolean') !== -1 || ft.indexOf('checkbox') !== -1) {
      cell.setText('[true | false]');
      cell.setAttributes(this._hintStyle());
    }
  },

  _boldStyle: function () {
    var style = {};
    style[DocumentApp.Attribute.BOLD] = true;
    return style;
  },

  _hintStyle: function () {
    var style = {};
    style[DocumentApp.Attribute.ITALIC] = true;
    style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#999999';
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
    var parts = [];

    for (var i = 0; i < numChildren; i++) {
      if (i === metaTableIndex) continue;
      var h = this._elementToHtml(body.getChild(i));
      if (h) parts.push(h);
    }

    // Merge consecutive <pre><code> blocks into a single block
    var merged = [];
    var PRE_OPEN = '<pre><code>';
    var PRE_CLOSE = '</code></pre>';
    for (var p = 0; p < parts.length; p++) {
      var cur = parts[p];
      if (cur.indexOf(PRE_OPEN) === 0 && merged.length > 0 &&
          merged[merged.length - 1].indexOf(PRE_OPEN) === 0) {
        var prev = merged[merged.length - 1];
        prev = prev.substring(0, prev.length - PRE_CLOSE.length);
        var inner = cur.substring(PRE_OPEN.length, cur.length - PRE_CLOSE.length);
        merged[merged.length - 1] = prev + '\n' + inner + PRE_CLOSE;
      } else {
        merged.push(cur);
      }
    }

    return merged.join('').trim();
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

    // Code block: entire paragraph is monospace → <pre><code>
    if (heading === DocumentApp.ParagraphHeading.NORMAL && this._isCodeParagraph(para)) {
      var plainText = para.getText();
      if (!plainText) return '';
      return '<pre><code>' + this._escapeHtml(plainText) + '</code></pre>';
    }

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
      var fontFamily = textEl.getFontFamily(i);

      var j = i + 1;
      while (j < text.length &&
             textEl.isBold(j) === bold &&
             textEl.isItalic(j) === italic &&
             textEl.isUnderline(j) === underline &&
             textEl.isStrikethrough(j) === strikethrough &&
             textEl.getLinkUrl(j) === link &&
             textEl.getFontFamily(j) === fontFamily) {
        j++;
      }

      var chunk = this._escapeHtml(text.substring(i, j));

      if (link) chunk = '<a href="' + this._escapeHtml(link) + '">' + chunk + '</a>';
      if (bold) chunk = '<strong>' + chunk + '</strong>';
      if (italic) chunk = '<em>' + chunk + '</em>';
      if (underline) chunk = '<u>' + chunk + '</u>';
      if (strikethrough) chunk = '<s>' + chunk + '</s>';
      if (this._isMonospace(fontFamily)) chunk = '<code>' + chunk + '</code>';

      html += chunk;
      i = j;
    }
    return html;
  },

  // ── Code detection helpers ──

  _MONOSPACE_FONTS: ['courier new', 'courier', 'consolas', 'source code pro',
    'roboto mono', 'fira code', 'fira mono', 'inconsolata', 'menlo', 'monaco',
    'ubuntu mono', 'droid sans mono', 'monospace', 'noto sans mono'],

  _isMonospace: function (fontFamily) {
    if (!fontFamily) return false;
    var lower = fontFamily.toLowerCase();
    for (var i = 0; i < this._MONOSPACE_FONTS.length; i++) {
      if (lower.indexOf(this._MONOSPACE_FONTS[i]) !== -1) return true;
    }
    return false;
  },

  /**
   * Check if an entire paragraph consists only of monospace-font text.
   */
  _isCodeParagraph: function (para) {
    var numChildren = para.getNumChildren();
    if (numChildren === 0) return false;
    var hasText = false;
    for (var i = 0; i < numChildren; i++) {
      var child = para.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.TEXT) return false;
      var textEl = child.asText();
      var text = textEl.getText();
      if (!text) continue;
      hasText = true;
      for (var j = 0; j < text.length; j++) {
        if (!this._isMonospace(textEl.getFontFamily(j))) return false;
      }
    }
    return hasText;
  },

  _escapeHtml: function (str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  },

  /**
   * Extract inline images from metadata table cells for image/file fields.
   * @param {string[]} imageFieldVars - list of field variables that are image/file type
   * @returns {Array<{variable: string, blob: Blob, name: string, blobHash: string, rowIndex: number}>}
   */
  getMetadataImageFields: function (imageFieldVars) {
    var result = this.findMetadataTable();
    if (!result) return [];

    var table = result.table;
    var images = [];
    var varSet = {};
    for (var v = 0; v < imageFieldVars.length; v++) {
      varSet[imageFieldVars[v]] = true;
    }

    for (var r = 1; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      var fieldVar = row.getCell(0).getText().trim();
      if (!varSet[fieldVar]) continue;

      var valueCell = row.getCell(1);
      // Get any existing text in the cell (could be an identifier from a previous upload)
      var cellText = valueCell.getText().trim();
      var hasImage = false;

      // Look for inline images in the value cell
      var numChildren = valueCell.getNumChildren();
      for (var c = 0; c < numChildren; c++) {
        var child = valueCell.getChild(c);
        // Child could be a Paragraph containing an InlineImage
        if (typeof child.getNumChildren === 'function') {
          for (var gc = 0; gc < child.getNumChildren(); gc++) {
            var grandChild = child.getChild(gc);
            if (grandChild.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
              hasImage = true;
              var blob = grandChild.getBlob();
              var hash = this._hashBlob(blob);
              var ext = this._getExtension(blob.getContentType());
              images.push({
                variable: fieldVar,
                blob: blob,
                name: fieldVar + '.' + ext,
                blobHash: hash,
                rowIndex: r,
                existingId: cellText  // identifier text already in the cell
              });
            }
          }
        }
        if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          hasImage = true;
          var blob = child.getBlob();
          var hash = this._hashBlob(blob);
          var ext = this._getExtension(blob.getContentType());
          images.push({
            variable: fieldVar,
            blob: blob,
            name: fieldVar + '.' + ext,
            blobHash: hash,
            rowIndex: r,
            existingId: cellText
          });
        }
      }
    }
    return images;
  },

  /**
   * Set the identifier text below an image in a metadata table cell,
   * preserving the inline image.
   */
  setImageFieldIdentifier: function (fieldVariable, identifier) {
    var result = this.findMetadataTable();
    if (!result) return;

    var table = result.table;
    for (var r = 1; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      if (row.getCell(0).getText().trim() !== fieldVariable) continue;

      var valueCell = row.getCell(1);
      // Find the paragraph that contains the inline image
      var numChildren = valueCell.getNumChildren();
      var foundImage = false;

      for (var c = 0; c < numChildren; c++) {
        var child = valueCell.getChild(c);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var para = child.asParagraph();
          for (var gc = 0; gc < para.getNumChildren(); gc++) {
            if (para.getChild(gc).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
              foundImage = true;
              break;
            }
          }
        }
      }

      if (foundImage) {
        // Remove any existing text paragraphs (old identifier), keep image paragraphs
        for (var c = numChildren - 1; c >= 0; c--) {
          var child = valueCell.getChild(c);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            var hasImg = false;
            var para = child.asParagraph();
            for (var gc = 0; gc < para.getNumChildren(); gc++) {
              if (para.getChild(gc).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
                hasImg = true;
                break;
              }
            }
            if (!hasImg && c > 0) {
              // Remove text-only paragraph (but not if it's the only child)
              valueCell.removeChild(child);
            }
          }
        }
        // Append a new paragraph with the identifier
        valueCell.appendParagraph(identifier).setAttributes(this._normalStyle());
      } else {
        // No image — just set text
        valueCell.setText(identifier);
      }
      return;
    }
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
