/**
 * Cleans Google Docs exported HTML by stripping inline styles,
 * unnecessary markup, and Google-specific artifacts.
 */
var HtmlCleaner = {

  clean: function (html) {
    // Remove everything before <body> and after </body>
    var bodyMatch = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    if (bodyMatch) {
      html = bodyMatch[1];
    }

    // Remove style attributes
    html = html.replace(/\s+style="[^"]*"/gi, '');

    // Remove class attributes
    html = html.replace(/\s+class="[^"]*"/gi, '');

    // Remove id attributes
    html = html.replace(/\s+id="[^"]*"/gi, '');

    // Remove Google's <span> wrappers that have no attributes left
    html = html.replace(/<span>([\s\S]*?)<\/span>/gi, '$1');

    // Remove empty paragraphs
    html = html.replace(/<p>\s*<\/p>/gi, '');

    // Remove Google's font and color spans
    html = html.replace(/<span[^>]*font-family[^>]*>([\s\S]*?)<\/span>/gi, '$1');

    // Remove <a> tags that are internal Google Doc anchors (id= links)
    html = html.replace(/<a\s+id="[^"]*"\s*>\s*<\/a>/gi, '');

    // Collapse multiple <br> tags
    html = html.replace(/(<br\s*\/?>){3,}/gi, '<br><br>');

    // Remove Google's comments/suggestions markup
    html = html.replace(/<!--[\s\S]*?-->/g, '');

    // Remove empty divs
    html = html.replace(/<div>\s*<\/div>/gi, '');

    // Trim whitespace
    html = html.trim();

    return html;
  }
};
