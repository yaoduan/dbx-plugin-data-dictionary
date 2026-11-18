/**
 * WordprocessingML (.docx) writer.
 *
 * Builds a document from a small block list (heading / paragraph / table /
 * pageBreak) so the report layer never touches raw OOXML. Fonts are declared
 * with an explicit `eastAsia` face, otherwise Word falls back to a Latin font
 * and CJK comments render in a mismatched typeface.
 */
(function (global) {
  "use strict";

  var LATIN_FONT = "Calibri";
  var EAST_ASIA_FONT = "Microsoft YaHei";

  function escapeXml(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function sanitize(value) {
    return String(value === null || value === undefined ? "" : value).replace(
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g,
      ""
    );
  }

  /** Split on newlines so multi-line comments keep their breaks in Word. */
  function runsFor(text, options) {
    var clean = sanitize(text);
    var style = "";
    if (options && options.bold) style += "<w:b/>";
    if (options && options.italic) style += "<w:i/>";
    if (options && options.color) style += '<w:color w:val="' + options.color + '"/>';
    if (options && options.size) style += '<w:sz w:val="' + options.size * 2 + '"/>';
    var fonts =
      '<w:rFonts w:ascii="' +
      LATIN_FONT +
      '" w:hAnsi="' +
      LATIN_FONT +
      '" w:eastAsia="' +
      EAST_ASIA_FONT +
      '" w:cs="' +
      LATIN_FONT +
      '"/>';
    var properties = "<w:rPr>" + fonts + style + "</w:rPr>";
    var lines = clean.split(/\r\n|\r|\n/);
    return lines
      .map(function (line, index) {
        var br = index > 0 ? "<w:br/>" : "";
        return "<w:r>" + properties + br + '<w:t xml:space="preserve">' + escapeXml(line) + "</w:t></w:r>";
      })
      .join("");
  }

  function paragraph(text, options) {
    var settings = options || {};
    var properties = [];
    if (settings.style) properties.push('<w:pStyle w:val="' + settings.style + '"/>');
    if (settings.align) properties.push('<w:jc w:val="' + settings.align + '"/>');
    var spacing = [];
    if (settings.spaceBefore !== undefined) spacing.push('w:before="' + settings.spaceBefore + '"');
    if (settings.spaceAfter !== undefined) spacing.push('w:after="' + settings.spaceAfter + '"');
    if (spacing.length) properties.push("<w:spacing " + spacing.join(" ") + "/>");
    var paragraphProperties = properties.length ? "<w:pPr>" + properties.join("") + "</w:pPr>" : "";
    return "<w:p>" + paragraphProperties + runsFor(text, settings) + "</w:p>";
  }

  function cellXml(cell, widthPct, isHeader) {
    var value = cell && typeof cell === "object" ? cell : { text: cell };
    var shading = isHeader
      ? '<w:shd w:val="clear" w:color="auto" w:fill="4F6BED"/>'
      : value.fill
        ? '<w:shd w:val="clear" w:color="auto" w:fill="' + value.fill + '"/>'
        : "";
    var span = value.span && value.span > 1 ? '<w:gridSpan w:val="' + value.span + '"/>' : "";
    var properties =
      "<w:tcPr>" +
      '<w:tcW w:w="' + widthPct + '" w:type="pct"/>' +
      span +
      shading +
      '<w:vAlign w:val="center"/>' +
      "</w:tcPr>";
    var text = paragraph(value.text, {
      bold: isHeader || value.bold,
      color: isHeader ? "FFFFFF" : value.color,
      spaceBefore: 20,
      spaceAfter: 20,
      size: value.size,
    });
    return "<w:tc>" + properties + text + "</w:tc>";
  }

  /**
   * `table` is `{ header: [string], rows: [[cell]], widths: [number] }` where
   * widths are percentages of the text column (they are normalised to 5000 =
   * 100%, the unit OOXML uses for `w:type="pct"`).
   */
  function tableXml(table) {
    var header = table.header || [];
    var rows = table.rows || [];
    var columnCount = header.length || (rows[0] ? rows[0].length : 1);
    var weights = table.widths && table.widths.length === columnCount ? table.widths : null;
    var total = weights
      ? weights.reduce(function (sum, value) {
          return sum + value;
        }, 0)
      : columnCount;
    var widths = [];
    for (var i = 0; i < columnCount; i += 1) {
      var weight = weights ? weights[i] : 1;
      widths.push(Math.max(1, Math.round((weight / total) * 5000)));
    }

    var borders =
      "<w:tblBorders>" +
      ["top", "left", "bottom", "right", "insideH", "insideV"]
        .map(function (side) {
          return '<w:' + side + ' w:val="single" w:sz="4" w:space="0" w:color="D0D7E5"/>';
        })
        .join("") +
      "</w:tblBorders>";
    var properties =
      "<w:tblPr>" +
      '<w:tblW w:w="5000" w:type="pct"/>' +
      '<w:tblLayout w:type="fixed"/>' +
      borders +
      '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>' +
      '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>' +
      "</w:tblPr>";
    var grid =
      "<w:tblGrid>" +
      widths
        .map(function (width) {
          return '<w:gridCol w:w="' + Math.round((width / 5000) * 9360) + '"/>';
        })
        .join("") +
      "</w:tblGrid>";

    var parts = ["<w:tbl>", properties, grid];
    if (header.length) {
      // tblHeader repeats the header row when a table spans pages.
      parts.push('<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>');
      header.forEach(function (cell, index) {
        parts.push(cellXml(cell, widths[index], true));
      });
      parts.push("</w:tr>");
    }
    rows.forEach(function (row) {
      parts.push("<w:tr>");
      var offset = 0;
      row.forEach(function (cell) {
        var value = cell && typeof cell === "object" ? cell : { text: cell };
        var span = value.span && value.span > 1 ? value.span : 1;
        var width = 0;
        for (var i = 0; i < span && offset + i < widths.length; i += 1) width += widths[offset + i];
        parts.push(cellXml(cell, width || widths[Math.min(offset, widths.length - 1)], false));
        offset += span;
      });
      parts.push("</w:tr>");
    });
    parts.push("</w:tbl>");
    // Word needs a paragraph between consecutive tables, or they merge.
    parts.push(paragraph("", { spaceBefore: 0, spaceAfter: 60 }));
    return parts.join("");
  }

  function blockXml(block) {
    if (!block) return "";
    if (block.type === "heading") {
      var level = Math.min(Math.max(block.level || 1, 1), 3);
      return paragraph(block.text, {
        style: "Heading" + level,
        spaceBefore: level === 1 ? 240 : 200,
        spaceAfter: 120,
      });
    }
    if (block.type === "paragraph") {
      return paragraph(block.text, {
        bold: block.bold,
        italic: block.italic,
        color: block.color,
        align: block.align,
        size: block.size,
        spaceBefore: block.spaceBefore === undefined ? 0 : block.spaceBefore,
        spaceAfter: block.spaceAfter === undefined ? 80 : block.spaceAfter,
      });
    }
    if (block.type === "table") return tableXml(block);
    if (block.type === "pageBreak") {
      return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    }
    return "";
  }

  function stylesXml() {
    var fonts =
      '<w:rFonts w:ascii="' +
      LATIN_FONT +
      '" w:hAnsi="' +
      LATIN_FONT +
      '" w:eastAsia="' +
      EAST_ASIA_FONT +
      '" w:cs="' +
      LATIN_FONT +
      '"/>';
    function heading(id, name, size, color, outline) {
      return (
        '<w:style w:type="paragraph" w:styleId="' +
        id +
        '"><w:name w:val="' +
        name +
        '"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
        '<w:pPr><w:keepNext/><w:outlineLvl w:val="' +
        outline +
        '"/><w:spacing w:before="240" w:after="120"/></w:pPr>' +
        "<w:rPr>" +
        fonts +
        '<w:b/><w:color w:val="' +
        color +
        '"/><w:sz w:val="' +
        size * 2 +
        '"/></w:rPr></w:style>'
      );
    }
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:docDefaults><w:rPrDefault><w:rPr>" +
      fonts +
      '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
      "</w:docDefaults>" +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      heading("Heading1", "heading 1", 18, "1F3864", 0) +
      heading("Heading2", "heading 2", 14, "2E5496", 1) +
      heading("Heading3", "heading 3", 12, "44546A", 2) +
      "</w:styles>"
    );
  }

  function documentXml(blocks) {
    var body = blocks
      .map(function (block) {
        return blockXml(block);
      })
      .join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      "<w:body>" +
      body +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>' +
      "</w:sectPr></w:body></w:document>"
    );
  }

  function contentTypesXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      "</Types>"
    );
  }

  function rootRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      "</Relationships>"
    );
  }

  function documentRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      "</Relationships>"
    );
  }

  function corePropsXml(title, generatedAt) {
    var iso = generatedAt.toISOString().replace(/\.\d{3}Z$/, "Z");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
      ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
      ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      "<dc:title>" +
      escapeXml(sanitize(title)) +
      "</dc:title>" +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' +
      iso +
      "</dcterms:created>" +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' +
      iso +
      "</dcterms:modified></cp:coreProperties>"
    );
  }

  function appPropsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
      ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      "<Application>DBX Data Dictionary Export</Application></Properties>"
    );
  }

  /** Build the .docx bytes from `{ title, generatedAt, blocks }`. */
  async function createDocument(options) {
    var entries = [
      { name: "[Content_Types].xml", content: contentTypesXml() },
      { name: "_rels/.rels", content: rootRelsXml() },
      { name: "docProps/core.xml", content: corePropsXml(options.title || "Data dictionary", options.generatedAt || new Date()) },
      { name: "docProps/app.xml", content: appPropsXml() },
      { name: "word/document.xml", content: documentXml(options.blocks || []) },
      { name: "word/_rels/document.xml.rels", content: documentRelsXml() },
      { name: "word/styles.xml", content: stylesXml() },
    ];
    return global.DictZip.createZip(entries);
  }

  global.DictDocx = {
    createDocument: createDocument,
    escapeXml: escapeXml,
    sanitize: sanitize,
  };
})(window);
