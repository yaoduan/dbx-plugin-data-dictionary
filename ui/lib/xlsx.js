/**
 * SpreadsheetML (.xlsx) writer.
 *
 * Writes the minimal part set Excel/WPS/LibreOffice need: content types,
 * package and workbook relationships, workbook, per-sheet worksheets and a
 * small style table. Strings are written inline (`t="inlineStr"`) so there is
 * no shared-string table to keep in sync.
 */
(function (global) {
  "use strict";

  function escapeXml(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  /**
   * XML 1.0 forbids most control characters outright — a stray 0x01 in a column
   * comment would make Excel refuse the whole file, so they are dropped.
   */
  function sanitize(value) {
    return String(value === null || value === undefined ? "" : value).replace(
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g,
      ""
    );
  }

  function columnLetter(index) {
    var letters = "";
    var remaining = index;
    while (remaining > 0) {
      var rest = (remaining - 1) % 26;
      letters = String.fromCharCode(65 + rest) + letters;
      remaining = Math.floor((remaining - rest - 1) / 26);
    }
    return letters || "A";
  }

  /** Excel sheet names: 31 chars, no []:*?/\ and no leading/trailing quote. */
  function sheetName(name, used) {
    var cleaned = sanitize(name).replace(/[\[\]:*?\/\\]/g, "_").trim();
    if (!cleaned) cleaned = "Sheet";
    if (cleaned.length > 31) cleaned = cleaned.slice(0, 31);
    var candidate = cleaned;
    var suffix = 2;
    while (used[candidate.toLowerCase()]) {
      var tail = "~" + suffix;
      candidate = cleaned.slice(0, Math.max(1, 31 - tail.length)) + tail;
      suffix += 1;
    }
    used[candidate.toLowerCase()] = true;
    return candidate;
  }

  function cellXml(reference, value, styleIndex) {
    var style = styleIndex ? ' s="' + styleIndex + '"' : "";
    if (value === null || value === undefined || value === "") {
      return '<c r="' + reference + '"' + style + "/>";
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      return '<c r="' + reference + '"' + style + "><v>" + value + "</v></c>";
    }
    if (typeof value === "boolean") {
      return '<c r="' + reference + '" t="b"' + style + "><v>" + (value ? 1 : 0) + "</v></c>";
    }
    return (
      '<c r="' +
      reference +
      '" t="inlineStr"' +
      style +
      '><is><t xml:space="preserve">' +
      escapeXml(sanitize(value)) +
      "</t></is></c>"
    );
  }

  /**
   * One worksheet.
   * `sheet` is `{ name, columns: [{ header, width }], rows: [[value]],
   * rowStyles?: [styleIndex], freezeHeader? }`.
   */
  function worksheetXml(sheet) {
    var columns = sheet.columns || [];
    var parts = [];
    parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    parts.push(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    );
    if (columns.length) {
      parts.push("<cols>");
      columns.forEach(function (column, index) {
        var width = column && column.width ? column.width : 18;
        parts.push('<col min="' + (index + 1) + '" max="' + (index + 1) + '" width="' + width + '" customWidth="1"/>');
      });
      parts.push("</cols>");
    }
    if (sheet.freezeHeader) {
      parts.push('<sheetViews><sheetView workbookViewId="0">');
      parts.push('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>');
      parts.push("</sheetView></sheetViews>");
    }
    parts.push("<sheetData>");

    var rowNumber = 1;
    if (columns.length) {
      var headerCells = columns
        .map(function (column, index) {
          return cellXml(columnLetter(index + 1) + rowNumber, column.header, 1);
        })
        .join("");
      parts.push('<row r="' + rowNumber + '" customHeight="1" ht="20">' + headerCells + "</row>");
      rowNumber += 1;
    }

    (sheet.rows || []).forEach(function (row, rowIndex) {
      var style = sheet.rowStyles && sheet.rowStyles[rowIndex] ? sheet.rowStyles[rowIndex] : 2;
      var cells = row
        .map(function (value, index) {
          return cellXml(columnLetter(index + 1) + rowNumber, value, style);
        })
        .join("");
      parts.push('<row r="' + rowNumber + '">' + cells + "</row>");
      rowNumber += 1;
    });

    parts.push("</sheetData>");
    if (columns.length && (sheet.rows || []).length) {
      parts.push(
        '<autoFilter ref="A1:' + columnLetter(columns.length) + (sheet.rows.length + 1) + '"/>'
      );
    }
    parts.push("</worksheet>");
    return parts.join("");
  }

  function stylesXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="3">' +
      '<font><sz val="10"/><name val="Microsoft YaHei"/></font>' +
      '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font>' +
      '<font><b/><sz val="12"/><name val="Microsoft YaHei"/></font>' +
      "</fonts>" +
      '<fills count="3">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF4F6BED"/><bgColor indexed="64"/></patternFill></fill>' +
      "</fills>" +
      '<borders count="2">' +
      "<border><left/><right/><top/><bottom/><diagonal/></border>" +
      '<border><left style="thin"><color rgb="FFD0D7E5"/></left><right style="thin"><color rgb="FFD0D7E5"/></right>' +
      '<top style="thin"><color rgb="FFD0D7E5"/></top><bottom style="thin"><color rgb="FFD0D7E5"/></bottom><diagonal/></border>' +
      "</borders>" +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="4">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1">' +
      '<alignment vertical="center" horizontal="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1">' +
      '<alignment vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      "</styleSheet>"
    );
  }

  function workbookXml(sheets) {
    var entries = sheets
      .map(function (sheet, index) {
        return (
          '<sheet name="' + escapeXml(sheet.resolvedName) + '" sheetId="' + (index + 1) + '" r:id="rId' + (index + 1) + '"/>'
        );
      })
      .join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      "<sheets>" +
      entries +
      "</sheets></workbook>"
    );
  }

  function workbookRelsXml(sheets) {
    var entries = sheets
      .map(function (sheet, index) {
        return (
          '<Relationship Id="rId' +
          (index + 1) +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"' +
          ' Target="worksheets/sheet' +
          (index + 1) +
          '.xml"/>'
        );
      })
      .join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      entries +
      '<Relationship Id="rId' +
      (sheets.length + 1) +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      "</Relationships>"
    );
  }

  function contentTypesXml(sheets) {
    var overrides = sheets
      .map(function (sheet, index) {
        return (
          '<Override PartName="/xl/worksheets/sheet' +
          (index + 1) +
          '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        );
      })
      .join("");
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      overrides +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      "</Types>"
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
      "</dcterms:modified>" +
      "</cp:coreProperties>"
    );
  }

  function appPropsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
      ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      "<Application>DBX Data Dictionary Export</Application>" +
      "</Properties>"
    );
  }

  function rootRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      "</Relationships>"
    );
  }

  /** Build the .xlsx bytes from `{ title, generatedAt, sheets }`. */
  async function createWorkbook(options) {
    var used = Object.create(null);
    var sheets = (options.sheets || []).map(function (sheet) {
      var copy = Object.assign({}, sheet);
      copy.resolvedName = sheetName(sheet.name, used);
      return copy;
    });
    if (!sheets.length) throw new Error("Workbook needs at least one sheet");

    var entries = [
      { name: "[Content_Types].xml", content: contentTypesXml(sheets) },
      { name: "_rels/.rels", content: rootRelsXml() },
      { name: "docProps/core.xml", content: corePropsXml(options.title || "Data dictionary", options.generatedAt || new Date()) },
      { name: "docProps/app.xml", content: appPropsXml() },
      { name: "xl/workbook.xml", content: workbookXml(sheets) },
      { name: "xl/_rels/workbook.xml.rels", content: workbookRelsXml(sheets) },
      { name: "xl/styles.xml", content: stylesXml() },
    ];
    sheets.forEach(function (sheet, index) {
      entries.push({ name: "xl/worksheets/sheet" + (index + 1) + ".xml", content: worksheetXml(sheet) });
    });
    return global.DictZip.createZip(entries);
  }

  global.DictXlsx = {
    createWorkbook: createWorkbook,
    sheetName: sheetName,
    columnLetter: columnLetter,
    escapeXml: escapeXml,
    sanitize: sanitize,
  };
})(window);
