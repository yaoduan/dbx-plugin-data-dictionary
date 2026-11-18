/**
 * Turns the collected schema into document blocks (docx) and sheets (xlsx).
 * Both formats carry the same information so the two exports stay comparable.
 */
(function (global) {
  "use strict";

  var COLUMN_HEADERS = ["#", "字段名", "类型", "长度/精度", "可空", "主键", "默认值", "注释"];
  var COLUMN_WIDTHS = [4, 17, 14, 9, 6, 6, 14, 30];

  function pad(value) {
    return value < 10 ? "0" + value : String(value);
  }

  function formatTimestamp(date) {
    return (
      date.getFullYear() +
      "-" +
      pad(date.getMonth() + 1) +
      "-" +
      pad(date.getDate()) +
      " " +
      pad(date.getHours()) +
      ":" +
      pad(date.getMinutes()) +
      ":" +
      pad(date.getSeconds())
    );
  }

  function formatTimestampForFile(date) {
    return (
      date.getFullYear() +
      pad(date.getMonth() + 1) +
      pad(date.getDate()) +
      "-" +
      pad(date.getHours()) +
      pad(date.getMinutes()) +
      pad(date.getSeconds())
    );
  }

  function scopeLabel(schema) {
    var parts = [];
    if (schema.database) parts.push(schema.database);
    if (schema.schema && schema.schema !== schema.database) parts.push(schema.schema);
    return parts.join(" / ") || schema.connectionName || "database";
  }

  function documentTitle(schema) {
    return scopeLabel(schema) + " 数据字典";
  }

  function fileBaseName(schema) {
    var scope = (schema.database || schema.schema || schema.connectionName || "database")
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "-");
    return scope + "-data-dictionary-" + formatTimestampForFile(schema.generatedAt);
  }

  function indexLabel(index) {
    var kind = index.primary ? "PRIMARY" : index.unique ? "UNIQUE" : "INDEX";
    return kind + (index.type ? " / " + index.type : "");
  }

  function foreignKeyTarget(fk) {
    var target = fk.refTable || "";
    if (fk.refSchema && fk.refSchema !== "") target = fk.refSchema + "." + target;
    var columns = fk.refColumns && fk.refColumns.length ? "(" + fk.refColumns.join(", ") + ")" : "";
    return target + columns;
  }

  function columnRow(column, position) {
    return [
      String(position),
      column.name,
      column.columnType || column.dataType,
      column.length,
      column.nullable ? "是" : "否",
      column.primaryKey ? "是" : "",
      column.defaultValue,
      column.comment,
    ];
  }

  function tableHeading(table) {
    return table.comment ? table.name + "（" + table.comment + "）" : table.name;
  }

  /* ---------------------------------------------------------------- docx */

  function buildDocxBlocks(schema, options) {
    var blocks = [];
    blocks.push({ type: "heading", level: 1, text: documentTitle(schema) });

    var meta = [
      ["连接", schema.connectionName || ""],
      ["数据库类型", schema.dbType || ""],
      ["数据库", schema.database || "—"],
      ["Schema", schema.schema || "—"],
      ["对象数量", String(schema.tables.length)],
      ["生成时间", formatTimestamp(schema.generatedAt)],
    ];
    blocks.push({
      type: "table",
      header: ["项目", "值"],
      widths: [22, 78],
      rows: meta.map(function (entry) {
        return [{ text: entry[0], bold: true }, entry[1]];
      }),
    });

    if (schema.warnings && schema.warnings.length) {
      blocks.push({ type: "heading", level: 2, text: "提示" });
      schema.warnings.forEach(function (warning) {
        blocks.push({ type: "paragraph", text: "• " + warning, color: "9A6700" });
      });
    }

    blocks.push({ type: "heading", level: 2, text: "表清单" });
    blocks.push({
      type: "table",
      header: ["#", "名称", "类型", "字段数", "行数(估)", "注释"],
      widths: [5, 26, 11, 8, 10, 40],
      rows: schema.tables.map(function (table, index) {
        return [
          String(index + 1),
          table.name,
          table.type,
          String(table.columns.length),
          table.rowCount === null || table.rowCount === undefined ? "" : String(table.rowCount),
          table.comment,
        ];
      }),
    });

    schema.tables.forEach(function (table) {
      blocks.push({ type: "pageBreak" });
      blocks.push({ type: "heading", level: 2, text: tableHeading(table) });

      var facts = [];
      facts.push("类型：" + table.type);
      if (table.engine) facts.push("引擎：" + table.engine);
      if (table.collation) facts.push("字符集：" + table.collation);
      if (table.rowCount !== null && table.rowCount !== undefined) facts.push("行数(估)：" + table.rowCount);
      facts.push("字段数：" + table.columns.length);
      blocks.push({ type: "paragraph", text: facts.join("　|　"), color: "595959" });

      if (table.columns.length) {
        blocks.push({
          type: "table",
          header: COLUMN_HEADERS,
          widths: COLUMN_WIDTHS,
          rows: table.columns.map(function (column, index) {
            return columnRow(column, index + 1);
          }),
        });
      } else {
        blocks.push({ type: "paragraph", text: "（没有读到字段）", italic: true, color: "9A6700" });
      }

      if (options.includeIndexes && table.indexes.length) {
        blocks.push({ type: "heading", level: 3, text: "索引" });
        blocks.push({
          type: "table",
          header: ["索引名", "类型", "字段"],
          widths: [30, 22, 48],
          rows: table.indexes.map(function (index) {
            return [index.name, indexLabel(index), index.columns.join(", ")];
          }),
        });
      }

      if (options.includeForeignKeys && table.foreignKeys.length) {
        blocks.push({ type: "heading", level: 3, text: "外键" });
        blocks.push({
          type: "table",
          header: ["约束名", "本表字段", "引用对象"],
          widths: [30, 28, 42],
          rows: table.foreignKeys.map(function (fk) {
            return [fk.name, fk.columns.join(", "), foreignKeyTarget(fk)];
          }),
        });
      }
    });

    return blocks;
  }

  /* ---------------------------------------------------------------- xlsx */

  function overviewSheet(schema, options) {
    var rows = [];
    rows.push(["数据字典", documentTitle(schema)]);
    rows.push(["连接", schema.connectionName || ""]);
    rows.push(["数据库类型", schema.dbType || ""]);
    rows.push(["数据库", schema.database || ""]);
    rows.push(["Schema", schema.schema || ""]);
    rows.push(["生成时间", formatTimestamp(schema.generatedAt)]);
    rows.push(["对象数量", schema.tables.length]);
    rows.push([]);
    (schema.warnings || []).forEach(function (warning) {
      rows.push(["提示", warning]);
    });
    if (schema.warnings && schema.warnings.length) rows.push([]);
    rows.push(["#", "名称", "类型", "字段数", "行数(估)", "注释"]);
    schema.tables.forEach(function (table, index) {
      rows.push([
        index + 1,
        table.name,
        table.type,
        table.columns.length,
        table.rowCount === null || table.rowCount === undefined ? "" : table.rowCount,
        table.comment,
      ]);
    });
    void options;
    return {
      name: "总览",
      columns: [
        { header: "项目", width: 10 },
        { header: "名称", width: 32 },
        { header: "类型", width: 14 },
        { header: "字段数", width: 9 },
        { header: "行数(估)", width: 12 },
        { header: "注释", width: 46 },
      ],
      rows: rows,
      freezeHeader: true,
    };
  }

  function tableSheetColumns() {
    return [
      { header: "#", width: 5 },
      { header: "字段名", width: 26 },
      { header: "类型", width: 20 },
      { header: "长度/精度", width: 12 },
      { header: "可空", width: 7 },
      { header: "主键", width: 7 },
      { header: "默认值", width: 20 },
      { header: "注释", width: 40 },
    ];
  }

  function tableSheetRows(table, options) {
    var rows = table.columns.map(function (column, index) {
      return columnRow(column, index + 1);
    });
    if (options.includeIndexes && table.indexes.length) {
      rows.push([]);
      rows.push(["索引", "索引名", "类型", "字段", "", "", "", ""]);
      table.indexes.forEach(function (index) {
        rows.push(["", index.name, indexLabel(index), index.columns.join(", "), "", "", "", ""]);
      });
    }
    if (options.includeForeignKeys && table.foreignKeys.length) {
      rows.push([]);
      rows.push(["外键", "约束名", "本表字段", "引用对象", "", "", "", ""]);
      table.foreignKeys.forEach(function (fk) {
        rows.push(["", fk.name, fk.columns.join(", "), foreignKeyTarget(fk), "", "", "", ""]);
      });
    }
    return rows;
  }

  /** One flat sheet with every column of every table. */
  function flatSheet(schema, options) {
    var rows = [];
    schema.tables.forEach(function (table) {
      table.columns.forEach(function (column, index) {
        rows.push([
          table.name,
          table.comment,
          index + 1,
          column.name,
          column.columnType || column.dataType,
          column.length,
          column.nullable ? "是" : "否",
          column.primaryKey ? "是" : "",
          column.defaultValue,
          column.comment,
        ]);
      });
    });
    void options;
    return {
      name: "字段明细",
      columns: [
        { header: "表名", width: 26 },
        { header: "表注释", width: 26 },
        { header: "#", width: 5 },
        { header: "字段名", width: 26 },
        { header: "类型", width: 20 },
        { header: "长度/精度", width: 12 },
        { header: "可空", width: 7 },
        { header: "主键", width: 7 },
        { header: "默认值", width: 20 },
        { header: "注释", width: 40 },
      ],
      rows: rows,
      freezeHeader: true,
    };
  }

  function indexSheet(schema) {
    var rows = [];
    schema.tables.forEach(function (table) {
      table.indexes.forEach(function (index) {
        rows.push([table.name, index.name, indexLabel(index), index.columns.join(", ")]);
      });
    });
    return {
      name: "索引",
      columns: [
        { header: "表名", width: 26 },
        { header: "索引名", width: 30 },
        { header: "类型", width: 20 },
        { header: "字段", width: 46 },
      ],
      rows: rows,
      freezeHeader: true,
    };
  }

  function foreignKeySheet(schema) {
    var rows = [];
    schema.tables.forEach(function (table) {
      table.foreignKeys.forEach(function (fk) {
        rows.push([table.name, fk.name, fk.columns.join(", "), foreignKeyTarget(fk)]);
      });
    });
    return {
      name: "外键",
      columns: [
        { header: "表名", width: 26 },
        { header: "约束名", width: 30 },
        { header: "本表字段", width: 28 },
        { header: "引用对象", width: 40 },
      ],
      rows: rows,
      freezeHeader: true,
    };
  }

  function buildSheets(schema, options) {
    var sheets = [overviewSheet(schema, options)];
    if (options.excelLayout === "sheet-per-table") {
      schema.tables.forEach(function (table) {
        sheets.push({
          name: table.name,
          columns: tableSheetColumns(),
          rows: tableSheetRows(table, options),
          freezeHeader: true,
        });
      });
    } else {
      sheets.push(flatSheet(schema, options));
      if (options.includeIndexes) sheets.push(indexSheet(schema));
      if (options.includeForeignKeys) sheets.push(foreignKeySheet(schema));
    }
    return sheets;
  }

  async function buildDocx(schema, options) {
    return global.DictDocx.createDocument({
      title: documentTitle(schema),
      generatedAt: schema.generatedAt,
      blocks: buildDocxBlocks(schema, options),
    });
  }

  async function buildXlsx(schema, options) {
    return global.DictXlsx.createWorkbook({
      title: documentTitle(schema),
      generatedAt: schema.generatedAt,
      sheets: buildSheets(schema, options),
    });
  }

  global.DictReport = {
    buildDocx: buildDocx,
    buildXlsx: buildXlsx,
    buildDocxBlocks: buildDocxBlocks,
    buildSheets: buildSheets,
    fileBaseName: fileBaseName,
    documentTitle: documentTitle,
    formatTimestamp: formatTimestamp,
  };
})(window);
