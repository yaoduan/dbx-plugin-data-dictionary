/**
 * Metadata collection on top of `host.queryData`.
 *
 * The host caps one query at 5000 rows (default 500) and 8 MiB, so every
 * catalog read is paged until a short page arrives. Rows come back as arrays,
 * so each page is mapped onto the column names the dialect profile aliased.
 */
(function (global) {
  "use strict";

  var PAGE_ROWS = 2000;
  var MAX_PAGES = 400; // 800k catalog rows: a hard stop, not an expected limit.
  var QUERY_TIMEOUT_MS = 60000;

  function rowsToObjects(result) {
    var columns = (result && result.columns) || [];
    var rows = (result && result.rows) || [];
    var names = columns.map(function (column) {
      return String((column && column.name) || "").toLowerCase();
    });
    return rows.map(function (row) {
      var object = {};
      for (var i = 0; i < names.length; i += 1) object[names[i]] = row[i];
      return object;
    });
  }

  function text(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "string") return value;
    return String(value);
  }

  function isYes(value) {
    var normalized = text(value).trim().toUpperCase();
    return normalized === "YES" || normalized === "Y" || normalized === "1" || normalized === "TRUE";
  }

  function number(value) {
    if (value === null || value === undefined || value === "") return null;
    var parsed = typeof value === "number" ? value : Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }

  /**
   * Run one paged catalog query. `baseSql` must already carry a deterministic
   * ORDER BY, otherwise paging could repeat or skip rows.
   *
   * The host caps a result by rows *and* by serialized bytes (8 MiB), so a page
   * can come back short with `truncated: true` -- wide comment-heavy catalogs
   * hit that first. Paging therefore advances by the rows actually received and
   * only stops on a short page that the host did not truncate.
   */
  async function queryPaged(ctx, baseSql, onProgress) {
    var all = [];
    var offset = 0;
    for (var page = 0; page < MAX_PAGES; page += 1) {
      var sql = ctx.profile.paginate(baseSql, PAGE_ROWS, offset);
      var request = {
        connectionId: ctx.connectionId,
        sql: sql,
        maxRows: PAGE_ROWS,
        timeoutMs: QUERY_TIMEOUT_MS,
      };
      // `hidesDatabase` profiles (SQLite) address a single attached database
      // and their connection `database` is a file path, not a catalog name.
      if (ctx.database && !ctx.profile.hidesDatabase) request.database = ctx.database;
      // Postgres/SQL Server/Oracle name the schema inside the SQL; passing it
      // as a request field too would make the host resolve a session that the
      // catalog query does not need.
      var result = await global.dbxPlugin.queryData(request);
      var rows = rowsToObjects(result);
      var truncated = !!(result && result.truncated);

      if (rows.length === 0) {
        // Truncation with nothing returned means one row alone exceeds the
        // byte cap; advancing would skip it silently.
        if (truncated) throw new Error("单行元数据超出宿主 8 MiB 上限，无法读取完整结构");
        return all;
      }

      all = all.concat(rows);
      offset += rows.length;
      if (onProgress) onProgress(all.length);
      if (rows.length < PAGE_ROWS && !truncated) return all;
    }
    throw new Error("Catalog query returned more than " + MAX_PAGES * PAGE_ROWS + " rows; narrow the selection");
  }

  async function listDatabases(ctx) {
    if (!ctx.profile.databasesSql) return [];
    var rows = await queryPaged(ctx, ctx.profile.databasesSql(ctx));
    return rows
      .map(function (row) {
        return text(row.name);
      })
      .filter(Boolean);
  }

  async function listSchemas(ctx) {
    if (!ctx.profile.schemasSql) return [];
    var rows = await queryPaged(ctx, ctx.profile.schemasSql(ctx));
    return rows
      .map(function (row) {
        return text(row.name);
      })
      .filter(Boolean);
  }

  async function listTables(ctx) {
    var rows = await queryPaged(ctx, ctx.profile.tablesSql(ctx));
    return rows.map(function (row) {
      return {
        name: text(row.table_name),
        type: text(row.table_type) || "BASE TABLE",
        comment: text(row.table_comment),
        rowCount: number(row.table_rows),
        engine: text(row.engine),
        collation: text(row.table_collation),
      };
    });
  }

  function formatLength(row) {
    var charLength = number(row.character_maximum_length);
    if (charLength !== null) return String(charLength);
    var precision = number(row.numeric_precision);
    if (precision === null) return "";
    var scale = number(row.numeric_scale);
    return scale === null || scale === 0 ? String(precision) : precision + "," + scale;
  }

  /**
   * `is_nullable` is the information_schema spelling ('YES'/'NO'); Oracle's
   * all_tab_columns uses `nullable` with 'Y'/'N', aliased to `nullable_flag`.
   */
  function readNullable(row) {
    if (row.is_nullable !== undefined && row.is_nullable !== null) return isYes(row.is_nullable);
    if (row.nullable_flag !== undefined && row.nullable_flag !== null) return isYes(row.nullable_flag);
    return true;
  }

  async function collect(ctx, options, onStatus) {
    var report = function (message) {
      if (onStatus) onStatus(message);
    };

    report("读取表清单…");
    var tables = await listTables(ctx);
    var wanted = options && options.tableNames && options.tableNames.length ? options.tableNames : null;
    if (wanted) {
      var keep = Object.create(null);
      wanted.forEach(function (name) {
        keep[name] = true;
      });
      tables = tables.filter(function (table) {
        return keep[table.name];
      });
    }

    var byName = Object.create(null);
    tables.forEach(function (table) {
      table.columns = [];
      table.indexes = [];
      table.foreignKeys = [];
      byName[table.name] = table;
    });

    report("读取字段…");
    var columnRows = await queryPaged(ctx, ctx.profile.columnsSql(ctx), function (count) {
      report("读取字段… " + count);
    });
    columnRows.forEach(function (row) {
      var table = byName[text(row.table_name)];
      if (!table) return;
      table.columns.push({
        name: text(row.column_name),
        position: number(row.ordinal_position),
        dataType: text(row.data_type),
        columnType: text(row.column_type) || text(row.data_type),
        length: formatLength(row),
        nullable: readNullable(row),
        defaultValue: text(row.column_default),
        primaryKey: isYes(row.is_primary_key),
        comment: text(row.column_comment),
        extra: text(row.extra),
      });
    });

    var warnings = [];

    if (options && options.includeIndexes) {
      if (ctx.profile.indexesSql) {
        report("读取索引…");
        try {
          var indexRows = await queryPaged(ctx, ctx.profile.indexesSql(ctx));
          var indexMap = Object.create(null);
          indexRows.forEach(function (row) {
            var tableName = text(row.table_name);
            var table = byName[tableName];
            if (!table) return;
            var indexName = text(row.index_name);
            var key = tableName + "\u0000" + indexName;
            var index = indexMap[key];
            if (!index) {
              index = {
                name: indexName,
                unique: number(row.non_unique) === 0,
                primary: isYes(row.is_primary),
                type: text(row.index_type),
                columns: [],
              };
              indexMap[key] = index;
              table.indexes.push(index);
            }
            var columnName = text(row.column_name);
            if (columnName) index.columns.push(columnName);
          });
        } catch (error) {
          warnings.push("索引元数据读取失败：" + errorText(error));
        }
      } else {
        warnings.push("当前数据库类型没有可用的索引元数据查询，已跳过索引。");
      }
    }

    if (options && options.includeForeignKeys) {
      if (ctx.profile.foreignKeysSql) {
        report("读取外键…");
        try {
          var fkRows = await queryPaged(ctx, ctx.profile.foreignKeysSql(ctx));
          var fkMap = Object.create(null);
          fkRows.forEach(function (row) {
            var tableName = text(row.table_name);
            var table = byName[tableName];
            if (!table) return;
            var constraintName = text(row.constraint_name);
            var key = tableName + "\u0000" + constraintName;
            var fk = fkMap[key];
            if (!fk) {
              fk = {
                name: constraintName,
                columns: [],
                refSchema: text(row.ref_schema),
                refTable: text(row.ref_table),
                refColumns: [],
              };
              fkMap[key] = fk;
              table.foreignKeys.push(fk);
            }
            var column = text(row.column_name);
            var refColumn = text(row.ref_column);
            if (column) fk.columns.push(column);
            if (refColumn) fk.refColumns.push(refColumn);
          });
        } catch (error) {
          warnings.push("外键元数据读取失败：" + errorText(error));
        }
      } else {
        warnings.push("当前数据库类型没有可用的外键元数据查询，已跳过外键。");
      }
    }

    tables.forEach(function (table) {
      table.columns.sort(function (a, b) {
        var left = a.position === null ? Number.MAX_SAFE_INTEGER : a.position;
        var right = b.position === null ? Number.MAX_SAFE_INTEGER : b.position;
        if (left !== right) return left - right;
        return a.name.localeCompare(b.name);
      });
      if (!table.columns.length) warnings.push("表 " + table.name + " 没有读到字段。");
    });

    if (!ctx.profile.supportsComments) {
      warnings.push("当前数据库类型不提供对象注释，注释列为空。");
    }

    return {
      generatedAt: new Date(),
      connectionName: ctx.connectionName,
      dbType: ctx.dbType,
      profileId: ctx.profile.id,
      database: ctx.database || "",
      schema: ctx.schema || "",
      tables: tables,
      warnings: warnings,
    };
  }

  function errorText(error) {
    if (!error) return "未知错误";
    if (error instanceof Error) return error.message || String(error);
    return String(error);
  }

  global.DictCollect = {
    collect: collect,
    listDatabases: listDatabases,
    listSchemas: listSchemas,
    listTables: listTables,
    queryPaged: queryPaged,
    rowsToObjects: rowsToObjects,
    errorText: errorText,
    PAGE_ROWS: PAGE_ROWS,
  };
})(window);
