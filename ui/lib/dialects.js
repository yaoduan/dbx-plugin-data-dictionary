/**
 * Per-dialect catalog SQL for the data dictionary exporter.
 *
 * Every statement is a single read-only SELECT, because `host.queryData`
 * accepts exactly one statement classified read-only by DBX and rejects
 * session-switching statements such as `USE` (the database is passed as a
 * request field instead).
 *
 * Each profile returns rows under stable aliases so the collector stays
 * dialect-agnostic:
 *   tables  -> table_name, table_type, table_comment, table_rows, engine
 *   columns -> table_name, column_name, ordinal_position, data_type,
 *              column_type, is_nullable | nullable_flag, column_default,
 *              is_primary_key, column_comment, extra
 *   indexes -> table_name, index_name, non_unique, seq_in_index, column_name,
 *              index_type, is_primary
 *   foreign keys -> table_name, constraint_name, ordinal_position,
 *              column_name, ref_schema, ref_table, ref_column
 */
(function (global) {
  "use strict";

  /** SQL string literal. Doubling quotes is the only escape ANSI SQL defines. */
  function quote(value) {
    return "'" + String(value === undefined || value === null ? "" : value).replace(/'/g, "''") + "'";
  }

  /**
   * Catalog identifiers reach the SQL text as literals, so they are checked
   * before interpolation: anything with a quote, a backslash, a semicolon, a
   * comment marker or a control character is refused instead of escaped.
   */
  function assertIdentifier(value, label) {
    var text = String(value === undefined || value === null ? "" : value);
    if (!text) throw new Error(label + " is required");
    if (text.length > 128) throw new Error(label + " is too long");
    if (/['"`;\\]/.test(text) || /--|\/\*|\*\//.test(text) || /[\u0000-\u001f]/.test(text)) {
      throw new Error(label + " contains characters that are not valid in a catalog name: " + text);
    }
    return text;
  }

  function limitOffset(sql, limit, offset) {
    return sql + " LIMIT " + limit + " OFFSET " + offset;
  }

  function fetchNext(sql, limit, offset) {
    return sql + " OFFSET " + offset + " ROWS FETCH NEXT " + limit + " ROWS ONLY";
  }

  /** ROWNUM wrapping works on every Oracle version, unlike OFFSET/FETCH (12c+). */
  function rownumWindow(sql, limit, offset) {
    return (
      "SELECT * FROM (SELECT a.*, ROWNUM AS dbx_rnum FROM (" +
      sql +
      ") a WHERE ROWNUM <= " +
      (offset + limit) +
      ") WHERE dbx_rnum > " +
      offset
    );
  }

  var PROFILES = {};

  PROFILES.mysql = {
    id: "mysql",
    label: "MySQL / MariaDB",
    schemaIsDatabase: true,
    supportsComments: true,
    paginate: limitOffset,
    scope: function (ctx) {
      return assertIdentifier(ctx.schema || ctx.database, "database");
    },
    databasesSql: function () {
      return (
        "SELECT SCHEMA_NAME AS name FROM information_schema.SCHEMATA" +
        " WHERE SCHEMA_NAME NOT IN ('information_schema','performance_schema','mysql','sys')" +
        " ORDER BY SCHEMA_NAME"
      );
    },
    schemasSql: null,
    tablesSql: function (ctx) {
      return (
        "SELECT TABLE_NAME AS table_name, TABLE_TYPE AS table_type, ENGINE AS engine," +
        " TABLE_COMMENT AS table_comment, TABLE_ROWS AS table_rows, TABLE_COLLATION AS table_collation" +
        " FROM information_schema.TABLES WHERE TABLE_SCHEMA = " +
        quote(this.scope(ctx)) +
        " AND TABLE_TYPE IN ('BASE TABLE','VIEW') ORDER BY TABLE_NAME"
      );
    },
    columnsSql: function (ctx) {
      return (
        "SELECT TABLE_NAME AS table_name, COLUMN_NAME AS column_name, ORDINAL_POSITION AS ordinal_position," +
        " DATA_TYPE AS data_type, COLUMN_TYPE AS column_type, IS_NULLABLE AS is_nullable," +
        " COLUMN_DEFAULT AS column_default, COLUMN_COMMENT AS column_comment, EXTRA AS extra," +
        " CHARACTER_MAXIMUM_LENGTH AS character_maximum_length, NUMERIC_PRECISION AS numeric_precision," +
        " NUMERIC_SCALE AS numeric_scale," +
        " CASE WHEN COLUMN_KEY = 'PRI' THEN 'YES' ELSE 'NO' END AS is_primary_key" +
        " FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = " +
        quote(this.scope(ctx)) +
        " ORDER BY TABLE_NAME, ORDINAL_POSITION"
      );
    },
    indexesSql: function (ctx) {
      return (
        "SELECT TABLE_NAME AS table_name, INDEX_NAME AS index_name, NON_UNIQUE AS non_unique," +
        " SEQ_IN_INDEX AS seq_in_index, COLUMN_NAME AS column_name, INDEX_TYPE AS index_type," +
        " CASE WHEN INDEX_NAME = 'PRIMARY' THEN 'YES' ELSE 'NO' END AS is_primary" +
        " FROM information_schema.STATISTICS WHERE TABLE_SCHEMA = " +
        quote(this.scope(ctx)) +
        " ORDER BY TABLE_NAME, INDEX_NAME, SEQ_IN_INDEX"
      );
    },
    foreignKeysSql: function (ctx) {
      return (
        "SELECT TABLE_NAME AS table_name, CONSTRAINT_NAME AS constraint_name," +
        " ORDINAL_POSITION AS ordinal_position, COLUMN_NAME AS column_name," +
        " REFERENCED_TABLE_SCHEMA AS ref_schema, REFERENCED_TABLE_NAME AS ref_table," +
        " REFERENCED_COLUMN_NAME AS ref_column" +
        " FROM information_schema.KEY_COLUMN_USAGE WHERE TABLE_SCHEMA = " +
        quote(this.scope(ctx)) +
        " AND REFERENCED_TABLE_NAME IS NOT NULL" +
        " ORDER BY TABLE_NAME, CONSTRAINT_NAME, ORDINAL_POSITION"
      );
    },
  };

  PROFILES.postgres = {
    id: "postgres",
    label: "PostgreSQL",
    schemaIsDatabase: false,
    supportsComments: true,
    defaultSchema: "public",
    paginate: limitOffset,
    scope: function (ctx) {
      return assertIdentifier(ctx.schema || this.defaultSchema, "schema");
    },
    databasesSql: function () {
      return "SELECT datname AS name FROM pg_database WHERE datallowconn AND NOT datistemplate ORDER BY datname";
    },
    schemasSql: function () {
      return (
        "SELECT nspname AS name FROM pg_namespace" +
        " WHERE nspname NOT LIKE 'pg_%' AND nspname <> 'information_schema' ORDER BY nspname"
      );
    },
    tablesSql: function (ctx) {
      return (
        "SELECT c.relname AS table_name," +
        " CASE c.relkind WHEN 'r' THEN 'BASE TABLE' WHEN 'p' THEN 'BASE TABLE' WHEN 'v' THEN 'VIEW'" +
        " WHEN 'm' THEN 'MATERIALIZED VIEW' WHEN 'f' THEN 'FOREIGN TABLE' ELSE c.relkind::text END AS table_type," +
        " obj_description(c.oid, 'pg_class') AS table_comment," +
        " CASE WHEN c.reltuples < 0 THEN NULL ELSE c.reltuples::bigint END AS table_rows," +
        " NULL AS engine" +
        " FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace" +
        " WHERE n.nspname = " +
        quote(this.scope(ctx)) +
        " AND c.relkind IN ('r','p','v','m','f') ORDER BY c.relname"
      );
    },
    columnsSql: function (ctx) {
      var scope = quote(this.scope(ctx));
      return (
        "SELECT c.relname AS table_name, a.attname AS column_name, a.attnum AS ordinal_position," +
        " format_type(a.atttypid, NULL) AS data_type," +
        " format_type(a.atttypid, a.atttypmod) AS column_type," +
        " CASE WHEN a.attnotnull THEN 'NO' ELSE 'YES' END AS is_nullable," +
        " pg_get_expr(d.adbin, d.adrelid) AS column_default," +
        " col_description(c.oid, a.attnum) AS column_comment," +
        " CASE WHEN a.attidentity <> '' THEN 'identity' WHEN a.attgenerated <> '' THEN 'generated' ELSE NULL END AS extra," +
        " information_schema._pg_char_max_length(a.atttypid, a.atttypmod) AS character_maximum_length," +
        " information_schema._pg_numeric_precision(a.atttypid, a.atttypmod) AS numeric_precision," +
        " information_schema._pg_numeric_scale(a.atttypid, a.atttypmod) AS numeric_scale," +
        " CASE WHEN EXISTS (SELECT 1 FROM pg_index i WHERE i.indrelid = c.oid AND i.indisprimary" +
        " AND a.attnum = ANY (i.indkey)) THEN 'YES' ELSE 'NO' END AS is_primary_key" +
        " FROM pg_attribute a" +
        " JOIN pg_class c ON c.oid = a.attrelid" +
        " JOIN pg_namespace n ON n.oid = c.relnamespace" +
        " LEFT JOIN pg_attrdef d ON d.adrelid = c.oid AND d.adnum = a.attnum" +
        " WHERE n.nspname = " +
        scope +
        " AND c.relkind IN ('r','p','v','m','f') AND a.attnum > 0 AND NOT a.attisdropped" +
        " ORDER BY c.relname, a.attnum"
      );
    },
    indexesSql: function (ctx) {
      return (
        "SELECT t.relname AS table_name, i.relname AS index_name," +
        " CASE WHEN x.indisunique THEN 0 ELSE 1 END AS non_unique," +
        " k.ord AS seq_in_index, a.attname AS column_name," +
        " am.amname AS index_type," +
        " CASE WHEN x.indisprimary THEN 'YES' ELSE 'NO' END AS is_primary" +
        " FROM pg_index x" +
        " JOIN pg_class t ON t.oid = x.indrelid" +
        " JOIN pg_class i ON i.oid = x.indexrelid" +
        " JOIN pg_namespace n ON n.oid = t.relnamespace" +
        " JOIN pg_am am ON am.oid = i.relam" +
        " JOIN LATERAL unnest(x.indkey) WITH ORDINALITY AS k(attnum, ord) ON TRUE" +
        " LEFT JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = k.attnum" +
        " WHERE n.nspname = " +
        quote(this.scope(ctx)) +
        " ORDER BY t.relname, i.relname, k.ord"
      );
    },
    foreignKeysSql: function (ctx) {
      return (
        "SELECT tc.table_name AS table_name, tc.constraint_name AS constraint_name," +
        " kcu.ordinal_position AS ordinal_position, kcu.column_name AS column_name," +
        " ccu.table_schema AS ref_schema, ccu.table_name AS ref_table, ccu.column_name AS ref_column" +
        " FROM information_schema.table_constraints tc" +
        " JOIN information_schema.key_column_usage kcu" +
        " ON kcu.constraint_name = tc.constraint_name AND kcu.constraint_schema = tc.constraint_schema" +
        " JOIN information_schema.constraint_column_usage ccu" +
        " ON ccu.constraint_name = tc.constraint_name AND ccu.constraint_schema = tc.constraint_schema" +
        " WHERE tc.constraint_type = 'FOREIGN KEY' AND tc.table_schema = " +
        quote(this.scope(ctx)) +
        " ORDER BY tc.table_name, tc.constraint_name, kcu.ordinal_position"
      );
    },
  };

  PROFILES.sqlserver = {
    id: "sqlserver",
    label: "SQL Server",
    schemaIsDatabase: false,
    supportsComments: true,
    defaultSchema: "dbo",
    paginate: fetchNext,
    scope: function (ctx) {
      return assertIdentifier(ctx.schema || this.defaultSchema, "schema");
    },
    databasesSql: function () {
      return "SELECT name AS name FROM sys.databases WHERE state = 0 ORDER BY name";
    },
    schemasSql: function () {
      return (
        "SELECT name AS name FROM sys.schemas" +
        " WHERE name NOT IN ('sys','INFORMATION_SCHEMA','guest','db_owner','db_accessadmin'," +
        "'db_securityadmin','db_ddladmin','db_backupoperator','db_datareader','db_datawriter'," +
        "'db_denydatareader','db_denydatawriter') ORDER BY name"
      );
    },
    tablesSql: function (ctx) {
      return (
        "SELECT o.name AS table_name," +
        " CASE o.type WHEN 'U' THEN 'BASE TABLE' ELSE 'VIEW' END AS table_type," +
        " CAST(ep.value AS nvarchar(max)) AS table_comment," +
        " (SELECT SUM(p.rows) FROM sys.partitions p WHERE p.object_id = o.object_id AND p.index_id IN (0,1)) AS table_rows," +
        " NULL AS engine" +
        " FROM sys.objects o" +
        " JOIN sys.schemas s ON s.schema_id = o.schema_id" +
        " LEFT JOIN sys.extended_properties ep ON ep.major_id = o.object_id AND ep.minor_id = 0 AND ep.name = 'MS_Description'" +
        " WHERE o.type IN ('U','V') AND s.name = " +
        quote(this.scope(ctx)) +
        " ORDER BY o.name"
      );
    },
    columnsSql: function (ctx) {
      return (
        "SELECT o.name AS table_name, c.name AS column_name, c.column_id AS ordinal_position," +
        " t.name AS data_type," +
        " t.name + CASE WHEN t.name IN ('decimal','numeric') THEN '(' + CAST(c.precision AS varchar(10)) + ',' + CAST(c.scale AS varchar(10)) + ')'" +
        " WHEN t.name IN ('varchar','char','varbinary','binary') AND c.max_length = -1 THEN '(max)'" +
        " WHEN t.name IN ('varchar','char','varbinary','binary') THEN '(' + CAST(c.max_length AS varchar(10)) + ')'" +
        " WHEN t.name IN ('nvarchar','nchar') AND c.max_length = -1 THEN '(max)'" +
        " WHEN t.name IN ('nvarchar','nchar') THEN '(' + CAST(c.max_length / 2 AS varchar(10)) + ')'" +
        " ELSE '' END AS column_type," +
        " CASE WHEN c.is_nullable = 1 THEN 'YES' ELSE 'NO' END AS is_nullable," +
        " dc.definition AS column_default," +
        " CAST(ep.value AS nvarchar(max)) AS column_comment," +
        " CASE WHEN c.is_identity = 1 THEN 'identity' WHEN c.is_computed = 1 THEN 'computed' ELSE NULL END AS extra," +
        " CASE WHEN t.name IN ('nvarchar','nchar') AND c.max_length > 0 THEN c.max_length / 2" +
        " WHEN c.max_length = -1 THEN NULL ELSE c.max_length END AS character_maximum_length," +
        " c.precision AS numeric_precision, c.scale AS numeric_scale," +
        " CASE WHEN EXISTS (SELECT 1 FROM sys.index_columns ic JOIN sys.indexes ix" +
        " ON ix.object_id = ic.object_id AND ix.index_id = ic.index_id" +
        " WHERE ic.object_id = o.object_id AND ic.column_id = c.column_id AND ix.is_primary_key = 1)" +
        " THEN 'YES' ELSE 'NO' END AS is_primary_key" +
        " FROM sys.columns c" +
        " JOIN sys.objects o ON o.object_id = c.object_id" +
        " JOIN sys.schemas s ON s.schema_id = o.schema_id" +
        " JOIN sys.types t ON t.user_type_id = c.user_type_id" +
        " LEFT JOIN sys.default_constraints dc ON dc.object_id = c.default_object_id" +
        " LEFT JOIN sys.extended_properties ep ON ep.major_id = o.object_id AND ep.minor_id = c.column_id AND ep.name = 'MS_Description'" +
        " WHERE o.type IN ('U','V') AND s.name = " +
        quote(this.scope(ctx)) +
        " ORDER BY o.name, c.column_id"
      );
    },
    indexesSql: function (ctx) {
      return (
        "SELECT o.name AS table_name, ix.name AS index_name," +
        " CASE WHEN ix.is_unique = 1 THEN 0 ELSE 1 END AS non_unique," +
        " ic.key_ordinal AS seq_in_index, c.name AS column_name, ix.type_desc AS index_type," +
        " CASE WHEN ix.is_primary_key = 1 THEN 'YES' ELSE 'NO' END AS is_primary" +
        " FROM sys.indexes ix" +
        " JOIN sys.objects o ON o.object_id = ix.object_id" +
        " JOIN sys.schemas s ON s.schema_id = o.schema_id" +
        " JOIN sys.index_columns ic ON ic.object_id = ix.object_id AND ic.index_id = ix.index_id" +
        " JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id" +
        " WHERE ix.name IS NOT NULL AND o.type IN ('U','V') AND s.name = " +
        quote(this.scope(ctx)) +
        " ORDER BY o.name, ix.name, ic.key_ordinal"
      );
    },
    foreignKeysSql: function (ctx) {
      return (
        "SELECT po.name AS table_name, fk.name AS constraint_name," +
        " fkc.constraint_column_id AS ordinal_position, pc.name AS column_name," +
        " rs.name AS ref_schema, ro.name AS ref_table, rc.name AS ref_column" +
        " FROM sys.foreign_keys fk" +
        " JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id" +
        " JOIN sys.objects po ON po.object_id = fk.parent_object_id" +
        " JOIN sys.schemas ps ON ps.schema_id = po.schema_id" +
        " JOIN sys.columns pc ON pc.object_id = fkc.parent_object_id AND pc.column_id = fkc.parent_column_id" +
        " JOIN sys.objects ro ON ro.object_id = fk.referenced_object_id" +
        " JOIN sys.schemas rs ON rs.schema_id = ro.schema_id" +
        " JOIN sys.columns rc ON rc.object_id = fkc.referenced_object_id AND rc.column_id = fkc.referenced_column_id" +
        " WHERE ps.name = " +
        quote(this.scope(ctx)) +
        " ORDER BY po.name, fk.name, fkc.constraint_column_id"
      );
    },
  };

  PROFILES.oracle = {
    id: "oracle",
    label: "Oracle",
    schemaIsDatabase: false,
    supportsComments: true,
    paginate: rownumWindow,
    scope: function (ctx) {
      return assertIdentifier(ctx.schema, "schema").toUpperCase();
    },
    databasesSql: null,
    schemasSql: function () {
      return "SELECT username AS name FROM all_users ORDER BY username";
    },
    tablesSql: function (ctx) {
      var scope = quote(this.scope(ctx));
      return (
        "SELECT t.table_name AS table_name, 'BASE TABLE' AS table_type," +
        " c.comments AS table_comment, t.num_rows AS table_rows, NULL AS engine" +
        " FROM all_tables t" +
        " LEFT JOIN all_tab_comments c ON c.owner = t.owner AND c.table_name = t.table_name" +
        " WHERE t.owner = " +
        scope +
        " UNION ALL" +
        " SELECT v.view_name AS table_name, 'VIEW' AS table_type," +
        " c.comments AS table_comment, NULL AS table_rows, NULL AS engine" +
        " FROM all_views v" +
        " LEFT JOIN all_tab_comments c ON c.owner = v.owner AND c.table_name = v.view_name" +
        " WHERE v.owner = " +
        scope +
        " ORDER BY table_name"
      );
    },
    columnsSql: function (ctx) {
      var scope = quote(this.scope(ctx));
      return (
        "SELECT c.table_name AS table_name, c.column_name AS column_name," +
        " c.column_id AS ordinal_position, c.data_type AS data_type," +
        " c.data_type || CASE WHEN c.data_type IN ('NUMBER') AND c.data_precision IS NOT NULL" +
        " THEN '(' || c.data_precision || ',' || NVL(c.data_scale, 0) || ')'" +
        " WHEN c.data_type IN ('VARCHAR2','NVARCHAR2','CHAR','NCHAR','RAW')" +
        " THEN '(' || c.char_length || ')' ELSE '' END AS column_type," +
        " c.nullable AS nullable_flag, c.data_default AS column_default," +
        " cc.comments AS column_comment, NULL AS extra," +
        " c.char_length AS character_maximum_length, c.data_precision AS numeric_precision," +
        " c.data_scale AS numeric_scale," +
        " CASE WHEN EXISTS (SELECT 1 FROM all_constraints ac" +
        " JOIN all_cons_columns acc ON acc.owner = ac.owner AND acc.constraint_name = ac.constraint_name" +
        " WHERE ac.constraint_type = 'P' AND ac.owner = c.owner AND ac.table_name = c.table_name" +
        " AND acc.column_name = c.column_name) THEN 'YES' ELSE 'NO' END AS is_primary_key" +
        " FROM all_tab_columns c" +
        " LEFT JOIN all_col_comments cc ON cc.owner = c.owner AND cc.table_name = c.table_name" +
        " AND cc.column_name = c.column_name" +
        " WHERE c.owner = " +
        scope +
        " ORDER BY c.table_name, c.column_id"
      );
    },
    indexesSql: function (ctx) {
      var scope = quote(this.scope(ctx));
      return (
        "SELECT i.table_name AS table_name, i.index_name AS index_name," +
        " CASE WHEN i.uniqueness = 'UNIQUE' THEN 0 ELSE 1 END AS non_unique," +
        " ic.column_position AS seq_in_index, ic.column_name AS column_name," +
        " i.index_type AS index_type," +
        " CASE WHEN EXISTS (SELECT 1 FROM all_constraints ac WHERE ac.constraint_type = 'P'" +
        " AND ac.owner = i.owner AND ac.index_name = i.index_name) THEN 'YES' ELSE 'NO' END AS is_primary" +
        " FROM all_indexes i" +
        " JOIN all_ind_columns ic ON ic.index_owner = i.owner AND ic.index_name = i.index_name" +
        " WHERE i.owner = " +
        scope +
        " ORDER BY i.table_name, i.index_name, ic.column_position"
      );
    },
    foreignKeysSql: function (ctx) {
      var scope = quote(this.scope(ctx));
      return (
        "SELECT ac.table_name AS table_name, ac.constraint_name AS constraint_name," +
        " acc.position AS ordinal_position, acc.column_name AS column_name," +
        " rcc.owner AS ref_schema, rcc.table_name AS ref_table, rcc.column_name AS ref_column" +
        " FROM all_constraints ac" +
        " JOIN all_cons_columns acc ON acc.owner = ac.owner AND acc.constraint_name = ac.constraint_name" +
        " JOIN all_cons_columns rcc ON rcc.owner = ac.r_owner" +
        " AND rcc.constraint_name = ac.r_constraint_name AND rcc.position = acc.position" +
        " WHERE ac.constraint_type = 'R' AND ac.owner = " +
        scope +
        " ORDER BY ac.table_name, ac.constraint_name, acc.position"
      );
    },
  };

  PROFILES.sqlite = {
    id: "sqlite",
    label: "SQLite",
    schemaIsDatabase: true,
    supportsComments: false,
    // One attached database per connection, and its `database` is a file path,
    // so the field is hidden and never interpolated into the catalog SQL.
    hidesDatabase: true,
    paginate: limitOffset,
    scope: function () {
      return "main";
    },
    databasesSql: null,
    schemasSql: null,
    tablesSql: function () {
      return (
        "SELECT name AS table_name," +
        " CASE type WHEN 'table' THEN 'BASE TABLE' ELSE 'VIEW' END AS table_type," +
        " NULL AS table_comment, NULL AS table_rows, NULL AS engine" +
        " FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%'" +
        " ORDER BY name"
      );
    },
    // pragma_* table-valued functions need SQLite 3.16+; DBX ships a newer one.
    columnsSql: function () {
      return (
        "SELECT m.name AS table_name, p.name AS column_name, p.cid + 1 AS ordinal_position," +
        " p.type AS data_type, p.type AS column_type," +
        " CASE WHEN p.\"notnull\" = 1 THEN 'NO' ELSE 'YES' END AS is_nullable," +
        " p.dflt_value AS column_default, NULL AS column_comment, NULL AS extra," +
        " NULL AS character_maximum_length, NULL AS numeric_precision, NULL AS numeric_scale," +
        " CASE WHEN p.pk > 0 THEN 'YES' ELSE 'NO' END AS is_primary_key" +
        " FROM sqlite_master m JOIN pragma_table_info(m.name) p" +
        " WHERE m.type IN ('table','view') AND m.name NOT LIKE 'sqlite_%'" +
        " ORDER BY m.name, p.cid"
      );
    },
    indexesSql: function () {
      return (
        "SELECT m.name AS table_name, il.name AS index_name," +
        " CASE WHEN il.\"unique\" = 1 THEN 0 ELSE 1 END AS non_unique," +
        " ii.seqno + 1 AS seq_in_index, ii.name AS column_name," +
        " il.origin AS index_type," +
        " CASE WHEN il.origin = 'pk' THEN 'YES' ELSE 'NO' END AS is_primary" +
        " FROM sqlite_master m" +
        " JOIN pragma_index_list(m.name) il" +
        " JOIN pragma_index_info(il.name) ii" +
        " WHERE m.type = 'table' AND m.name NOT LIKE 'sqlite_%'" +
        " ORDER BY m.name, il.name, ii.seqno"
      );
    },
    foreignKeysSql: function () {
      return (
        "SELECT m.name AS table_name," +
        " 'fk_' || m.name || '_' || f.id AS constraint_name," +
        " f.seq + 1 AS ordinal_position, f.\"from\" AS column_name," +
        " NULL AS ref_schema, f.\"table\" AS ref_table, f.\"to\" AS ref_column" +
        " FROM sqlite_master m JOIN pragma_foreign_key_list(m.name) f" +
        " WHERE m.type = 'table' AND m.name NOT LIKE 'sqlite_%'" +
        " ORDER BY m.name, f.id, f.seq"
      );
    },
  };

  /**
   * ANSI information_schema fallback. Used for dialects without a dedicated
   * profile; comments and index metadata are not part of the standard views, so
   * those sections come back empty rather than guessed.
   */
  PROFILES.ansi = {
    id: "ansi",
    label: "ANSI information_schema",
    schemaIsDatabase: false,
    supportsComments: false,
    paginate: limitOffset,
    scope: function (ctx) {
      return assertIdentifier(ctx.schema || ctx.database, "schema");
    },
    databasesSql: null,
    schemasSql: function () {
      return "SELECT schema_name AS name FROM information_schema.schemata ORDER BY schema_name";
    },
    tablesSql: function (ctx) {
      return (
        "SELECT table_name AS table_name, table_type AS table_type," +
        " NULL AS table_comment, NULL AS table_rows, NULL AS engine" +
        " FROM information_schema.tables WHERE table_schema = " +
        quote(this.scope(ctx)) +
        " ORDER BY table_name"
      );
    },
    columnsSql: function (ctx) {
      return (
        "SELECT table_name AS table_name, column_name AS column_name," +
        " ordinal_position AS ordinal_position, data_type AS data_type, data_type AS column_type," +
        " is_nullable AS is_nullable, column_default AS column_default," +
        " NULL AS column_comment, NULL AS extra," +
        " character_maximum_length AS character_maximum_length," +
        " numeric_precision AS numeric_precision, numeric_scale AS numeric_scale," +
        " 'NO' AS is_primary_key" +
        " FROM information_schema.columns WHERE table_schema = " +
        quote(this.scope(ctx)) +
        " ORDER BY table_name, ordinal_position"
      );
    },
    indexesSql: null,
    foreignKeysSql: function (ctx) {
      return (
        "SELECT tc.table_name AS table_name, tc.constraint_name AS constraint_name," +
        " kcu.ordinal_position AS ordinal_position, kcu.column_name AS column_name," +
        " ccu.table_schema AS ref_schema, ccu.table_name AS ref_table, ccu.column_name AS ref_column" +
        " FROM information_schema.table_constraints tc" +
        " JOIN information_schema.key_column_usage kcu" +
        " ON kcu.constraint_name = tc.constraint_name AND kcu.constraint_schema = tc.constraint_schema" +
        " JOIN information_schema.constraint_column_usage ccu" +
        " ON ccu.constraint_name = tc.constraint_name AND ccu.constraint_schema = tc.constraint_schema" +
        " WHERE tc.constraint_type = 'FOREIGN KEY' AND tc.table_schema = " +
        quote(this.scope(ctx)) +
        " ORDER BY tc.table_name, tc.constraint_name, kcu.ordinal_position"
      );
    },
  };

  /** dbType -> profile. Keys are DBX `dbType` values (generated/databaseTypes.ts). */
  var DB_TYPE_PROFILES = {
    mysql: "mysql",
    mariadb: "mysql",
    oceanbase: "mysql",
    goldendb: "mysql",
    doris: "mysql",
    starrocks: "mysql",
    gbase: "mysql",
    tidb: "mysql",
    postgres: "postgres",
    postgresql: "postgres",
    redshift: "postgres",
    gaussdb: "postgres",
    opengauss: "postgres",
    kingbase: "postgres",
    highgo: "postgres",
    vastbase: "postgres",
    uxdb: "postgres",
    gbase8c: "postgres",
    greenplum: "postgres",
    cockroachdb: "postgres",
    sqlserver: "sqlserver",
    mssql: "sqlserver",
    oracle: "oracle",
    "oceanbase-oracle": "oracle",
    dameng: "oracle",
    sqlite: "sqlite",
    turso: "sqlite",
    rqlite: "sqlite",
    "cloudflare-d1": "sqlite",
    duckdb: "ansi",
    clickhouse: "ansi",
    databend: "ansi",
    trino: "ansi",
    prestosql: "ansi",
    hive: "ansi",
    spark: "ansi",
    kyuubi: "ansi",
    impala: "ansi",
    snowflake: "ansi",
    vertica: "ansi",
    exasol: "ansi",
    h2: "ansi",
    db2: "ansi",
    saphana: "ansi",
    teradata: "ansi",
    yashandb: "ansi",
    bigquery: "ansi",
  };

  function profileFor(dbType) {
    var key = String(dbType || "").toLowerCase();
    var name = DB_TYPE_PROFILES[key];
    return name ? PROFILES[name] : null;
  }

  global.DictDialects = {
    quote: quote,
    assertIdentifier: assertIdentifier,
    profiles: PROFILES,
    profileFor: profileFor,
    dbTypeProfiles: DB_TYPE_PROFILES,
  };
})(window);
