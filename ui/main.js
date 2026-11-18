/**
 * Data dictionary exporter workbench.
 *
 * Reads catalog metadata through the host's consent-gated read-only data API
 * and writes a .docx or .xlsx file through the host's native save dialog. No
 * driver, no credentials and no network access are involved: the plugin only
 * ever sends one read-only SELECT at a time for a connection the user granted.
 */
(function () {
  "use strict";

  var dbx = window.dbxPlugin;

  var state = {
    connections: [],
    connection: null,
    profile: null,
    tables: [],
    selected: Object.create(null),
    filter: "",
    busy: false,
  };

  var I18N = {
    zh: {
      title: "数据字典导出",
      subtitle: "把所选数据库的表结构导出为 DOCX 或 Excel 文档。",
      connection: "连接",
      database: "数据库",
      schema: "Schema",
      format: "导出格式",
      layout: "Excel 布局",
      indexes: "包含索引",
      foreignKeys: "包含外键",
      loadTables: "读取表清单",
      exportDoc: "导出文档",
      tables: "表选择",
      refresh: "刷新",
      selectAll: "全选",
      selectNone: "清空",
      filterPlaceholder: "搜索表名或注释",
      formatDocx: "Word 文档（.docx）",
      formatXlsx: "Excel 工作簿（.xlsx）",
      layoutSingle: "汇总表（字段明细 + 索引 + 外键）",
      layoutPerTable: "每张表一个工作表",
    },
    en: {
      title: "Data dictionary export",
      subtitle: "Export the selected database schema as a DOCX or Excel document.",
      connection: "Connection",
      database: "Database",
      schema: "Schema",
      format: "Format",
      layout: "Excel layout",
      indexes: "Include indexes",
      foreignKeys: "Include foreign keys",
      loadTables: "Load tables",
      exportDoc: "Export document",
      tables: "Tables",
      refresh: "Refresh",
      selectAll: "All",
      selectNone: "None",
      filterPlaceholder: "Search name or comment",
      formatDocx: "Word document (.docx)",
      formatXlsx: "Excel workbook (.xlsx)",
      layoutSingle: "Summary sheets (columns + indexes + foreign keys)",
      layoutPerTable: "One sheet per table",
    },
  };

  var messages = {
    zh: {
      noDataApi: "当前 DBX 宿主不支持插件只读数据查询（host.data:read / dataApi），无法读取数据库结构。",
      noConnectionsApi: "当前 DBX 宿主不支持插件读取连接列表（host.connections:read / connectionsApi）。",
      noConnections: "没有已打开的 SQL 连接。请先在 DBX 中打开一个关系型数据库连接，再回到本页刷新。",
      unsupportedDbType: "连接类型 {0} 还没有对应的字典查询实现。",
      pickConnection: "请选择一个连接。",
      needDatabase: "请填写数据库名。",
      needSchema: "请填写 Schema。",
      loadingTables: "读取表清单…",
      tablesLoaded: "共 {0} 张表/视图，已选 {1}。",
      noTables: "该范围内没有表或视图。",
      selectTables: "请至少选择一张表。",
      exporting: "正在生成文档…",
      saving: "正在保存文件…",
      saved: "已导出：{0}",
      savedNoPath: "已导出文件。",
      cancelled: "已取消保存。",
      selectionCount: "共 {0} 项，已选 {1} 项。",
      noSaveApi: "当前宿主不提供文件保存能力，无法导出。",
      grantNeeded: "需要你在 DBX 弹窗中允许本插件读取该连接。",
      listLoaded: "读到 {0} 项。",
      warningsTitle: "导出完成，但有以下降级：",
    },
    en: {
      noDataApi: "This DBX host does not support plugin read-only data queries (host.data:read / dataApi).",
      noConnectionsApi: "This DBX host does not support plugin connection listing (host.connections:read / connectionsApi).",
      noConnections: "No open SQL connection. Open a relational connection in DBX, then reload this tab.",
      unsupportedDbType: "Connection type {0} has no dictionary query implementation yet.",
      pickConnection: "Select a connection.",
      needDatabase: "Enter a database name.",
      needSchema: "Enter a schema.",
      loadingTables: "Loading tables…",
      tablesLoaded: "{0} tables/views, {1} selected.",
      noTables: "No tables or views in this scope.",
      selectTables: "Select at least one table.",
      exporting: "Building the document…",
      saving: "Saving the file…",
      saved: "Exported: {0}",
      savedNoPath: "File exported.",
      cancelled: "Save cancelled.",
      selectionCount: "{0} items, {1} selected.",
      noSaveApi: "This host offers no file saving, so the document cannot be exported.",
      grantNeeded: "Allow this plugin to read the connection in the DBX prompt.",
      listLoaded: "{0} entries loaded.",
      warningsTitle: "Exported, with these degradations:",
    },
  };

  var lang = "zh";

  function t(key) {
    var table = I18N[lang] || I18N.zh;
    return table[key] || (I18N.zh[key] || key);
  }

  function m(key) {
    var args = Array.prototype.slice.call(arguments, 1);
    var table = messages[lang] || messages.zh;
    var template = table[key] || messages.zh[key] || key;
    return template.replace(/\{(\d+)\}/g, function (match, index) {
      var value = args[Number(index)];
      return value === undefined || value === null ? "" : String(value);
    });
  }

  function el(id) {
    return document.getElementById(id);
  }

  var ui = {};

  function cacheElements() {
    [
      "title",
      "subtitle",
      "unsupported",
      "unsupported-text",
      "setup",
      "connection",
      "database-field",
      "database",
      "database-options",
      "reload-database",
      "schema-field",
      "schema",
      "schema-options",
      "reload-schema",
      "format",
      "layout-field",
      "layout",
      "opt-indexes",
      "opt-foreign-keys",
      "load-tables",
      "export",
      "status",
      "error",
      "warnings",
      "tables-card",
      "filter",
      "select-all",
      "select-none",
      "selection-count",
      "table-list",
      "label-connection",
      "label-database",
      "label-schema",
      "label-format",
      "label-layout",
      "label-indexes",
      "label-foreign-keys",
      "label-tables",
    ].forEach(function (id) {
      ui[id] = el(id);
    });
  }

  function applyLanguage() {
    document.documentElement.lang = lang === "zh" ? "zh-CN" : "en";
    ui.title.textContent = t("title");
    ui.subtitle.textContent = t("subtitle");
    ui["label-connection"].textContent = t("connection");
    ui["label-database"].textContent = t("database");
    ui["label-schema"].textContent = t("schema");
    ui["label-format"].textContent = t("format");
    ui["label-layout"].textContent = t("layout");
    ui["label-indexes"].textContent = t("indexes");
    ui["label-foreign-keys"].textContent = t("foreignKeys");
    ui["label-tables"].textContent = t("tables");
    ui["load-tables"].textContent = t("loadTables");
    ui.export.textContent = t("exportDoc");
    ui["reload-database"].textContent = t("refresh");
    ui["reload-schema"].textContent = t("refresh");
    ui["select-all"].textContent = t("selectAll");
    ui["select-none"].textContent = t("selectNone");
    ui.filter.placeholder = t("filterPlaceholder");
    ui.format.options[0].textContent = t("formatDocx");
    ui.format.options[1].textContent = t("formatXlsx");
    ui.layout.options[0].textContent = t("layoutSingle");
    ui.layout.options[1].textContent = t("layoutPerTable");
  }

  function setStatus(text) {
    ui.status.textContent = text || "";
  }

  function showError(error) {
    if (!error) {
      ui.error.hidden = true;
      ui.error.textContent = "";
      return;
    }
    var text = window.DictCollect.errorText(error);
    if (text.indexOf("PLUGIN_DATA_ACCESS_NOT_GRANTED") >= 0) text = m("grantNeeded") + "\n" + text;
    ui.error.hidden = false;
    ui.error.textContent = text;
  }

  /**
   * Degradations collected during a successful export (a dialect without
   * comments, a catalog view the server refused). They belong next to the
   * result, not in the error slot.
   */
  function showWarnings(warnings) {
    if (!warnings || !warnings.length) {
      ui.warnings.hidden = true;
      ui.warnings.textContent = "";
      return;
    }
    ui.warnings.hidden = false;
    ui.warnings.textContent = m("warningsTitle") + "\n" + warnings.map(function (warning) {
      return "• " + warning;
    }).join("\n");
  }

  function setBusy(busy) {
    state.busy = busy;
    ui["load-tables"].disabled = busy;
    ui["reload-database"].disabled = busy;
    ui["reload-schema"].disabled = busy;
    ui.connection.disabled = busy;
    ui.export.disabled = busy || !selectedNames().length;
  }

  /* ----------------------------------------------------------- bootstrap */

  function blockWith(text) {
    ui["unsupported-text"].textContent = text;
    ui.unsupported.hidden = false;
    ui.setup.hidden = true;
    ui["tables-card"].hidden = true;
  }

  function isSqlConnection(connection) {
    return !!window.DictDialects.profileFor(connection.dbType || connection.dialect);
  }

  function describeConnection(connection) {
    var parts = [connection.name || connection.id];
    var type = connection.dbType || connection.dialect;
    if (type) parts.push(type);
    if (!connection.connected && !connection.isConnected) parts.push(lang === "zh" ? "未连接" : "not open");
    return parts.join(" · ");
  }

  function renderConnections() {
    ui.connection.innerHTML = "";
    state.connections.forEach(function (connection) {
      var option = document.createElement("option");
      option.value = connection.id;
      option.textContent = describeConnection(connection);
      ui.connection.appendChild(option);
    });
  }

  function currentConnection() {
    var id = ui.connection.value;
    for (var i = 0; i < state.connections.length; i += 1) {
      if (state.connections[i].id === id) return state.connections[i];
    }
    return null;
  }

  function syncProfile() {
    var connection = currentConnection();
    state.connection = connection;
    state.profile = connection ? window.DictDialects.profileFor(connection.dbType || connection.dialect) : null;
    var profile = state.profile;

    // SQLite addresses one attached database, so there is nothing to pick; the
    // connection's `database` is a file path and must not reach the SQL.
    var hidesDatabase = !!(profile && profile.hidesDatabase);
    var schemaIsDatabase = !profile || profile.schemaIsDatabase;
    ui["database-field"].hidden = hidesDatabase;
    ui["schema-field"].hidden = schemaIsDatabase;

    // A refresh button with no catalog query behind it would silently do
    // nothing, so it is hidden instead of disabled.
    ui["reload-database"].hidden = !profile || !profile.databasesSql;
    ui["reload-schema"].hidden = !profile || !profile.schemasSql;

    if (connection && !hidesDatabase && !ui.database.value) {
      ui.database.value = connection.database || "";
    }
    if (connection && !schemaIsDatabase && !ui.schema.value) {
      ui.schema.value = connection.schema || (profile && profile.defaultSchema) || "";
    }
  }

  /**
   * Build the collector context. The profile's own `scope()` decides which of
   * database/schema it needs, so validation stays in one place instead of being
   * re-encoded per dialect here.
   */
  function contextFor() {
    var connection = state.connection;
    if (!connection) throw new Error(m("pickConnection"));
    var profile = state.profile;
    if (!profile) throw new Error(m("unsupportedDbType", connection.dbType || connection.dialect || "?"));

    var ctx = {
      connectionId: connection.id,
      connectionName: connection.name || connection.id,
      dbType: connection.dbType || connection.dialect || "",
      database: ui["database-field"].hidden ? "" : ui.database.value.trim(),
      schema: ui["schema-field"].hidden ? "" : ui.schema.value.trim(),
      profile: profile,
    };

    try {
      profile.scope(ctx);
    } catch (error) {
      var message = window.DictCollect.errorText(error);
      // Turn the profile's "database is required" / "schema is required" into
      // the localized prompt, but keep any other validation message verbatim.
      if (/^database is required/.test(message)) throw new Error(m("needDatabase"));
      if (/^schema is required/.test(message)) throw new Error(m("needSchema"));
      throw error;
    }
    return ctx;
  }

  /** Context for catalog listings, which do not need a resolved scope. */
  function listingContext() {
    var connection = state.connection;
    if (!connection) throw new Error(m("pickConnection"));
    if (!state.profile) throw new Error(m("unsupportedDbType", connection.dbType || connection.dialect || "?"));
    return {
      connectionId: connection.id,
      connectionName: connection.name || connection.id,
      dbType: connection.dbType || connection.dialect || "",
      database: ui["database-field"].hidden ? "" : ui.database.value.trim(),
      schema: ui["schema-field"].hidden ? "" : ui.schema.value.trim(),
      profile: state.profile,
    };
  }

  function exportOptions() {
    return {
      format: ui.format.value,
      excelLayout: ui.layout.value,
      includeIndexes: ui["opt-indexes"].checked,
      includeForeignKeys: ui["opt-foreign-keys"].checked,
    };
  }

  /* ------------------------------------------------------------- storage */

  var STORAGE_KEY = "exporter-state";

  function canStore() {
    return !!(dbx.capabilities && dbx.capabilities.storage && dbx.storage);
  }

  async function saveSettings() {
    if (!canStore()) return;
    try {
      await dbx.storage.set(STORAGE_KEY, {
        connectionId: ui.connection.value,
        database: ui.database.value,
        schema: ui.schema.value,
        format: ui.format.value,
        layout: ui.layout.value,
        includeIndexes: ui["opt-indexes"].checked,
        includeForeignKeys: ui["opt-foreign-keys"].checked,
      });
    } catch (error) {
      // Settings are a convenience; a failed write must not break the export.
    }
  }

  async function restoreSettings() {
    if (!canStore()) return null;
    try {
      return await dbx.storage.get(STORAGE_KEY);
    } catch (error) {
      return null;
    }
  }

  /* --------------------------------------------------------------- tables */

  function selectedNames() {
    return state.tables
      .filter(function (table) {
        return state.selected[table.name];
      })
      .map(function (table) {
        return table.name;
      });
  }

  function updateSelectionCount() {
    var selected = selectedNames().length;
    ui["selection-count"].textContent = m("selectionCount", state.tables.length, selected);
    ui.export.disabled = state.busy || selected === 0;
  }

  function renderTables() {
    var list = ui["table-list"];
    list.innerHTML = "";
    var needle = state.filter.trim().toLowerCase();
    var visible = state.tables.filter(function (table) {
      if (!needle) return true;
      return (
        table.name.toLowerCase().indexOf(needle) >= 0 ||
        (table.comment || "").toLowerCase().indexOf(needle) >= 0
      );
    });

    var fragment = document.createDocumentFragment();
    visible.forEach(function (table) {
      var row = document.createElement("label");
      row.className = "table-row";

      var checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = !!state.selected[table.name];
      checkbox.addEventListener("change", function () {
        state.selected[table.name] = checkbox.checked;
        updateSelectionCount();
      });

      var name = document.createElement("span");
      name.className = "table-name";
      name.textContent = table.name;

      var comment = document.createElement("span");
      comment.className = "table-comment";
      comment.textContent = table.comment || "";

      var meta = document.createElement("span");
      meta.className = "table-meta";
      meta.textContent = table.type + (table.rowCount === null || table.rowCount === undefined ? "" : " · ~" + table.rowCount);

      row.appendChild(checkbox);
      row.appendChild(name);
      row.appendChild(comment);
      row.appendChild(meta);
      fragment.appendChild(row);
    });
    list.appendChild(fragment);
    updateSelectionCount();
  }

  async function loadTables() {
    showError(null);
    showWarnings(null);
    setBusy(true);
    setStatus(m("loadingTables"));
    try {
      var ctx = contextFor();
      var tables = await window.DictCollect.listTables(ctx);
      state.tables = tables;
      state.selected = Object.create(null);
      tables.forEach(function (table) {
        state.selected[table.name] = true;
      });
      ui["tables-card"].hidden = tables.length === 0;
      renderTables();
      setStatus(tables.length ? m("tablesLoaded", tables.length, tables.length) : m("noTables"));
      await saveSettings();
    } catch (error) {
      state.tables = [];
      state.selected = Object.create(null);
      ui["tables-card"].hidden = true;
      showError(error);
      setStatus("");
    } finally {
      setBusy(false);
    }
  }

  async function loadDatabaseList() {
    showError(null);
    try {
      var ctx = listingContext();
      if (!ctx.profile.databasesSql) return;
      setBusy(true);
      var names = await window.DictCollect.listDatabases(ctx);
      fillDatalist(ui["database-options"], names);
      setStatus(m("listLoaded", names.length));
    } catch (error) {
      showError(error);
    } finally {
      setBusy(false);
    }
  }

  async function loadSchemaList() {
    showError(null);
    try {
      var ctx = listingContext();
      if (!ctx.profile.schemasSql) return;
      setBusy(true);
      var names = await window.DictCollect.listSchemas(ctx);
      fillDatalist(ui["schema-options"], names);
      setStatus(m("listLoaded", names.length));
    } catch (error) {
      showError(error);
    } finally {
      setBusy(false);
    }
  }

  function fillDatalist(list, names) {
    list.innerHTML = "";
    names.forEach(function (name) {
      var option = document.createElement("option");
      option.value = name;
      list.appendChild(option);
    });
  }

  /* --------------------------------------------------------------- export */

  /**
   * Hand the bytes to the host so it can run the native save dialog.
   * `saveFile` is the one-shot path; `fileTransfer.beginSave` is the chunked
   * fallback for hosts that only advertise the transfer namespace.
   */
  async function saveBytes(fileName, contentType, bytes) {
    if (typeof dbx.saveFile === "function") {
      try {
        var result = await dbx.saveFile({ fileName: fileName, contentType: contentType }, bytes);
        return { ok: result !== null && result !== undefined, path: result && result.path };
      } catch (error) {
        var transfer = dbx.capabilities && dbx.capabilities.fileTransfer;
        if (!transfer || !transfer.beginSave) throw error;
      }
    }

    var canChunk = dbx.capabilities && dbx.capabilities.fileTransfer && dbx.capabilities.fileTransfer.beginSave;
    if (!canChunk || !dbx.fileTransfer) throw new Error(m("noSaveApi"));

    var target = await dbx.fileTransfer.beginSave({ name: fileName, contentType: contentType, size: bytes.length });
    if (!target) return { ok: false };
    var chunkSize = target.chunkBytes && target.chunkBytes > 0 ? Math.min(target.chunkBytes, 4 * 1024 * 1024) : 1024 * 1024;
    try {
      for (var offset = 0; offset < bytes.length; offset += chunkSize) {
        var slice = bytes.subarray(offset, Math.min(offset + chunkSize, bytes.length));
        await dbx.fileTransfer.write(target.handleId, offset, slice);
      }
      await dbx.fileTransfer.finish(target.handleId);
      return { ok: true, path: target.path };
    } catch (error) {
      try {
        await dbx.fileTransfer.cancel(target.handleId);
      } catch (ignored) {
        // The handle is already gone; the original error is what matters.
      }
      throw error;
    }
  }

  async function runExport() {
    showError(null);
    showWarnings(null);
    var names = selectedNames();
    if (!names.length) {
      showError(new Error(m("selectTables")));
      return;
    }
    setBusy(true);
    try {
      var ctx = contextFor();
      var options = exportOptions();
      var schema = await window.DictCollect.collect(ctx, {
        tableNames: names,
        includeIndexes: options.includeIndexes,
        includeForeignKeys: options.includeForeignKeys,
      }, setStatus);

      setStatus(m("exporting"));
      var isDocx = options.format === "docx";
      var bytes = isDocx
        ? await window.DictReport.buildDocx(schema, options)
        : await window.DictReport.buildXlsx(schema, options);
      var fileName = window.DictReport.fileBaseName(schema) + (isDocx ? ".docx" : ".xlsx");
      var contentType = isDocx
        ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

      setStatus(m("saving"));
      var saved = await saveBytes(fileName, contentType, bytes);
      if (!saved.ok) {
        setStatus(m("cancelled"));
      } else {
        setStatus(saved.path ? m("saved", saved.path) : m("savedNoPath"));
      }
      showWarnings(schema.warnings);
      await saveSettings();
    } catch (error) {
      showError(error);
      setStatus("");
    } finally {
      setBusy(false);
    }
  }

  /* ----------------------------------------------------------------- wire */

  function bindEvents() {
    ui.connection.addEventListener("change", function () {
      ui.database.value = "";
      ui.schema.value = "";
      state.tables = [];
      state.selected = Object.create(null);
      ui["tables-card"].hidden = true;
      syncProfile();
      setStatus("");
      showError(null);
      showWarnings(null);
      void saveSettings();
    });
    ui.format.addEventListener("change", function () {
      ui["layout-field"].hidden = ui.format.value !== "xlsx";
      void saveSettings();
    });
    ui.layout.addEventListener("change", saveSettings);
    ui["opt-indexes"].addEventListener("change", saveSettings);
    ui["opt-foreign-keys"].addEventListener("change", saveSettings);
    ui["load-tables"].addEventListener("click", loadTables);
    ui.export.addEventListener("click", runExport);
    ui["reload-database"].addEventListener("click", loadDatabaseList);
    ui["reload-schema"].addEventListener("click", loadSchemaList);
    ui.filter.addEventListener("input", function () {
      state.filter = ui.filter.value;
      renderTables();
    });
    ui["select-all"].addEventListener("click", function () {
      state.tables.forEach(function (table) {
        state.selected[table.name] = true;
      });
      renderTables();
    });
    ui["select-none"].addEventListener("click", function () {
      state.selected = Object.create(null);
      renderTables();
    });
  }

  /** Context menus hand over a connection summary or a table identity. */
  function applyContext(context) {
    if (!context || typeof context !== "object") return;
    var connectionId = context.connectionId || context.id;
    if (connectionId) {
      for (var i = 0; i < state.connections.length; i += 1) {
        if (state.connections[i].id === connectionId) {
          ui.connection.value = connectionId;
          break;
        }
      }
    }
    if (context.database) ui.database.value = String(context.database);
    if (context.schema) ui.schema.value = String(context.schema);
    syncProfile();
  }

  async function start() {
    cacheElements();
    await dbx.ready;
    lang = String(dbx.locale || "zh-CN").toLowerCase().indexOf("zh") === 0 ? "zh" : "en";
    applyLanguage();

    var capabilities = dbx.capabilities || {};
    if (!capabilities.dataApi) {
      blockWith(m("noDataApi"));
      return;
    }
    if (!capabilities.connectionsApi) {
      blockWith(m("noConnectionsApi"));
      return;
    }

    bindEvents();

    try {
      var all = await dbx.getConnections();
      state.connections = (all || []).filter(function (connection) {
        return (connection.connected || connection.isConnected) && isSqlConnection(connection);
      });
    } catch (error) {
      blockWith(window.DictCollect.errorText(error));
      return;
    }

    if (!state.connections.length) {
      blockWith(m("noConnections"));
      return;
    }

    renderConnections();

    var saved = await restoreSettings();
    if (saved && typeof saved === "object") {
      if (saved.connectionId) {
        for (var i = 0; i < state.connections.length; i += 1) {
          if (state.connections[i].id === saved.connectionId) {
            ui.connection.value = saved.connectionId;
            if (saved.database) ui.database.value = String(saved.database);
            if (saved.schema) ui.schema.value = String(saved.schema);
            break;
          }
        }
      }
      if (saved.format === "docx" || saved.format === "xlsx") ui.format.value = saved.format;
      if (saved.layout === "single-sheet" || saved.layout === "sheet-per-table") ui.layout.value = saved.layout;
      if (typeof saved.includeIndexes === "boolean") ui["opt-indexes"].checked = saved.includeIndexes;
      if (typeof saved.includeForeignKeys === "boolean") ui["opt-foreign-keys"].checked = saved.includeForeignKeys;
    }

    ui["layout-field"].hidden = ui.format.value !== "xlsx";
    syncProfile();
    applyContext(dbx.context);

    // Reopening the same workbench refreshes the context instead of reloading.
    dbx.onContext(function (next) {
      applyContext(next);
    });

    window.addEventListener("dbx-plugin-env", function () {
      var nextLang = String(dbx.locale || "zh-CN").toLowerCase().indexOf("zh") === 0 ? "zh" : "en";
      if (nextLang !== lang) {
        lang = nextLang;
        applyLanguage();
        renderConnections();
        syncProfile();
        renderTables();
      }
    });
  }

  start().catch(function (error) {
    cacheElements();
    blockWith(window.DictCollect ? window.DictCollect.errorText(error) : String(error));
  });
})();
