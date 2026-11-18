# 数据字典导出（DBX 插件）

把 DBX 中一个已打开连接的库表结构导出为数据字典文档，支持 **DOCX** 和 **Excel（.xlsx）** 两种格式。

纯前端插件，没有原生 Sidecar：元数据通过 DBX 的只读数据接口读取，文档在插件沙箱内生成，再交给宿主的原生保存对话框落盘。插件自身不带驱动、不接触凭据、没有网络访问。

## 功能

- 选择连接 + 数据库/Schema，读取表与视图清单，按名称或注释筛选后勾选导出范围。
- 导出内容：文档信息页、表清单总览、每张表的字段明细（序号、字段名、类型、长度/精度、可空、主键、默认值、注释），可选索引与外键。
- DOCX：一页一张表，表头跨页重复，中文使用 `eastAsia` 字体声明。
- Excel 两种布局：
  - 汇总表：`总览` + `字段明细` + `索引` + `外键`，首行冻结并带筛选器；
  - 每表一个工作表：表名自动按 Excel 规则截断去重（31 字符、去掉 `[]:*?/\`）。
- 上次选择的连接、库、格式和选项记在插件私有存储里，下次打开自动回填。

## 支持的数据库

| Profile | 覆盖的 `dbType` | 注释 | 索引 | 外键 |
| --- | --- | --- | --- | --- |
| `mysql` | mysql, mariadb, tidb, oceanbase, doris, starrocks, goldendb, gbase | ✓ | ✓ | ✓ |
| `postgres` | postgres, redshift, gaussdb, opengauss, kingbase, highgo, vastbase, uxdb, gbase8c, greenplum, cockroachdb | ✓ | ✓ | ✓ |
| `sqlserver` | sqlserver, mssql | ✓ | ✓ | ✓ |
| `oracle` | oracle, oceanbase-oracle, dameng | ✓ | ✓ | ✓ |
| `sqlite` | sqlite, turso, rqlite, cloudflare-d1 | — | ✓ | ✓ |
| `ansi` | duckdb, clickhouse, databend, trino, presto, hive, spark, kyuubi, impala, snowflake, vertica, exasol, h2, db2, saphana, teradata, yashandb, bigquery | — | — | ✓ |

`ansi` 是 `information_schema` 兜底：标准视图里没有注释和索引元数据，这两项会留空并在文档的「提示」里写明，而不是编造。Redis、MongoDB 等非 SQL 连接不在支持范围内，宿主也不会为它们提供只读查询能力。

## 权限与边界

```json
"permissions": ["host.connections:read", "host.data:read", "host.storage"]
```

- `host.connections:read`：列出已配置连接（脱敏，不含主机、端口、用户名、密码）。
- `host.data:read`：每次只提交**一条**被 DBX 判定为只读的 SQL。首次查询某个连接时，DBX 会弹窗征求用户同意；授权可在「插件中心 → 已安装」随时撤销。
- `host.storage`：记住上一次的导出选项。

保存文件走 `dbxPlugin.saveFile`，由宿主弹出原生保存对话框，不需要额外权限；宿主只提供分块通道时自动回退到 `fileTransfer.beginSave`。

几条硬性限制来自宿主，插件按它们设计而不是绕过：

- 连接必须**已经打开**。已保存但断开的连接会被拒绝（`Connection is not open`），插件不会替用户重连。
- 不能用 `USE db` 切库，数据库通过请求字段传递。
- 单次查询最多 5000 行，所以每类元数据都按 2000 行一页翻页，直到读到不满一页为止。
- 目录名（库名、Schema 名）进入 SQL 前先过白名单校验：含引号、反斜杠、分号、注释符或控制字符的名字直接报错，不做转义拼接。

## 开发

```bash
npm install --global @dbx-app/plugin-cli

dbx-plugin dev --path . --port 5190   # 本地开发宿主
dbx-plugin package .                  # 生成 dist/<id>-<version>-universal.dbxp
```

`dbx-plugin dev` 的开发宿主只模拟了 Host API 的一个子集，不提供 `queryData` / `saveFile`，页面会直接显示「宿主不支持」并停下。完整流程需要在 DBX 里安装未签名开发包验证：插件中心 → 设置里打开「允许安装未签名开发包」。

## 身份与上架

```
id:        io.github.wangyd.data-dictionary
publisher: wangyd
license:   Apache-2.0
source:    https://github.com/yaoduan/dbx-plugin-data-dictionary
```

`id` 和 `publisher` 首次上架后不可更改。

上架流程见[官方文档](https://dbxio.com/cn/docs/plugin-development)和 [`dbx-store/CONTRIBUTING.md`](https://github.com/t8y2/dbx-store/blob/main/CONTRIBUTING.md)：

1. 本仓库推到 GitHub，打版本 Tag 并发 Release——`.github/workflows/plugin-release.yml` 会自动构建未签名候选包，产出 `.dbxp`、`.artifact.json` 和 `release-candidates.json` 并挂到 Release 上。
2. Fork `t8y2/dbx-store`，向 `main` 提**一个** PR，包含 `publishers/wangyd.json`（仅首次）和 `candidates/io.github.wangyd.data-dictionary.json`。
3. CI 会故意停在红色的 `open candidate(s) awaiting DBX Store signing`，这是设计如此——它阻止未签名内容被合并。维护者审核后运行受保护的签名工作流，把最终目录文件写回同一个 PR，CI 转绿后再合并。

作者只发布未签名候选包，不持有也不填写官方签名密钥（`signingKeyId` 由 DBX Store 的签名流程写入）。已发布版本的字节不可覆盖，任何改动都要递增版本号重新送审。

## 代码结构

```
manifest.json          身份、权限、贡献点（工作台 / 连接与表右键菜单 / 命令 + 工具栏）
ui/index.html          工作台页面
ui/main.js             连接选择、表清单、导出流程、宿主能力探测
ui/lib/dialects.js     各方言的目录查询 SQL 与分页写法
ui/lib/collect.js      翻页查询、行列映射、聚合成字典模型
ui/lib/docx.js         WordprocessingML 写出
ui/lib/xlsx.js         SpreadsheetML 写出
ui/lib/report.js       字典模型 → 文档块 / 工作表
ui/lib/zip.js          ZIP 容器（deflate-raw，缺失时退化为 store）
```

OOXML 和 ZIP 都是手写的：插件沙箱没有网络也没有打包器（CSP 的 `connect-src` 默认为 `none`，CDN 取不到），因此不引入 `exceljs` / `docx` 这类依赖。
