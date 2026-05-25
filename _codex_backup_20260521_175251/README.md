# 工程量清单清洗与飞书导入工具链

用途：清洗工程量清单 Excel，提取主清单字段、层级路径和定价模式，并通过 `lark-cli` 导入飞书多维表格。

## 当前字段

- 备注
- 合同名称
- 页签
- 层级路径
- 项目编号
- 项目名称
- 项目特征描述
- 计量单位
- 工程数量
- 不含税单价
- 汇总合价
- 定价模式

`层级路径` 由原 `科目` 与 `二级页签` 通过 `/` 合并。综合单价分析费用明细字段及对应脚本流程已清理。

## 文件结构

```text
工具链/
├── excel_reader.py
├── auto_split_table.py
├── config.py
├── import_rules.yaml
├── README.md
├── NEXT_AGENT_HANDOFF.md
└── scripts/
    ├── full_reimport.py
    ├── import_source_dir_to_lark.py
    ├── repair_csv_layer_path.py
    ├── fitout_view_fields.json
    └── scan_source_dir.py
```

当前可用工具链目录：

```text
C:\Users\56237\Documents\project-hysw\_codex_backup_20260521_175251
```

正式链路不依赖调试副本；`excel_reader_debug.py` 等临时文件已清理。

## 基本用法

单合同 dry-run：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\auto_split_table.py" "<合同名称>" "<Excel文件路径>" --dry-run
```

单合同 dry-run 只做本地解析、层级识别、单价检查和归档，不访问飞书、不查表空间、不写入。

单合同导入：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\auto_split_table.py" "<合同名称>" "<Excel文件路径>"
```

全量重导入 dry-run：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\full_reimport.py" --dry-run
```

全量重导入正式执行：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\full_reimport.py" --confirm
```

## 批量扫描

新批次导入前，先用 `scripts\scan_source_dir.py` 做本地扫描，不访问飞书、不创建字段、不导入：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\scan_source_dir.py" `
  "D:\单价库\精装工程\一批次" `
  --out-dir ".\runs\fitout_batch_scan"
```

输出文件：

- `scan_summary.csv`：按文件汇总清单行数、异常数、缺单价、缺工程量、空层级路径、空单位。
- `scan_issues.csv`：抽样列出需人工确认的问题行。
- `scan_items_sample.csv`：抽样导出清洗后记录，用于核查字段落列。

## 批量分表导入

新批次按需要分好文件夹后，先扫描再导入。是否拆表、拆几张表、每个目录导入哪张表，应按本次专业、数据量和飞书单表上限判断；工具链不固定住宅/非住宅表 ID。

2026-05-24 精装一批次已验证目标，仅作为本次记录：

- 住宅清单：`tblY7o8bNiBxEYGA`
- 非住宅清单：`tblRdm1rcN46iGuw`
- Base：`ARDubM91FaL62esCg1hcnHYtnFh`
- 单表上限：20000 行；脚本按 200 行一批写入。

推荐目录结构：

```text
D:\单价库\精装工程\一批次\
├── 住宅清单\
└── 非住宅清单\
```

扫描住宅：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\scan_source_dir.py" `
  "D:\单价库\精装工程\一批次\住宅清单" `
  --out-dir ".\runs\fitout_residential_scan"
```

扫描非住宅：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\scan_source_dir.py" `
  "D:\单价库\精装工程\一批次\非住宅清单" `
  --out-dir ".\runs\fitout_non_residential_scan"
```

生成本地待复制 CSV，发现 `层级路径`、`计量单位`、`工程数量`、`不含税单价` 为空会直接拦截：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\import_source_dir_to_lark.py" `
  "D:\单价库\精装工程\一批次\住宅清单" `
  --table-id tblY7o8bNiBxEYGA

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\import_source_dir_to_lark.py" `
  "D:\单价库\精装工程\一批次\非住宅清单" `
  --table-id tblRdm1rcN46iGuw
```

默认只生成本地文件，不调用飞书。脚本会在 `runs\fitout_import_YYYYMMDD_HHMMSS_microseconds_<table_id>\` 保存 `cleaned_records.csv` 和 `blocked_issues.csv`。如需调用飞书检查请求体，可加 `--dry-run`；如需正式写入，必须显式加 `--push`。

导出的 `合同名称` 会清理开头列表序号和残留标点；`页签` 会保留成对括号，例如 `公区（包干）`。

第一列 `备注` 会写入来源追溯信息：

```text
来源文件:<文件名> | Sheet:<页签名> | Excel行:<行号> | 导入时间:<时间>
```

固定飞书视图列顺序时，需要先确认目标表字段 ID 与 `scripts\fitout_view_fields.json` 一致；不一致时应按新表字段 ID 重新生成 JSON：

```powershell
& "C:\Users\56237\AppData\Roaming\npm\lark-cli.cmd" base +view-set-visible-fields `
  --base-token ARDubM91FaL62esCg1hcnHYtnFh `
  --table-id <table_id> `
  --view-id vewwDxH04m `
  --json "@scripts\fitout_view_fields.json" `
  --as user
```

固定导入字段顺序为：备注、合同名称、页签、层级路径、项目编号、项目名称、项目特征描述、计量单位、工程数量、不含税单价、汇总合价、定价模式。目标表 ID 必须每次通过 `--table-id` 显式指定。

## CSV 层级修复

默认修复景观工程 CSV：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py"
```

指定新 CSV：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py" `
  --input "D:\单价库\深圳公司单价库_景观工程_新.csv" `
  --output "D:\单价库\深圳公司单价库_景观工程_新.fixed.csv"
```

增加新的 Excel 源目录：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py" `
  --source-dir "D:\单价库\景观工程" `
  --source-dir "D:\单价库\景观工程\第二批" `
  --source-dir "D:\单价库\景观工程\第三批"
```

## 层级规则

- `页签` 来自主清单 sheet 名中的专业名称。
- `层级路径` = `科目/二级页签`；如果原表没有科目或二级页签行，则用 `页签` 兜底，避免空层级。
- `类别 = 科目` 的行只更新科目路径，不输出。
- `类别 = 二级页签` 的行切换当前二级标题。
- `类别 = 三级页签` 的行追加到当前二级标题之后。
- 无编号、无单位、无单价但有名称的行，视为同级标题，切换当前二级标题。
- 隐藏行正常读取，不因 Excel 行隐藏状态跳过。
- 隐藏 sheet / veryHidden sheet 默认跳过，不参与扫描、导出或导入。
- 主清单 sheet 兼容 `实体工程量清单`、`实体工程清单`、`实体清单`、`工程量清单` 等命名。
- 主清单 sheet 会跳过 `综合单价分析`、`单价分析`、`单价分析表` 等分析页。
- 表头兼容 `项目编号`、`项目编码`、`序号`；项目名称列兼容误写为 `项` 的四行表头；B 型价格兼容 `基准价`、`基准单价`、两层投标价表头。
- 导入前扫描要求 `不含税单价`、`工程数量`、`计量单位`、`层级路径` 不为空。
- 分楼栋/分区域工程量存在但汇总工程量为空时，自动汇总分项工程量。
- 总综合单价为空但同组供应综合单价、安装综合单价存在时，自动相加补足。
- 同文件同名项目只有一个明确单位时，用该单位补齐偶发漏填单位。
- `马桶（甲供）` 源表漏填单位时按 `套` 补齐。
- 非清单项、无单位且数量/单价/合价均为 0 的占位行不输出。
- 误标为 `清单项` 的层级/占位行会跳过：`单位=项` 且数量/单价/合价为空；或无单位、数量为空/0、单价为空/0、合价为空/0。
- 批量扫描/导出时，如果同目录存在同名 `.xls` 和已转档的 `.xlsx`/`.xlsm`，默认跳过旧 `.xls`。

`auto_split_table.py`、`scripts/full_reimport.py` 和 `scripts/repair_csv_layer_path.py` 均复用 `excel_reader.clean_excel()` 的清洗结果，避免多套层级规则不一致。

## 保留能力

- Excel 读取和合并单元格处理
- 隐藏行读取
- 隐藏 sheet 跳过
- 多行表头识别和列映射
- A/B 型定价模式识别
- 基准价、浮率、投标价处理
- 浮率校验
- 缺失单价检查
- 飞书导入、分表、dry-run、全量重导入主流程
- CSV 层级路径修复脚本
- 批量源目录扫描脚本：`scripts/scan_source_dir.py`
- 批量源目录导入脚本：`scripts/import_source_dir_to_lark.py`
- 通用固定字段顺序配置：`config.IMPORT_FIELD_ORDER`
- 视图列顺序示例：`scripts/fitout_view_fields.json`
- Windows PowerShell 兼容的纯文本命令行输出

## 已清理能力

- 综合单价分析结构校验
- 费用明细字段抽取
- 分析明细回填流程

## 精装二批次复用记录

- 住宅目录：`D:\单价库\精装工程\二批次\住宅清单`，导入 `tblY7o8bNiBxEYGA`，4725 行。
- 非住宅目录：`D:\单价库\精装工程\二批次\非住宅清单`，导入 `tblRdm1rcN46iGuw`，5734 行。
- 扫描清零报告：`runs\fitout_batch2_residential_scan_fixed4\`、`runs\fitout_batch2_non_residential_scan_fixed4\`。
- 正式导入归档：`runs\fitout_import_20260525_153053_895417_tblY7o8bNiBxEYGA\`、`runs\fitout_import_20260525_153246_373029_tblRdm1rcN46iGuw\`。
- 二批次导入后边界校验：住宅总行数 14381，非住宅总行数 19623。

## 清单目录关联

导入时会通过 `catalog_linker.py` 自动读取目录表 `tblobzr4LplrbiYT`，用来源文件名/合同名称匹配目录记录，并把目标表的 `清单目录` 关联字段一并写入。当前已配置：

- `tblY7o8bNiBxEYGA` -> `fldPFg6p1L`
- `tbleZCShQmE1aRhQ` -> `fld6BcTVBu`
- `tblRdm1rcN46iGuw` -> `fldEoK947G`

如果后续新增目标表，需要先在表里增加链接到 `tblobzr4LplrbiYT` 的 `清单目录` 字段，再把表 ID 和 field ID 加到 `catalog_linker.CATALOG_LINK_FIELD_BY_TABLE`。
