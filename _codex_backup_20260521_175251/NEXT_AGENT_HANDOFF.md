# 工程量清单清洗与飞书导入 - 交接说明

## 当前目标

工具链只维护工程量清单主表数据：主清单字段、层级路径、定价模式。不再维护综合单价分析费用明细。

当前可用工具链目录：

```text
C:\Users\56237\Documents\project-hysw\_codex_backup_20260521_175251
```

正式链路不依赖调试副本；`excel_reader_debug.py` 已清理。

## 输出字段

| 字段 | 说明 |
|---|---|
| 备注 | 来源或导入时间 |
| 合同名称 | 文件名提取或导入参数 |
| 页签 | Excel 实体清单 sheet 的专业名称 |
| 层级路径 | `科目/二级页签` 合并字段，使用 `/` 分层；无显式层级时用 `页签` 兜底 |
| 项目编号 | 清单编号 |
| 项目名称 | 清单名称 |
| 项目特征描述 | 清单特征 |
| 计量单位 | 单位 |
| 工程数量 | 数量 |
| 不含税单价 | 单价 |
| 汇总合价 | 合价 |
| 定价模式 | 自主报价或基准价/浮率信息 |

## 层级识别

- `科目` 行只更新科目路径，不直接输出。
- `二级页签` 行切换当前二级标题。
- `三级页签` 行追加到当前二级标题之后。
- 无编号、无单位、无单价但有名称的行，按同级标题处理，切换当前二级标题。
- 隐藏行正常读取，不因 Excel 行隐藏状态跳过。
- 主清单 sheet 兼容 `实体工程量清单`、`实体工程清单`、`实体清单`、`工程量清单` 等命名。
- 表头兼容 `项目编号`、`项目编码`、`序号`；B 型价格兼容 `基准价`、`基准单价`、两层投标价表头。
- 空值导入前必须清零：`不含税单价`、`工程数量`、`计量单位`、`层级路径` 不允许为空。
- 可自动修复：分楼栋/分区域工程量求和、供应综合单价+安装综合单价求和、同文件同名项目唯一单位回填。
- 可自动跳过：非清单项、无单位且数量/单价/合价均为 0 的占位行。

典型修复：`恢复对景观赏水景`、`恢复休憩空间`、`恢复归家叠水水景`、`恢复3B架空层地面`、`展示区现状绿化...` 这类空类别标题互为同级，不应合并成 `A/B/C/D/E`。

重要实现约束：`auto_split_table.py`、`scripts/full_reimport.py` 和 `scripts/repair_csv_layer_path.py` 均应复用 `excel_reader.clean_excel()` 的结果，不要再维护第二套 pandas 层级解析规则。

## 保留能力

- Excel 读取和合并单元格处理
- 隐藏行读取
- 表头识别和列映射
- A/B 型定价模式识别
- B 型浮率验证
- 缺失单价检查
- 飞书导入、分表、dry-run、全量重导入主流程
- CSV 层级路径修复脚本：`scripts/repair_csv_layer_path.py`
- 批量源目录扫描脚本：`scripts/scan_source_dir.py`
- 批量源目录导入脚本：`scripts/import_source_dir_to_lark.py`
- 通用固定字段顺序配置：`config.IMPORT_FIELD_ORDER`
- 视图列顺序示例：`scripts/fitout_view_fields.json`
- 单合同 `--dry-run` 只做本地解析和检查，不访问飞书、不写入
- Windows PowerShell 兼容的纯文本命令行输出

## 已清理内容

- 综合单价分析结构校验
- 费用明细字段抽取
- 费用字段映射
- 分析明细回填流程

## 主要脚本

| 脚本 | 用途 |
|---|---|
| `excel_reader.py` | 主清单清洗核心，适合批量读取 Excel |
| `auto_split_table.py` | 单合同解析、检查、导入飞书；解析阶段复用 `excel_reader.clean_excel()` |
| `scripts/full_reimport.py` | 主清单全量重导入 |
| `scripts/repair_csv_layer_path.py` | 用源 Excel 重新生成 CSV 的 `层级路径` |
| `scripts/scan_source_dir.py` | 批量扫描源 Excel 目录，生成本地质量报告，不访问飞书 |
| `scripts/import_source_dir_to_lark.py` | 批量导入一个源目录到指定飞书表，导入前硬校验空层级、空单位、空工程量、空单价 |

## 精装一批次已跑通结果

- 住宅源目录：`D:\单价库\精装工程\一批次\住宅清单`
- 住宅目标表：`tblY7o8bNiBxEYGA`
- 住宅已导入：9656 行
- 非住宅源目录：`D:\单价库\精装工程\一批次\非住宅清单`
- 非住宅目标表：`tblRdm1rcN46iGuw`
- 非住宅已导入：13889 行
- 两表字段和视图顺序一致，字段顺序见 `config.IMPORT_FIELD_ORDER`
- 注意：上述表 ID 是本次精装一批次记录，不是工具链默认值。新专业、新批次应重新判断是否拆表，并通过 `--table-id` 显式指定目标表。

导入前质量报告：

- `runs\fitout_residential_scan\scan_summary.csv`
- `runs\fitout_residential_scan\scan_issues.csv`
- `runs\fitout_non_residential_scan\scan_summary.csv`
- `runs\fitout_non_residential_scan\scan_issues.csv`

正式导入归档：

- `runs\fitout_import_20260524_191143`
- `runs\fitout_import_20260524_191453`

## 验证建议

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\auto_split_table.py" "<合同名称>" "<Excel文件路径>" --dry-run

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\full_reimport.py" --dry-run

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\scan_source_dir.py" `
  "D:\单价库\精装工程\一批次" `
  --out-dir ".\runs\fitout_batch_scan"

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\import_source_dir_to_lark.py" `
  "D:\单价库\精装工程\一批次\住宅清单" `
  --table-id tblY7o8bNiBxEYGA `
  --dry-run

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\import_source_dir_to_lark.py" `
  "D:\单价库\精装工程\一批次\非住宅清单" `
  --table-id tblRdm1rcN46iGuw `
  --dry-run

& "C:\Users\56237\AppData\Roaming\npm\lark-cli.cmd" base +view-set-visible-fields `
  --base-token ARDubM91FaL62esCg1hcnHYtnFh `
  --table-id <table_id> `
  --view-id vewwDxH04m `
  --json "@scripts\fitout_view_fields.json" `
  --as user

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py" `
  --output "D:\单价库\深圳公司单价库_景观工程.verify.csv"
```

检查 `repair_csv_layer_path.py` 输出中的 `matched_rows`。未匹配行不会丢失，会回退为旧 CSV 的 `科目/二级页签` 合成路径。
