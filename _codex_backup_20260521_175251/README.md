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
    └── repair_csv_layer_path.py
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

- `页签` 来自实体工程量清单 sheet 名中的专业名称。
- `层级路径` = `科目/二级页签`。
- `类别 = 科目` 的行只更新科目路径，不输出。
- `类别 = 二级页签` 的行切换当前二级标题。
- `类别 = 三级页签` 的行追加到当前二级标题之后。
- 无编号、无单位、无单价但有名称的行，视为同级标题，切换当前二级标题。
- 隐藏行正常读取，不因 Excel 行隐藏状态跳过。

`auto_split_table.py`、`scripts/full_reimport.py` 和 `scripts/repair_csv_layer_path.py` 均复用 `excel_reader.clean_excel()` 的清洗结果，避免多套层级规则不一致。

## 保留能力

- Excel 读取和合并单元格处理
- 隐藏行读取
- 多行表头识别和列映射
- A/B 型定价模式识别
- 基准价、浮率、投标价处理
- 浮率校验
- 缺失单价检查
- 飞书导入、分表、dry-run、全量重导入主流程
- CSV 层级路径修复脚本
- Windows PowerShell 兼容的纯文本命令行输出

## 已清理能力

- 综合单价分析结构校验
- 费用明细字段抽取
- 分析明细回填流程
