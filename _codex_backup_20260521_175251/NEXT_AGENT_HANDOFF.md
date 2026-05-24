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
| 层级路径 | `科目/二级页签` 合并字段，使用 `/` 分层 |
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

## 验证建议

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\auto_split_table.py" "<合同名称>" "<Excel文件路径>" --dry-run

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\full_reimport.py" --dry-run

& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py" `
  --output "D:\单价库\深圳公司单价库_景观工程.verify.csv"
```

检查 `repair_csv_layer_path.py` 输出中的 `matched_rows`。未匹配行不会丢失，会回退为旧 CSV 的 `科目/二级页签` 合成路径。
