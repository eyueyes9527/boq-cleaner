# CSV 层级路径修复脚本使用说明

脚本位置：

```text
scripts/repair_csv_layer_path.py
```

作用：

1. 读取当前工具链的 `excel_reader.py`
2. 扫描源 Excel 文件
3. 用 `合同名称 + 页签 + 项目编号 + 项目名称 + 项目特征描述前60字` 建立层级索引
4. 读取旧 CSV
5. 输出只保留一列 `层级路径` 的新 CSV

## 默认运行

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py"
```

默认读取：

```text
D:\单价库\深圳公司单价库_景观工程.csv
D:\单价库\景观工程
D:\单价库\景观工程\第二批
```

默认输出：

```text
D:\单价库\深圳公司单价库_景观工程.fixed.csv
```

## 指定新 CSV

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py" `
  --input "D:\单价库\深圳公司单价库_景观工程_新.csv" `
  --output "D:\单价库\深圳公司单价库_景观工程_新.fixed.csv"
```

## 指定新的源 Excel 目录

只要传了 `--source-dir`，就要把需要参与匹配的目录全部写上：

```powershell
& "C:\Users\56237\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\scripts\repair_csv_layer_path.py" `
  --source-dir "D:\单价库\景观工程" `
  --source-dir "D:\单价库\景观工程\第二批" `
  --source-dir "D:\单价库\景观工程\第三批"
```

## 输出结果怎么看

脚本结束时会输出：

```text
source_excel_files=16
source_items=9019
index_keys=9019
csv_rows=6855
matched_rows=6445
output=D:\单价库\xxx.fixed.csv
```

重点看：

- `csv_rows`：CSV 总行数
- `matched_rows`：从源 Excel 精准匹配到层级路径的行数
- `output`：生成文件位置

未匹配的行不会丢失，会使用旧 CSV 的 `科目/二级页签` 合成 `层级路径`。
