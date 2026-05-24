#!/usr/bin/env python3
"""Scan a source Excel directory with the current cleaner, without importing.

Outputs:
- scan_summary.csv
- scan_issues.csv
- scan_items_sample.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_reader import clean_excel  # noqa: E402


ITEM_FIELDS = [
    "source_file",
    "source_sheet",
    "source_row",
    "合同名称",
    "页签",
    "层级路径",
    "项目编号",
    "项目名称",
    "计量单位",
    "工程数量",
    "不含税单价",
    "汇总合价",
    "定价信息",
    "异常标记",
]


def attr(obj, name, default=""):
    return getattr(obj, name, default)


def iter_excel_files(source_dir: Path):
    for path in sorted(source_dir.glob("*.xls*")):
        if not path.name.startswith("~$"):
            yield path


def scan_file(path: Path):
    try:
        items, anomalies = clean_excel(str(path))
        return "ok", items, anomalies, ""
    except SystemExit as exc:
        return "blocked_system_exit", [], [], str(exc)
    except Exception as exc:
        return "error", [], [], repr(exc)


def summarize_pricing(items):
    pricing = Counter()
    for item in items:
        info = attr(item, "定价信息", "")
        if info.startswith("基准价"):
            pricing["基准价"] += 1
        elif info == "自主报价":
            pricing["自主报价"] += 1
        elif info:
            pricing["其他定价"] += 1
        else:
            pricing["空"] += 1
    return pricing


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_dir", type=Path, help="Excel source directory")
    parser.add_argument("--out-dir", type=Path, default=ROOT / "runs" / "source_scan")
    parser.add_argument("--sample-limit", type=int, default=5000)
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    summary_rows = []
    issue_rows = []
    total_counter = Counter()
    sample_written = 0

    summary_path = args.out_dir / "scan_summary.csv"
    issues_path = args.out_dir / "scan_issues.csv"
    sample_path = args.out_dir / "scan_items_sample.csv"

    with sample_path.open("w", encoding="utf-8-sig", newline="") as sample_file:
        sample_writer = csv.DictWriter(sample_file, fieldnames=ITEM_FIELDS)
        sample_writer.writeheader()

        for path in iter_excel_files(args.source_dir):
            status, items, anomalies, error = scan_file(path)

            sheets = Counter(attr(item, "source_sheet") for item in items)
            pricing = summarize_pricing(items)
            missing_price = [item for item in items if attr(item, "不含税单价") is None]
            missing_qty = [item for item in items if attr(item, "工程数量") is None]
            empty_layer = [item for item in items if not str(attr(item, "层级路径")).strip()]
            empty_unit = [item for item in items if not str(attr(item, "计量单位")).strip()]

            row = {
                "文件": path.name,
                "状态": status,
                "清单行数": len(items),
                "sheet数": len(sheets),
                "异常数": len(anomalies),
                "缺单价": len(missing_price),
                "缺工程量": len(missing_qty),
                "空层级路径": len(empty_layer),
                "空单位": len(empty_unit),
                "自主报价": pricing["自主报价"],
                "基准价": pricing["基准价"],
                "其他定价": pricing["其他定价"] + pricing["空"],
                "错误": error,
                "sheets": "; ".join(f"{name}:{count}" for name, count in sheets.items()),
            }
            summary_rows.append(row)

            for key in ["清单行数", "异常数", "缺单价", "缺工程量", "空层级路径", "空单位"]:
                total_counter[key] += int(row[key])

            for anomaly in anomalies:
                issue_rows.append({
                    "文件": path.name,
                    "类型": attr(anomaly, "异常类型"),
                    "Sheet": attr(anomaly, "source_sheet"),
                    "行": attr(anomaly, "source_row"),
                    "项目编号": attr(anomaly, "项目编号"),
                    "项目名称": attr(anomaly, "项目名称"),
                    "描述": attr(anomaly, "异常描述"),
                })

            for label, bad_items in [
                ("缺单价", missing_price[:20]),
                ("缺工程量", missing_qty[:20]),
                ("空层级路径", empty_layer[:20]),
                ("空单位", empty_unit[:20]),
            ]:
                for item in bad_items:
                    issue_rows.append({
                        "文件": path.name,
                        "类型": label,
                        "Sheet": attr(item, "source_sheet"),
                        "行": attr(item, "source_row"),
                        "项目编号": attr(item, "项目编号"),
                        "项目名称": attr(item, "项目名称"),
                        "描述": attr(item, "异常标记"),
                    })

            for item in items:
                if sample_written >= args.sample_limit:
                    break
                sample_writer.writerow({field: attr(item, field) for field in ITEM_FIELDS})
                sample_written += 1

    summary_fields = [
        "文件", "状态", "清单行数", "sheet数", "异常数", "缺单价", "缺工程量",
        "空层级路径", "空单位", "自主报价", "基准价", "其他定价", "错误", "sheets",
    ]
    with summary_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=summary_fields)
        writer.writeheader()
        writer.writerows(summary_rows)

    issue_fields = ["文件", "类型", "Sheet", "行", "项目编号", "项目名称", "描述"]
    with issues_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=issue_fields)
        writer.writeheader()
        writer.writerows(issue_rows)

    result = {
        "source_dir": str(args.source_dir),
        "files": len(summary_rows),
        "ok_files": sum(1 for row in summary_rows if row["状态"] == "ok"),
        "blocked_or_error_files": sum(1 for row in summary_rows if row["状态"] != "ok"),
        "totals": dict(total_counter),
        "summary_path": str(summary_path),
        "issues_path": str(issues_path),
        "sample_path": str(sample_path),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
