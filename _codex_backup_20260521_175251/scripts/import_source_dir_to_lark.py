#!/usr/bin/env python3
"""Export a cleaned Excel source directory for Feishu Base import."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_reader import clean_excel  # noqa: E402
import config  # noqa: E402
from catalog_linker import CatalogLinkError, attach_catalog_links  # noqa: E402


DEFAULT_BASE_TOKEN = config.BASE_TOKEN
LARK_CLI = Path(config.LARK_CLI)
BATCH_SIZE = 200
FIELD_ORDER = config.IMPORT_FIELD_ORDER
MAX_BATCH_RETRIES = 5


def iter_excel_files(source_dir: Path):
    converted_stems = {
        path.stem
        for path in source_dir.glob("*.xls*")
        if path.suffix.lower() in {".xlsx", ".xlsm"}
    }
    for path in sorted(source_dir.glob("*.xls*")):
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() == ".xls" and path.stem in converted_stems:
            continue
        else:
            yield path


def item_value(item, name):
    return getattr(item, name, "")


def build_remark(item, import_time: str) -> str:
    return (
        f"来源文件:{item.source_file} | "
        f"Sheet:{item.source_sheet} | "
        f"Excel行:{item.source_row} | "
        f"导入时间:{import_time}"
    )


def clean_records(source_dir: Path):
    import_time = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    records = []
    issues = []

    for path in iter_excel_files(source_dir):
        items, anomalies = clean_excel(str(path))
        for anomaly in anomalies:
            issues.append({
                "文件": path.name,
                "类型": anomaly.异常类型,
                "Sheet": anomaly.source_sheet,
                "行": anomaly.source_row,
                "项目编号": anomaly.项目编号,
                "项目名称": anomaly.项目名称,
                "描述": anomaly.异常描述,
            })

        for item in items:
            record = {
                "备注": build_remark(item, import_time),
                "合同名称": item.合同名称,
                "页签": item.页签,
                "层级路径": item.层级路径,
                "项目编号": item.项目编号,
                "项目名称": item.项目名称,
                "项目特征描述": item.项目特征描述,
                "计量单位": item.计量单位,
                "工程数量": item.工程数量,
                "不含税单价": item.不含税单价,
                "汇总合价": item.汇总合价,
                "定价模式": item.定价信息,
                "_source_file": item.source_file,
                "_source_sheet": item.source_sheet,
                "_source_row": item.source_row,
            }
            records.append(record)

            required = ["层级路径", "计量单位", "工程数量", "不含税单价"]
            for field in required:
                value = record[field]
                if value is None or str(value).strip() == "":
                    issues.append({
                        "文件": item.source_file,
                        "类型": f"{field}为空",
                        "Sheet": item.source_sheet,
                        "行": item.source_row,
                        "项目编号": item.项目编号,
                        "项目名称": item.项目名称,
                        "描述": "导入前硬校验拦截",
                    })

    return records, issues


def write_csv(path: Path, records: list[dict], extra_fields: list[str] | None = None):
    if not records:
        return
    fields = FIELD_ORDER + (extra_fields or []) + ["_source_file", "_source_sheet", "_source_row"]
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(records)


def write_issues(path: Path, issues: list[dict]):
    fields = ["文件", "类型", "Sheet", "行", "项目编号", "项目名称", "描述"]
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(issues)


def run_lark(payload_path: Path, base_token: str, table_id: str, dry_run: bool):
    relative_payload = payload_path.resolve().relative_to(ROOT.resolve())
    cmd = [
        str(LARK_CLI),
        "base",
        "+record-batch-create",
        "--base-token",
        base_token,
        "--table-id",
        table_id,
        "--json",
        f"@{relative_payload}",
        "--as",
        "user",
    ]
    if dry_run:
        cmd.append("--dry-run")
    return subprocess.run(
        cmd,
        cwd=str(ROOT),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=120,
    )


def import_records(
    records: list[dict],
    run_dir: Path,
    base_token: str,
    table_id: str,
    dry_run: bool,
    field_order: list[str],
):
    total = len(records)
    written = 0
    for start in range(0, total, BATCH_SIZE):
        chunk = records[start:start + BATCH_SIZE]
        rows = [[record.get(field, "") for field in field_order] for record in chunk]
        payload = {"fields": field_order, "rows": rows}
        payload_path = run_dir / f"batch_{start // BATCH_SIZE + 1:04d}.json"
        payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

        log_path = run_dir / f"batch_{start // BATCH_SIZE + 1:04d}.log"
        batch_no = start // BATCH_SIZE + 1
        for attempt in range(1, MAX_BATCH_RETRIES + 1):
            result = run_lark(payload_path, base_token, table_id, dry_run)
            output = result.stdout + "\n" + result.stderr
            log_path.write_text(output, encoding="utf-8")
            if result.returncode == 0:
                break
            if "limited" not in output and "800004135" not in output:
                raise RuntimeError(f"batch {batch_no} failed: {result.stderr or result.stdout}")
            if attempt == MAX_BATCH_RETRIES:
                raise RuntimeError(f"batch {batch_no} failed after retries: {result.stderr or result.stdout}")
            wait_seconds = 3 * attempt
            print(f"batch {batch_no} rate limited, retrying in {wait_seconds}s ({attempt}/{MAX_BATCH_RETRIES})")
            time.sleep(wait_seconds)
        written += len(chunk)
        print(f"imported {written}/{total}")
        time.sleep(1.0 if not dry_run else 0.2)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_dir", type=Path)
    parser.add_argument("--base-token", default=DEFAULT_BASE_TOKEN)
    parser.add_argument("--table-id", default="local", help="目标飞书表 ID；仅生成本地文件时可省略")
    parser.add_argument("--dry-run", action="store_true", help="调用 lark-cli dry-run，默认不调用")
    parser.add_argument("--push", action="store_true", help="正式调用 lark-cli 写入飞书，慎用")
    args = parser.parse_args()

    if args.push and args.dry_run:
        parser.error("--push 和 --dry-run 只能选择一个")
    if args.push and args.table_id == "local":
        parser.error("--push 必须显式指定 --table-id")
    if args.dry_run and args.table_id == "local":
        parser.error("--dry-run 必须显式指定 --table-id")

    run_dir = ROOT / "runs" / (
        f"fitout_import_{dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{args.table_id}"
    )
    run_dir.mkdir(parents=True, exist_ok=True)

    records, issues = clean_records(args.source_dir)
    field_order = list(FIELD_ORDER)
    extra_csv_fields: list[str] = []
    catalog_summary = {"enabled": False, "reason": "local-only run"}

    if issues:
        write_csv(run_dir / "cleaned_records.csv", records)
        write_issues(run_dir / "blocked_issues.csv", issues)
        print(json.dumps({
            "ok": False,
            "records": len(records),
            "issues": len(issues),
            "issues_path": str(run_dir / "blocked_issues.csv"),
        }, ensure_ascii=False, indent=2))
        raise SystemExit(1)

    if not records:
        raise SystemExit("no records to import")

    if args.push or args.dry_run:
        try:
            records, catalog_field_id, catalog_summary = attach_catalog_links(
                records,
                table_id=args.table_id,
                lark_cli=LARK_CLI,
                base_token=args.base_token,
                cwd=ROOT,
                source_name_field="合同名称",
                remark_field="备注",
                strict=True,
            )
        except CatalogLinkError as exc:
            failed_path = run_dir / "catalog_link_error.json"
            failed_path.write_text(
                json.dumps({"error": str(exc)}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            raise SystemExit(f"catalog link failed; see {failed_path}") from exc
        if catalog_field_id and catalog_field_id not in field_order:
            field_order.append(catalog_field_id)
            extra_csv_fields.append(catalog_field_id)

    write_csv(run_dir / "cleaned_records.csv", records, extra_csv_fields=extra_csv_fields)
    write_issues(run_dir / "blocked_issues.csv", issues)

    if args.push or args.dry_run:
        import_records(records, run_dir, args.base_token, args.table_id, args.dry_run, field_order)

    print(json.dumps({
        "ok": True,
        "local_only": not (args.push or args.dry_run),
        "dry_run": args.dry_run,
        "push": args.push,
        "table_id": args.table_id,
        "records": len(records),
        "run_dir": str(run_dir),
        "cleaned_records_path": str(run_dir / "cleaned_records.csv"),
        "blocked_issues_path": str(run_dir / "blocked_issues.csv"),
        "field_order": field_order,
        "catalog_link": catalog_summary,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
