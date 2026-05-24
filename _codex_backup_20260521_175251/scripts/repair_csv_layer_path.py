#!/usr/bin/env python3
"""Regenerate CSV hierarchy as a single layer path column.

This script re-reads source Excel files with the fixed toolchain reader,
builds a row-level hierarchy index, then rewrites an exported CSV with:

    科目 + 二级页签 -> 层级路径

Default paths match the current landscape-unit-price workflow. Override them
with CLI arguments when reusing this for another discipline or folder.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path


DEFAULT_TOOLCHAIN_DIR = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_CSV = Path(r"D:\单价库\深圳公司单价库_景观工程.csv")
DEFAULT_OUTPUT_CSV = Path(r"D:\单价库\深圳公司单价库_景观工程.fixed.csv")
DEFAULT_SOURCE_DIRS = [
    Path(r"D:\单价库\景观工程"),
    Path(r"D:\单价库\景观工程\第二批"),
]

OUTPUT_FIELDS = [
    "备注",
    "合同名称",
    "页签",
    "层级路径",
    "项目编号",
    "项目名称",
    "项目特征描述",
    "计量单位",
    "工程数量",
    "不含税单价",
    "汇总合价",
    "定价模式",
]


def join_path(*parts: str) -> str:
    result: list[str] = []
    for part in parts:
        if not part:
            continue
        for piece in str(part).replace("|", "/").split("/"):
            piece = piece.strip()
            if piece and (not result or result[-1] != piece):
                result.append(piece)
    return "/".join(result)


def import_reader(toolchain_dir: Path):
    if not (toolchain_dir / "excel_reader.py").exists():
        raise FileNotFoundError(f"excel_reader.py not found: {toolchain_dir}")
    sys.path.insert(0, str(toolchain_dir))
    import excel_reader  # type: ignore

    return excel_reader


def iter_excel_files(source_dirs: list[Path]):
    for source_dir in source_dirs:
        if not source_dir.exists():
            continue
        for path in sorted(source_dir.glob("*.xlsx")):
            if not path.name.startswith("~$"):
                yield path


def make_key(contract_name: str, page: str, code: str, name: str, desc: str):
    return (
        str(contract_name or "").strip(),
        str(page or "").strip(),
        str(code or "").strip(),
        str(name or "").strip(),
        str(desc or "")[:60],
    )


def build_layer_index(toolchain_dir: Path, source_dirs: list[Path]):
    excel_reader = import_reader(toolchain_dir)
    index: dict[tuple[str, str, str, str, str], str] = {}
    files = 0
    items = 0

    for excel_path in iter_excel_files(source_dirs):
        files += 1
        try:
            cleaned_items, _ = excel_reader.clean_excel(str(excel_path))
        except BaseException as exc:
            print(f"skip: {excel_path.name}: {exc}")
            continue

        for item in cleaned_items:
            items += 1
            layer_path = getattr(item, "层级路径", "")
            if not layer_path:
                layer_path = join_path(getattr(item, "科目", ""), getattr(item, "二级页签", ""))

            for contract_name in {getattr(item, "合同名称", ""), excel_path.stem}:
                key = make_key(
                    contract_name,
                    getattr(item, "页签", ""),
                    getattr(item, "项目编号", ""),
                    getattr(item, "项目名称", ""),
                    getattr(item, "项目特征描述", ""),
                )
                index[key] = layer_path

    return index, files, items


def rewrite_csv(input_csv: Path, output_csv: Path, index: dict[tuple[str, str, str, str, str], str]):
    rows = 0
    matched = 0

    with input_csv.open("r", encoding="utf-8-sig", newline="") as fin:
        reader = csv.DictReader(fin)
        output_csv.parent.mkdir(parents=True, exist_ok=True)
        with output_csv.open("w", encoding="utf-8-sig", newline="") as fout:
            writer = csv.DictWriter(fout, fieldnames=OUTPUT_FIELDS)
            writer.writeheader()

            for row in reader:
                rows += 1
                key = make_key(
                    row.get("合同名称", ""),
                    row.get("页签", ""),
                    row.get("项目编号", ""),
                    row.get("项目名称", ""),
                    row.get("项目特征描述", ""),
                )
                layer_path = index.get(key)
                if layer_path:
                    matched += 1
                else:
                    layer_path = join_path(row.get("科目", ""), row.get("二级页签", ""))

                writer.writerow(
                    {
                        "备注": row.get("备注", ""),
                        "合同名称": row.get("合同名称", ""),
                        "页签": row.get("页签", ""),
                        "层级路径": layer_path,
                        "项目编号": row.get("项目编号", ""),
                        "项目名称": row.get("项目名称", ""),
                        "项目特征描述": row.get("项目特征描述", ""),
                        "计量单位": row.get("计量单位", ""),
                        "工程数量": row.get("工程数量", ""),
                        "不含税单价": row.get("不含税单价", ""),
                        "汇总合价": row.get("汇总合价", ""),
                        "定价模式": row.get("定价模式", ""),
                    }
                )

    return rows, matched


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--toolchain-dir", type=Path, default=DEFAULT_TOOLCHAIN_DIR)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_CSV)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument(
        "--source-dir",
        type=Path,
        action="append",
        default=None,
        help="Excel source directory. Can be used more than once.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    source_dirs = args.source_dir or DEFAULT_SOURCE_DIRS

    index, files, items = build_layer_index(args.toolchain_dir, source_dirs)
    rows, matched = rewrite_csv(args.input, args.output, index)

    print(f"source_excel_files={files}")
    print(f"source_items={items}")
    print(f"index_keys={len(index)}")
    print(f"csv_rows={rows}")
    print(f"matched_rows={matched}")
    print(f"output={args.output}")


if __name__ == "__main__":
    main()
