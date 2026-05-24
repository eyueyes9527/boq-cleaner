#!/usr/bin/env python3
"""全量重新导入脚本：清空景观工程表 -> 按主清单重新导入。

当前版本只处理主清单字段、层级路径和定价模式；不再处理费用分析明细。
"""

import datetime
import json
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import config
from auto_split_table import import_contract

BASE_TOKEN = config.BASE_TOKEN
LARK_CLI = config.LARK_CLI
TABLE_ID = config.PRIMARY_TABLE_ID
SOURCE_DIRS = [
    r"D:\单价库\景观工程",
    r"D:\单价库\景观工程\第二批",
]


def run_lark(cmd, timeout=60):
    env = os.environ.copy()
    env["LARK_CLI_NO_PROXY"] = "1"
    result = subprocess.run(
        [LARK_CLI] + cmd,
        capture_output=True,
        timeout=timeout,
        env=env,
    )
    stdout = result.stdout.decode("utf-8", errors="replace").strip()
    stderr = result.stderr.decode("utf-8", errors="replace").strip()
    return stdout, stderr, result.returncode


def extract_json(text):
    start = next((i for i, ch in enumerate(text) if ch in "[{"), -1)
    if start > 0:
        text = text[start:]
    try:
        return json.loads(text)
    except Exception:
        return None


def iter_contracts():
    contracts = []
    for directory in SOURCE_DIRS:
        if not os.path.exists(directory):
            continue
        for fname in sorted(os.listdir(directory)):
            if fname.startswith("~$") or not fname.endswith(".xlsx"):
                continue
            contracts.append({
                "name": fname.replace(".xlsx", ""),
                "path": os.path.join(directory, fname),
            })
    return contracts


def clear_feishu_table():
    print("\n清空飞书景观工程表")
    record_ids = []
    offset = 0

    while True:
        stdout, _, _ = run_lark([
            "base", "+record-list",
            "--base-token", BASE_TOKEN,
            "--table-id", TABLE_ID,
            "--limit", "200",
            "--offset", str(offset),
            "--format", "json",
            "--as", "user",
        ])
        data = extract_json(stdout)
        if not data or not data.get("ok"):
            break
        batch = data["data"].get("record_id_list", [])
        record_ids.extend(batch)
        print(f"  已读取 {len(record_ids)} 条")
        if not data["data"].get("has_more", False):
            break
        offset += len(batch)

    for i in range(0, len(record_ids), 200):
        cmd = [
            "base", "+record-delete",
            "--base-token", BASE_TOKEN,
            "--table-id", TABLE_ID,
            "--yes",
            "--as", "user",
        ]
        for rid in record_ids[i:i + 200]:
            cmd.extend(["--record-id", rid])
        run_lark(cmd, timeout=120)
        print(f"  已删除 {min(i + 200, len(record_ids))}/{len(record_ids)}")
        time.sleep(0.3)


def main():
    dry_run = "--dry-run" in sys.argv
    contracts = iter_contracts()

    print("\n" + "=" * 60)
    print("景观工程 全量重新导入（主清单版）")
    print(f"时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"合同文件: {len(contracts)} 个")
    print("=" * 60)
    for contract in contracts:
        print(f"  {contract['name']}")

    if dry_run:
        print("\nDRY-RUN 完成，未清空、未导入")
        return

    if "--confirm" not in sys.argv:
        confirm = input("\n即将清空景观工程表并重新导入全部主清单数据，输入 YES 确认: ")
        if confirm != "YES":
            print("已取消")
            return

    clear_feishu_table()
    for i, contract in enumerate(contracts, 1):
        print(f"\n[{i}/{len(contracts)}] {contract['name']}")
        import_contract(contract["name"], contract["path"], dry_run=False)

    print("\n全量重新导入完成")


if __name__ == "__main__":
    main()
