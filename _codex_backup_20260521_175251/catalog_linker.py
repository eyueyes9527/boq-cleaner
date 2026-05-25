#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Resolve and attach Feishu Base catalog link values during import."""

from __future__ import annotations

import json
import os
import re
import subprocess
import time
from pathlib import Path
from urllib.parse import unquote


BASE_TOKEN = "ARDubM91FaL62esCg1hcnHYtnFh"
DIRECTORY_TABLE_ID = "tblobzr4LplrbiYT"
DIRECTORY_LIST_NAME_FIELD_ID = "fldUMgZolF"
DIRECTORY_CONTRACT_FIELD_ID = "fldup23xB1"

# Target detail table -> catalog link field id.
CATALOG_LINK_FIELD_BY_TABLE = {
    "tblY7o8bNiBxEYGA": "fldPFg6p1L",
    "tbleZCShQmE1aRhQ": "fld6BcTVBu",
    "tblRdm1rcN46iGuw": "fldEoK947G",
}

MAX_RETRIES = 5


class CatalogLinkError(RuntimeError):
    pass


def catalog_link_field_id(table_id: str | None) -> str | None:
    return CATALOG_LINK_FIELD_BY_TABLE.get(table_id or "")


def _strip_ext(value: str | None) -> str | None:
    if not value:
        return value
    return re.sub(r"\.(xlsx|xlsm|xls|csv)$", "", value, flags=re.IGNORECASE)


def _normalize_key(value: str | None) -> str | None:
    if value is None:
        return None
    text = unquote(str(value)).strip()
    if not text:
        return None
    text = re.sub(r"\s+", "", text)
    text = (
        text.replace("（", "(")
        .replace("）", ")")
        .replace("，", ",")
        .lower()
    )
    return text


def _try_repair_mojibake(value: str | None) -> str | None:
    if not value:
        return None
    try:
        return value.encode("gbk", errors="strict").decode("utf-8", errors="strict")
    except UnicodeError:
        return None


def _key_variants(value: str | None) -> list[str]:
    key = _normalize_key(value)
    if not key:
        return []

    variants: list[str] = []

    def add(candidate: str | None) -> None:
        if candidate and candidate not in variants:
            variants.append(candidate)

    add(key)
    without_leading_number = re.sub(r"^[0-9]+[_-]?", "", key)
    add(without_leading_number)
    zero_normalized = re.sub(r"(^|[^0-9])0+([0-9]+-)", r"\1\2", key)
    add(zero_normalized)
    add(re.sub(r"(^|[^0-9])0+([0-9]+-)", r"\1\2", without_leading_number))
    return variants


def _source_file_from_remark(remark: str | None) -> str | None:
    if not remark:
        return None
    text = str(remark)
    after = text
    first_ext = re.search(r"\.(xlsx|xlsm|xls|csv)", text, flags=re.IGNORECASE)
    colon_at = text.find(":")
    if colon_at >= 0 and (not first_ext or colon_at < first_ext.start()):
        after = text[colon_at + 1 :]

    ext_match = re.search(r"^(.+?\.(xlsx|xlsm|xls|csv))", after, flags=re.IGNORECASE)
    if ext_match:
        return ext_match.group(1).strip()
    marker = after.find(" | Sheet:")
    if marker > 0:
        return after[:marker].strip()
    return after.strip() or None


def _run_lark_json(
    lark_cli: str | Path,
    args: list[str],
    cwd: str | Path | None = None,
) -> dict:
    cmd = [str(lark_cli), *args]
    env = os.environ.copy()
    env["LARK_CLI_NO_PROXY"] = "1"
    last_output = ""
    for attempt in range(1, MAX_RETRIES + 1):
        result = subprocess.run(
            cmd,
            cwd=str(cwd) if cwd else None,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
            env=env,
        )
        last_output = (result.stdout or "") + "\n" + (result.stderr or "")
        if result.returncode == 0:
            start = min(
                [idx for idx in (last_output.find("{"), last_output.find("[")) if idx >= 0],
                default=-1,
            )
            if start < 0:
                raise CatalogLinkError(f"lark-cli returned no JSON: {last_output[:500]}")
            return json.loads(last_output[start:])
        if "limited" not in last_output and "800004135" not in last_output:
            break
        if attempt < MAX_RETRIES:
            time.sleep(3 * attempt)
    raise CatalogLinkError(f"lark-cli failed: {last_output[:1000]}")


def _read_table_records(
    lark_cli: str | Path,
    base_token: str,
    table_id: str,
    field_ids: list[str],
    cwd: str | Path | None = None,
) -> list[dict]:
    rows: list[dict] = []
    offset = 0
    limit = 200
    while True:
        args = [
            "base",
            "+record-list",
            "--base-token",
            base_token,
            "--table-id",
            table_id,
            "--offset",
            str(offset),
            "--limit",
            str(limit),
            "--format",
            "json",
            "--as",
            "user",
        ]
        for field_id in field_ids:
            args.extend(["--field-id", field_id])

        envelope = _run_lark_json(lark_cli, args, cwd=cwd)
        data = envelope.get("data", {})
        field_id_list = [str(field_id) for field_id in data.get("field_id_list", [])]
        record_ids = [str(record_id) for record_id in data.get("record_id_list", [])]
        value_rows = data.get("data", [])
        for idx, record_id in enumerate(record_ids):
            values = value_rows[idx] if idx < len(value_rows) else []
            row = {"record_id": record_id}
            for field_index, field_id in enumerate(field_id_list):
                row[field_id] = values[field_index] if field_index < len(values) else None
            rows.append(row)

        if not data.get("has_more", False):
            break
        offset += limit
    return rows


def load_catalog_map(
    lark_cli: str | Path,
    base_token: str = BASE_TOKEN,
    cwd: str | Path | None = None,
) -> dict[str, str]:
    directory_rows = _read_table_records(
        lark_cli,
        base_token,
        DIRECTORY_TABLE_ID,
        [DIRECTORY_LIST_NAME_FIELD_ID, DIRECTORY_CONTRACT_FIELD_ID],
        cwd=cwd,
    )
    key_map: dict[str, str] = {}
    ambiguous: dict[str, set[str]] = {}

    def add_key(raw: str | None, record_id: str) -> None:
        for key in _key_variants(raw):
            if key in ambiguous:
                continue
            existing = key_map.get(key)
            if existing and existing != record_id:
                ambiguous[key] = {existing, record_id}
                key_map.pop(key, None)
            else:
                key_map[key] = record_id

    for row in directory_rows:
        record_id = row["record_id"]
        for raw in (
            row.get(DIRECTORY_LIST_NAME_FIELD_ID),
            _strip_ext(row.get(DIRECTORY_LIST_NAME_FIELD_ID)),
            row.get(DIRECTORY_CONTRACT_FIELD_ID),
            _strip_ext(row.get(DIRECTORY_CONTRACT_FIELD_ID)),
        ):
            add_key(raw, record_id)

    if ambiguous:
        raise CatalogLinkError(f"ambiguous catalog keys: {len(ambiguous)}")
    return key_map


def resolve_catalog_record_id(
    catalog_map: dict[str, str],
    *,
    source_file: str | None = None,
    source_name: str | None = None,
    remark: str | None = None,
) -> str | None:
    remark_source_file = _source_file_from_remark(remark)
    candidates = [
        source_file,
        _strip_ext(source_file),
        remark_source_file,
        _strip_ext(remark_source_file),
        source_name,
        _strip_ext(source_name),
        _try_repair_mojibake(source_file),
        _strip_ext(_try_repair_mojibake(source_file)),
        _try_repair_mojibake(remark_source_file),
        _strip_ext(_try_repair_mojibake(remark_source_file)),
        _try_repair_mojibake(source_name),
        _strip_ext(_try_repair_mojibake(source_name)),
    ]
    for candidate in candidates:
        for key in _key_variants(candidate):
            record_id = catalog_map.get(key)
            if record_id:
                return record_id
    return None


def attach_catalog_links(
    records: list[dict],
    *,
    table_id: str,
    lark_cli: str | Path,
    base_token: str = BASE_TOKEN,
    cwd: str | Path | None = None,
    source_name_field: str | None = None,
    remark_field: str | None = None,
    source_file_field: str = "_source_file",
    strict: bool = True,
) -> tuple[list[dict], str | None, dict]:
    link_field_id = catalog_link_field_id(table_id)
    if not link_field_id:
        return records, None, {
            "enabled": False,
            "reason": f"no catalog link field configured for table {table_id}",
        }

    catalog_map = load_catalog_map(lark_cli, base_token=base_token, cwd=cwd)
    unmatched: list[dict] = []
    matched = 0

    for index, record in enumerate(records):
        source_name = record.get(source_name_field) if source_name_field else None
        remark = record.get(remark_field) if remark_field else None
        source_file = record.get(source_file_field)
        catalog_record_id = resolve_catalog_record_id(
            catalog_map,
            source_file=source_file,
            source_name=source_name,
            remark=remark,
        )
        if catalog_record_id:
            record[link_field_id] = [{"id": catalog_record_id}]
            matched += 1
        else:
            unmatched.append(
                {
                    "index": index,
                    "source_file": source_file,
                    "source_name": source_name,
                    "remark": remark,
                }
            )

    summary = {
        "enabled": True,
        "field_id": link_field_id,
        "records": len(records),
        "matched": matched,
        "unmatched": len(unmatched),
        "unmatched_samples": unmatched[:20],
    }
    if unmatched and strict:
        raise CatalogLinkError(
            f"catalog link match failed for {len(unmatched)}/{len(records)} records"
        )
    return records, link_field_id, summary
