#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动分表导入脚本 v1.3
功能：
1. 解析Excel，提取单价记录
2. 检查是否有缺失单价的记录（导入前检查）
3. 导入前检查目标表剩余空间
4. 空间不足时提示手动创建新表
5. 确保同一合同的所有记录在同一张表内
6. 不再上传附件（已移除）
7. 【新增】导入后自动检查刚导入的记录，验证单价完整性

使用方式：
1. 修改底部 CONTRACT_NAME 和 EXCEL_PATH
2. 确认顶部 TABLE_ID 是正确的表ID
3. 运行: python auto_split_table.py
"""

import subprocess
import json
import time
import os
import re
import datetime
import pandas as pd
from typing import List, Dict
import sys

from excel_reader import clean_excel

# ========== 配置区 ==========
BASE_TOKEN = "ARDubM91FaL62esCg1hcnHYtnFh"
TABLE_ID = "tbleZCShQmE1aRhQ"  # 当前使用的表ID（需要时在飞书里复制新表ID替换这里）
LARK_CLI = r"C:\Users\56237\AppData\Roaming\npm\lark-cli.cmd"
MAX_ROWS = 20000
SAFETY_MARGIN = 500  # 预留500行缓冲

# 审计日志路径
AUDIT_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runs", "lark_audit.csv")


def append_audit_log(cmd_label: str, returncode: int, elapsed: float, error: str = ""):
    """记录 lark-cli 调用日志"""
    os.makedirs(os.path.dirname(AUDIT_LOG), exist_ok=True)
    import csv
    is_new = not os.path.exists(AUDIT_LOG)
    with open(AUDIT_LOG, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if is_new:
            w.writerow(["时间", "命令", "状态码", "耗时(秒)", "错误"])
        import datetime
        w.writerow([datetime.datetime.now().isoformat(), cmd_label, returncode, f"{elapsed:.2f}", error])

# 字段ID映射（与原表一致，已移除附件字段 fldt8TxJ6q）
FIELD_IDS = [
    "fldJkEw41C", "flddaIcKu7", "fldEvvjofm", "fldlIfsrTX",
    "fldgEQEjf9", "fldiVMozmY", "fldxVGnJpn", "fldjhCzC8y",
    "fldPBDN6QO", "fldbFXoOC5", "fldGGC9sXw", "fldpLjVfSv",
]

FIELD_NAME_MAP = {
    'fldJkEw41C': '备注',
    'flddaIcKu7': '合同名称',
    'fldEvvjofm': '页签',
    'fldlIfsrTX': '层级路径',
    'fldgEQEjf9': '项目编号',
    'fldiVMozmY': '项目名称',
    'fldxVGnJpn': '项目特征描述',
    'fldjhCzC8y': '计量单位',
    'fldPBDN6QO': '工程数量',
    'fldbFXoOC5': '不含税单价',
    'fldGGC9sXw': '汇总合价',
    'fldpLjVfSv': '定价信息',
}


def join_path(*parts):
    result = []
    for part in parts:
        if not part:
            continue
        for piece in str(part).replace('|', '/').split('/'):
            piece = piece.strip()
            if piece and (not result or result[-1] != piece):
                result.append(piece)
    return '/'.join(result)


# ========== Excel解析函数 ==========

def discover_target_sheets(excel_path):
    """发现需要处理的sheet"""
    xl_file = pd.ExcelFile(excel_path)
    results = []
    for sname in xl_file.sheet_names:
        if '实体工程量清单' not in sname:
            continue
        if '汇总表' in sname or '综合单价分析' in sname:
            continue
        results.append(sname)
    return results


def find_header_row(df):
    """查找表头行"""
    for i in range(min(15, len(df))):
        row = df.iloc[i]
        if any('项目编号' in str(cell) for cell in row if not pd.isna(cell)):
            return i
    return None


def get_col_indices(header_row):
    """映射Excel列索引"""
    mapping = {}
    for idx, val in enumerate(header_row):
        if pd.isna(val):
            continue
        s = str(val).strip()
        if '项目编号' in s:
            mapping['项目编号'] = idx
        elif '项目名称' in s and '项目编号' not in s:
            mapping['项目名称'] = idx
        elif '项目特征描述' in s:
            mapping['项目特征描述'] = idx
        elif '类别' in s:
            mapping['类别'] = idx
        elif '计量单位' in s or '单位' in s:
            mapping['计量单位'] = idx
        elif '工程数量' in s or '工程量' in s:
            mapping['工程数量'] = idx
        elif '不含' in s and ('单价' in s or '综合单价' in s):
            mapping['不含税单价'] = idx
        elif '不含' in s and '合价' in s:
            mapping['汇总合价'] = idx
        elif '合价' in s and '不含' not in s:
            mapping['汇总合价'] = idx
        # 新增：基准价招标相关列（只匹配第一个，防止重复列覆盖）
        elif '基准价' in s and '不含' in s and '基准价' not in mapping:
            mapping['基准价'] = idx
        elif ('上下浮率' in s or '浮率' in s) and '上下浮率' not in mapping:
            mapping['上下浮率'] = idx
        elif '投标价' in s and '不含' in s and '投标价' not in mapping:
            mapping['投标价'] = idx
    return mapping


def parse_num(v):
    """安全转换数字"""
    if pd.isna(v):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except:
        return None


def extract_all_records(sheet_name, excel_path):
    """从指定sheet提取所有记录"""
    try:
        df_sample = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=20)
        header_idx = find_header_row(df_sample)
        if header_idx is None:
            return []
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"  [WARN] 读取失败: {str(e)}")
        return []

    header_row = df.iloc[header_idx]
    col_map = get_col_indices(header_row)
    results = []
    current_subject = ''
    current_subtab = ''

    for i in range(header_idx + 1, len(df)):
        row = df.iloc[i]
        cat_idx = col_map.get('类别')
        if cat_idx is None or cat_idx >= len(row):
            continue
        cat_val = row.iloc[cat_idx]
        if pd.isna(cat_val):
            continue
        cat = str(cat_val).strip()

        if cat == '科目':
            name_idx = col_map.get('项目名称')
            if name_idx is not None and name_idx < len(row):
                name = row.iloc[name_idx]
                if not pd.isna(name):
                    current_subject = str(name).strip()
            continue
        elif cat in ('二级页签', '页签'):
            name_idx = col_map.get('项目名称')
            if name_idx is not None and name_idx < len(row):
                name = row.iloc[name_idx]
                if not pd.isna(name):
                    current_subtab = str(name).strip()
            continue
        elif cat == '三级页签':
            name_idx = col_map.get('项目名称')
            if name_idx is not None and name_idx < len(row):
                name = row.iloc[name_idx]
                if not pd.isna(name):
                    current_subtab = join_path(current_subtab, str(name).strip())
            continue
        elif cat != '清单项':
            name_idx = col_map.get('项目名称')
            price_idx = col_map.get('不含税单价') or col_map.get('投标价')
            unit_idx = col_map.get('计量单位')
            name = row.iloc[name_idx] if name_idx is not None and name_idx < len(row) else None
            price = parse_num(row.iloc[price_idx]) if price_idx is not None and price_idx < len(row) else None
            unit = row.iloc[unit_idx] if unit_idx is not None and unit_idx < len(row) else None
            if name_idx is not None and not pd.isna(name) and price is None and (unit_idx is None or pd.isna(unit)):
                current_subtab = str(name).strip()
            continue

        code_idx = col_map.get('项目编号')
        if code_idx is None:
            continue
        code_val = row.iloc[code_idx] if code_idx < len(row) else None
        if pd.isna(code_val) or str(code_val).strip() == '':
            continue

        # 处理不含税单价：检查是否为基准价招标
        unit_price = None
        pricing_info = ""
        is_b_type = bool(col_map.get('基准价') and col_map.get('上下浮率'))

        if is_b_type:
            # 基准价招标：提取基准价和上下浮率
            base_price = parse_num(row.iloc[col_map['基准价']]) if col_map['基准价'] < len(row) else None
            float_rate = parse_num(row.iloc[col_map['上下浮率']]) if col_map['上下浮率'] < len(row) else None

            # 优先使用"不含税单价"列（如果有）
            if col_map.get('不含税单价') and col_map['不含税单价'] < len(row):
                unit_price = parse_num(row.iloc[col_map['不含税单价']])
            elif col_map.get('投标价') and col_map['投标价'] < len(row):
                unit_price = parse_num(row.iloc[col_map['投标价']])
            elif base_price is not None and float_rate is not None:
                unit_price = base_price * (1 - float_rate / 100)

            # 生成定价信息字符串
            if base_price is not None and float_rate is not None:
                pricing_info = f"基准价:{base_price}, 上下浮率:{float_rate}%"

            # 浮率验证：区分乘数模式(|浮率|>0.5)和幅度模式(|浮率|<=0.5)
            if base_price is not None and float_rate is not None and unit_price is not None:
                if abs(float_rate) > 0.5:
                    expected = round(base_price * float_rate, 2)
                    formula = f"基准价{base_price}×{float_rate}"
                else:
                    expected = round(base_price * (1 + float_rate), 2)
                    formula = f"基准价{base_price}×(1+{float_rate})"
                if abs(unit_price - expected) > 0.05:
                    print(f"\n[ERROR] 浮率异常！导入已停止")
                    print(f"  投标价={unit_price} ≠ 计算值={expected}")
                    print(f"  公式: {formula}")
                    print(f"  文件: {excel_path} | Sheet: {sheet_name} | 行: {i+2}")
                    print(f"  项目编号: {str(row.iloc[code_idx]).strip() if code_idx < len(row) else 'N/A'}")
                    print(f"  请检查数据源，修复后重新运行。\n")
                    exit(1)  # 停止导入，修复后重跑
        else:
            # 自主报价
            if col_map.get('不含税单价') and col_map['不含税单价'] < len(row):
                unit_price = parse_num(row.iloc[col_map['不含税单价']])
            elif col_map.get('投标价') and col_map['投标价'] < len(row):
                unit_price = parse_num(row.iloc[col_map['投标价']])
            pricing_info = "自主报价"

        now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        record = {
            '备注':         now_str,
            '合同名称':    '',
            '页签':         sheet_name,
            '层级路径':  join_path(current_subject, current_subtab),
            '项目编号':    str(code_val).strip(),
            '项目名称':    (str(row.iloc[col_map['项目名称']]).strip() if col_map.get('项目名称') and col_map['项目名称'] < len(row) and not pd.isna(row.iloc[col_map['项目名称']]) else ''),
            '项目特征描述': (str(row.iloc[col_map['项目特征描述']]).strip() if col_map.get('项目特征描述') and col_map['项目特征描述'] < len(row) and not pd.isna(row.iloc[col_map['项目特征描述']]) else ''),
            '计量单位':    (str(row.iloc[col_map['计量单位']]).strip() if col_map.get('计量单位') and col_map['计量单位'] < len(row) and not pd.isna(row.iloc[col_map['计量单位']]) else ''),
            '工程数量':    parse_num(row.iloc[col_map['工程数量']]) if col_map.get('工程数量') and col_map['工程数量'] < len(row) and not pd.isna(row.iloc[col_map['工程数量']]) else None,
            '不含税单价':  unit_price,
            '汇总合价':    parse_num(row.iloc[col_map['汇总合价']]) if col_map.get('汇总合价') and col_map['汇总合价'] < len(row) and not pd.isna(row.iloc[col_map['汇总合价']]) else None,
            '定价信息':    pricing_info,
        }
        results.append(record)
    return results


# ========== 单价检查函数 ==========

def check_missing_price(records: List[Dict]) -> List[Dict]:
    """
    检查缺失不含税单价的记录
    返回: 缺失单价的记录列表
    """
    missing = []
    for rec in records:
        price = rec.get('不含税单价')
        if price is None or (isinstance(price, (int, float)) and price == 0):
            missing.append(rec)
    return missing


def print_missing_price_report(missing: List[Dict], max_show: int = 20):
    """
    打印缺失单价的记录报告
    """
    if not missing:
        print(f"[OK] 所有记录均含不含税单价")
        return

    print(f"\n{'='*60}")
    print(f"[WARN]  发现 {len(missing)} 条记录缺失不含税单价：")
    print(f"{'='*60}")

    show_count = min(len(missing), max_show)
    for i, rec in enumerate(missing[:show_count]):
        print(f"  {i+1}. 项目编号: {rec.get('项目编号', 'N/A')}")
        print(f"     项目名称: {rec.get('项目名称', 'N/A')[:40]}")
        print(f"     层级路径: {rec.get('层级路径', 'N/A')}")
        print()

    if len(missing) > max_show:
        print(f"  ... 还有 {len(missing) - max_show} 条未显示")

    print(f"{'='*60}\n")


# ========== 飞书操作函数 ==========

def get_table_record_count(table_id: str) -> int:
    """
    获取表的记录总数（翻页遍历所有记录）
    使用 --format json 获取 JSON 格式输出
    """
    total = 0
    offset = 0
    limit = 200

    while True:
        cmd = [
            LARK_CLI, 'base', '+record-list',
            '--base-token', BASE_TOKEN,
            '--table-id', table_id,
            '--limit', str(limit),
            '--offset', str(offset),
            '--format', 'json'
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=30)

        try:
            data = json.loads(result.stdout)
            records = data.get('data', {}).get('items', [])
            total += len(records)

            # 检查是否还有更多记录
            if not data.get('data', {}).get('has_more', False):
                break
            offset += limit
        except Exception as e:
            print(f"  [WARN] 解析错误: {e}")
            break

    return total


def get_records_by_import_time(table_id: str, import_time: str) -> List[Dict]:
    """
    根据导入时间获取记录（筛选"备注"字段匹配 import_time 的记录）
    使用 --format json 获取结构化数据
    """
    all_records = []
    offset = 0
    limit = 500  # 增加每页数量，减少请求次数

    print(f"  [CHECK] 开始查询导入时间为 '{import_time}' 的记录...")

    while True:
        cmd = [
            LARK_CLI, 'base', '+record-list',
            '--base-token', BASE_TOKEN,
            '--table-id', table_id,
            '--limit', str(limit),
            '--offset', str(offset),
            '--format', 'json'
        ]
        
        # 设置环境变量禁用代理警告
        env = os.environ.copy()
        env['LARK_CLI_NO_PROXY'] = '1'
        
        result = subprocess.run(cmd, capture_output=True, timeout=60, env=env)
        stdout_str = result.stdout.decode('utf-8', errors='replace').strip()
        stderr_str = result.stderr.decode('utf-8', errors='replace').strip()
        
        # 检查是否有错误
        if result.returncode != 0:
            print(f"  [WARN] CLI错误: {stderr_str[:200]}")
            break
        
        try:
            # 提取JSON部分（过滤掉非JSON的警告信息）
            json_start = -1
            for i, char in enumerate(stdout_str):
                if char in '[{':
                    json_start = i
                    break
            if json_start > 0:
                stdout_str = stdout_str[json_start:]
            
            data = json.loads(stdout_str)
            
            # 检查返回格式
            if not data.get('ok'):
                print(f"  [WARN] API返回错误: {data}")
                break
            
            # lark-cli 返回的格式：
            # data['data']['data'] = 二维数组（记录值）
            # data['data']['record_id_list'] = record_id列表
            # data['data']['field_id_list'] = 字段ID列表
            
            records_data = data['data'].get('data', [])
            record_ids = data['data'].get('record_id_list', [])
            field_ids = data['data'].get('field_id_list', [])
            
            print(f"    获取 {len(records_data)} 条记录 (offset={offset})")
            
            # 将二维数组转换为对象数组，并筛选匹配导入时间的记录
            for i, record_values in enumerate(records_data):
                record_id = record_ids[i] if i < len(record_ids) else f'unknown_{i}'
                
                # 构造记录对象
                rec = {'record_id': record_id}
                for j, field_id in enumerate(field_ids):
                    if j < len(record_values):
                        rec[field_id] = record_values[j]
                
                # 筛选备注匹配导入时间的记录
                remark = rec.get('fldJkEw41C', '')
                if import_time in str(remark):
                    # 构造统一格式的记录
                    clean_rec = {
                        'record_id': record_id,
                        '备注': remark,
                        '不含税单价': rec.get('fldbFXoOC5'),
                        '合同名称': rec.get('flddaIcKu7', ''),
                        '层级路径': rec.get('fldlIfsrTX', ''),
                        '项目编号': rec.get('fldgEQEjf9', ''),
                        '项目名称': rec.get('fldiVMozmY', ''),
                    }
                    all_records.append(clean_rec)

            # 检查是否还有更多记录
            if not data['data'].get('has_more', False):
                break
            offset += limit
        except Exception as e:
            print(f"  [WARN] 解析错误: {e}")
            print(f"     stdout前500字符: {stdout_str[:500]}")
            break

    print(f"  [OK] 找到 {len(all_records)} 条匹配的记录")
    return all_records


def check_imported_records(table_id: str, import_time: str) -> List[Dict]:
    """
    检查刚导入的记录是否有缺失不含税单价
    返回: 缺失单价的记录列表
    """
    print(f"\n[CHECK] 检查刚导入的记录（导入时间: {import_time}）...")

    # 获取匹配导入时间的记录
    records = get_records_by_import_time(table_id, import_time)
    print(f"   找到 {len(records)} 条刚导入的记录")

    if not records:
        print(f"  [WARN] 未找到刚导入的记录，无法检查")
        return []

    # 检查缺失单价
    missing = []
    for rec in records:
        price = rec.get('不含税单价')
        if price is None or price == '' or (isinstance(price, (int, float)) and price == 0):
            missing.append(rec)

    return missing


def print_import_check_report(missing: List[Dict], import_time: str):
    """打印导入检查报告"""
    if not missing:
        print(f"\n{'='*60}")
        print(f"[OK] 导入验证通过！所有刚导入的记录均含不含税单价")
        print(f"   导入时间: {import_time}")
        print(f"{'='*60}\n")
        return

    print(f"\n{'='*60}")
    print(f"[WARN]  导入验证失败！发现 {len(missing)} 条记录缺失不含税单价：")
    print(f"   导入时间: {import_time}")
    print(f"{'='*60}")

    for i, rec in enumerate(missing[:20]):  # 最多显示20条
        print(f"  {i+1}. 项目编号: {rec.get('项目编号', 'N/A')}")
        print(f"     项目名称: {str(rec.get('项目名称', 'N/A'))[:40]}")
        print(f"     层级路径: {rec.get('层级路径', 'N/A')}")
        print(f"     不含税单价: {rec.get('不含税单价')}")
        print()

    if len(missing) > 20:
        print(f"  ... 还有 {len(missing) - 20} 条未显示")

    print(f"{'='*60}\n")


def check_table_space(table_id: str, needed: int) -> bool:
    """
    检查表是否有足够空间
    返回: True=空间足够, False=需要新表
    """
    current = get_table_record_count(table_id)
    available = MAX_ROWS - SAFETY_MARGIN - current

    print(f"[STAT] 当前表记录数: {current}/{MAX_ROWS}")
    print(f"[STAT] 可用空间: {available} 条")
    print(f"[STAT] 需要空间: {needed} 条")

    if needed <= available:
        return True
    else:
        print(f"\n[WARN] 空间不足！需要 {needed} 条，但只剩 {available} 条")
        return False


def prompt_create_new_table(current_count: int, needed: int):
    """提示用户手动创建新表"""
    print(f"\n{'='*60}")
    print("[WARN]  需要创建新数据表")
    print(f"{'='*60}")
    print(f"\n请按以下步骤操作：")
    print(f"1. 打开飞书多维表格")
    print(f"2. 点击左侧『+』号创建新数据表")
    print(f"3. 表名格式：单价数据库_02、单价数据库_03（按顺序编号）")
    print(f"4. 复制原表的所有字段（项目编号、项目名称、不含税单价等）")
    print(f"5. 创建完成后，在脚本顶部修改 TABLE_ID = '新表ID'")
    print(f"\n当前表ID: {TABLE_ID}")
    print(f"当前记录数: {current_count}")
    print(f"需要导入: {needed} 条")
    print(f"{'='*60}\n")


def lark_batch_create(table_id: str, rows_chunk: List[List]) -> Dict:
    """批量创建记录（通过 lark-cli），带重试+退避"""
    # 注意：--json 参数需要直接传JSON字符串，不能用 @filename
    payload_str = json.dumps({'fields': FIELD_IDS, 'rows': rows_chunk}, ensure_ascii=False)
    cmd = [
        LARK_CLI, 'base', '+record-batch-create',
        '--base-token', BASE_TOKEN,
        '--table-id', table_id,
        '--json', payload_str,
    ]
    cmd_label = f"batch-create {table_id} ({len(rows_chunk)} 条)"

    max_retries = 3
    for attempt in range(max_retries):
        env = os.environ.copy()
        env['LARK_CLI_NO_PROXY'] = '1'
        
        start = time.time()
        result = subprocess.run(cmd, capture_output=True, timeout=60, env=env)
        elapsed = time.time() - start
        
        stdout_str = result.stdout.decode('utf-8', errors='replace').strip()
        stderr_str = result.stderr.decode('utf-8', errors='replace').strip()
        
        if result.returncode == 0:
            append_audit_log(cmd_label, 0, elapsed)
            # 提取JSON部分（过滤掉非JSON的警告信息）
            try:
                json_start = -1
                for i, char in enumerate(stdout_str):
                    if char in '[{':
                        json_start = i
                        break
                if json_start > 0:
                    stdout_str = stdout_str[json_start:]
                return json.loads(stdout_str)
            except Exception as e:
                return {'ok': False, 'error': f'JSON解析失败: {str(e)}', 'stdout': stdout_str[:500], 'stderr': stderr_str[:500]}
        
        # 重试前等待（退避：2s → 4s → 8s）
        wait = 2 ** attempt
        error_msg = f'CLI返回码{result.returncode}'
        append_audit_log(cmd_label, result.returncode, elapsed, error_msg)
        if attempt < max_retries - 1:
            print(f"  [WARN]  {error_msg}，{wait}s 后重试 ({attempt+1}/{max_retries})")
            time.sleep(wait)
        else:
            return {'ok': False, 'error': error_msg, 'stderr': stderr_str, 'stdout': stdout_str}


def records_to_rows(records: List[Dict]) -> List[List]:
    """将记录列表转换为 lark-cli 所需的行格式"""
    rows = []
    for rec in records:
        row = []
        for fid in FIELD_IDS:
            fname = FIELD_NAME_MAP.get(fid, '')
            row.append(rec.get(fname, ''))
        rows.append(row)
    return rows


# ========== 运行归档 ==========

def save_run_archive(run_dir: str, contract_name: str, records: list, table_id: str, import_time: str, dry_run: bool):
    """保存本次运行的归档文件到 runs/ 目录"""
    import csv
    os.makedirs(run_dir, exist_ok=True)

    # 1. CSV 数据快照
    csv_path = os.path.join(run_dir, "cleaned_items.csv")
    if records:
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=records[0].keys())
            w.writeheader()
            w.writerows(records)
        print(f"  [SAVE] 数据快照: {csv_path} ({len(records)} 条)")

    # 2. 配置快照
    config_snapshot = {
        "contract_name": contract_name,
        "table_id": table_id,
        "import_time": import_time,
        "dry_run": dry_run,
        "record_count": len(records),
        "timestamp": __import__("datetime").datetime.now().isoformat(),
    }
    config_path = os.path.join(run_dir, "config_snapshot.json")
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config_snapshot, f, ensure_ascii=False, indent=2)
    print(f"  [SAVE] 配置快照: {config_path}")


# ========== 主流程 ==========

def import_contract(contract_name: str, excel_path: str, table_id: str = None, skip_no_price: bool = False, dry_run: bool = False):
    """
    导入单个合同到指定表

    参数:
        contract_name: 合同名称
        excel_path: Excel文件路径
        table_id: 目标表ID（None则使用全局 TABLE_ID）
        skip_no_price: 是否跳过缺失单价的记录（True=跳过，False=保留）
        dry_run: 是否仅测试不实际导入（True=只检查不导入）
    """
    if table_id is None:
        table_id = TABLE_ID

    # 生成导入时间戳（用于标记刚导入的记录）
    import_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    print(f"\n{'='*60}")
    print(f"开始导入合同: {contract_name}")
    print(f"Excel: {excel_path}")
    print(f"目标表ID: {table_id}")
    print(f"导入时间标记: {import_time}")
    print(f"{'='*60}\n")

    # 1. 解析Excel
    print(f"[READ] 正在解析Excel...")
    if not os.path.exists(excel_path):
        print(f"[ERROR] 文件不存在: {excel_path}")
        return

    try:
        cleaned_items, anomalies = clean_excel(excel_path)
    except SystemExit:
        raise
    except Exception as exc:
        print(f"[ERROR] 解析失败: {exc}")
        return

    if anomalies:
        print(f"   [WARN] 解析异常: {len(anomalies)} 条")
        for anomaly in anomalies[:10]:
            print(
                f"     {anomaly.source_sheet or '-'} 行{anomaly.source_row or '-'} "
                f"{anomaly.异常类型}: {anomaly.异常描述}"
            )

    all_records = []
    sheet_counts = {}
    for item in cleaned_items:
        sheet_counts[item.source_sheet] = sheet_counts.get(item.source_sheet, 0) + 1
        all_records.append({
            '备注': import_time,
            '合同名称': contract_name,
            '页签': item.页签,
            '层级路径': item.层级路径 or join_path(item.科目, item.二级页签),
            '项目编号': item.项目编号,
            '项目名称': item.项目名称,
            '项目特征描述': item.项目特征描述,
            '计量单位': item.计量单位,
            '工程数量': item.工程数量,
            '不含税单价': item.不含税单价,
            '汇总合价': item.汇总合价,
            '定价信息': item.定价信息,
        })

    print(f"   解析到 {len(sheet_counts)} 个目标sheet:")
    for sheet_name, count in sorted(sheet_counts.items()):
        print(f"     - {sheet_name}: {count} 条")

    if not all_records:
        print(f"\n[WARN] 无数据，退出")
        return

    record_count = len(all_records)
    print(f"\n[STAT] 合同记录数: {record_count}")

    # 创建运行目录（用于归档）
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    run_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "runs", f"{timestamp}_import")
    save_run_archive(run_dir, contract_name, all_records, table_id, import_time, dry_run)

    # 2. 检查缺失单价
    print(f"\n[CHECK] 检查缺失不含税单价的记录...")
    missing_price = check_missing_price(all_records)
    print_missing_price_report(missing_price)

    if missing_price:
        if skip_no_price:
            print(f"[WARN] 已设置跳过缺失单价记录，将导入 {record_count - len(missing_price)} 条")
            all_records = [r for r in all_records if r not in missing_price]
            record_count = len(all_records)
        else:
            print(f"[WARN] 注意：有 {len(missing_price)} 条记录不含单价，将继续导入")

    # dry-run 只做本地解析和数据检查，不访问飞书、不写入。
    if dry_run:
        print(f"\n[DRY] DRY RUN 模式，不实际导入")
        print(f"   将导入 {record_count} 条记录到表 {table_id}")
        if missing_price:
            print(f"   [WARN] 其中 {len(missing_price)} 条缺失单价")
        print(f"\n[OK] DRY RUN 完成，无实际导入")
        return

    # 3. 检查表空间
    if not check_table_space(table_id, record_count):
        prompt_create_new_table(
            get_table_record_count(table_id),
            record_count
        )
        return

    print(f"\n[OK] 空间检查通过")

    # 4. 批量导入
    print(f"\n[RUN] 开始导入...")
    BATCH_SIZE = 50
    total_written = 0

    for i in range(0, len(all_records), BATCH_SIZE):
        batch_recs = all_records[i:i+BATCH_SIZE]
        rows_chunk = records_to_rows(batch_recs)

        resp = lark_batch_create(table_id, rows_chunk)
        if resp.get('ok'):
            rids = resp.get('data', {}).get('record_id_list', [])
            total_written += len(rids)
            print(f"  批次 {i//BATCH_SIZE + 1}: OK {len(rids)} 条 (累计 {total_written})")
        else:
            print(f"  批次 {i//BATCH_SIZE + 1}: FAIL {resp.get('error', '未知错误')}")
            # 打印详细错误信息
            if resp.get('stderr'):
                print(f"     stderr: {resp['stderr'][:200]}")
            if resp.get('stdout'):
                print(f"     stdout: {resp['stdout'][:200]}")
        time.sleep(0.3)

    print(f"\n[OK] 导入完成，共 {total_written} 条")

    # 保存导入统计到归档
    import_stats = {
        "contract_name": contract_name,
        "table_id": table_id,
        "total_records": record_count,
        "total_written": total_written,
        "missing_price": len(missing_price),
        "dry_run": dry_run,
        "import_time": import_time,
    }
    stats_path = os.path.join(run_dir, "import_stats.json")
    with open(stats_path, "w", encoding="utf-8") as f:
        json.dump(import_stats, f, ensure_ascii=False, indent=2)
    print(f"  [SAVE] 导入统计: {stats_path}")

    # 5. 检查刚导入的记录
    print(f"\n[CHECK] 开始验证导入记录的单价完整性...")
    missing = check_imported_records(table_id, import_time)
    print_import_check_report(missing, import_time)
    
    # 6. 检查并删除重复记录（仅当辅助脚本存在时执行）
    print(f"\n[CHECK] 开始检查重复记录...")
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        check_script = os.path.join(script_dir, 'check_duplicates_feishu.py')
        delete_script = os.path.join(script_dir, 'delete_duplicates.py')
        if not os.path.exists(check_script) or not os.path.exists(delete_script):
            print("  [WARN] 未找到重复记录辅助脚本，跳过该步骤")
            return

        # 调用 check_duplicates_feishu.py 生成 duplicates_to_delete.json
        result = subprocess.run(
            [sys.executable, check_script],
            capture_output=True,
            text=True,
            encoding='utf-8',
            timeout=120,
            cwd=script_dir
        )
        if result.returncode == 0:
            # 打印输出（去掉最后的提示信息）
            output_lines = result.stdout.strip().split('\n')
            for line in output_lines[:-2]:  # 去掉最后两行提示
                print(line)
            
            # 检查是否有重复记录需要删除
            duplicates_file = os.path.join(script_dir, 'duplicates_to_delete.json')
            
            if os.path.exists(duplicates_file):
                with open(duplicates_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if data.get('total_to_delete', 0) > 0:
                    # 调用 delete_duplicates.py 删除重复记录
                    result2 = subprocess.run(
                        [sys.executable, delete_script],
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        timeout=180,
                        cwd=script_dir
                    )
                    if result2.returncode == 0:
                        print(result2.stdout)
                    else:
                        print(f"  [WARN] 删除重复记录失败: {result2.stderr}")
                else:
                    print(f"\n[OK] 无重复记录")
        else:
            print(f"  [WARN] 检查重复记录失败: {result.stderr}")
    except Exception as e:
        print(f"  [WARN] 检查重复记录时发生错误: {str(e)}")

def main():
    import argparse

    parser = argparse.ArgumentParser(description="解析并导入单个工程量清单 Excel")
    parser.add_argument("contract_name", help="合同名称")
    parser.add_argument("excel_path", help="Excel 文件路径")
    parser.add_argument("--table-id", default=None, help="目标飞书表 ID，默认使用脚本配置")
    parser.add_argument("--skip-no-price", action="store_true", help="跳过缺失单价的记录")
    parser.add_argument("--dry-run", action="store_true", help="只解析和检查，不访问飞书、不写入")
    args = parser.parse_args()

    import_contract(
        args.contract_name,
        args.excel_path,
        table_id=args.table_id,
        skip_no_price=args.skip_no_price,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
