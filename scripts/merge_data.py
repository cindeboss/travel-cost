#!/usr/bin/env python3
"""
数据合并脚本

合并所有按月分片的差旅数据，生成完整的数据集和索引。
"""

import sys
import json
from pathlib import Path
from typing import Dict, List, Any
from datetime import datetime
from collections import defaultdict

# 添加父目录到路径以导入utils
sys.path.insert(0, str(Path(__file__).parent))


def parse_amount(record: Dict) -> float:
    """
    从记录中提取金额

    Args:
        record: 差旅记录

    Returns:
        金额
    """
    record_type = record.get('type', '')

    if record_type == 'flight':
        return float(record.get('price', 0))
    elif record_type == 'hotel':
        return float(record.get('price', 0))
    elif record_type == 'train':
        return float(record.get('price', 0))
    elif record_type == 'car':
        return float(record.get('totalAmount', 0))

    return 0


def parse_date_from_record(record: Dict) -> str:
    """
    从记录中提取日期

    Args:
        record: 差旅记录

    Returns:
        日期字符串 (YYYY-MM-DD)
    """
    record_type = record.get('type', '')

    if record_type == 'flight':
        time_str = record.get('departTime', '')
    elif record_type == 'hotel':
        time_str = record.get('checkInTime', '')
    elif record_type == 'train':
        time_str = record.get('departTime', '')
    elif record_type == 'car':
        time_str = record.get('pickupTime', '')
    else:
        return ''

    # 提取日期部分 (假设格式为 YYYY-MM-DD HH:MM:SS 或类似)
    if ' ' in time_str:
        return time_str.split(' ')[0]

    return time_str[:10] if len(time_str) >= 10 else ''


def get_employee_name(record: Dict) -> str:
    """
    从记录中获取员工姓名

    Args:
        record: 差旅记录

    Returns:
        员工姓名
    """
    record_type = record.get('type', '')

    if record_type == 'flight':
        return record.get('passenger', '')
    elif record_type == 'car':
        return record.get('passenger', '')
    else:
        return record.get('employee', '')


def merge_monthly_data(by_month_dir: Path) -> Dict[str, Any]:
    """
    合并所有按月分片的数据

    Args:
        by_month_dir: 按月分片数据目录

    Returns:
        合并后的数据字典
    """
    all_records = []
    months = set()
    sources = set()

    # 扫描所有JSON文件
    for filepath in by_month_dir.glob('*.json'):
        if filepath.name.startswith('roster_'):
            continue  # 跳过花名册文件

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # 提取记录
            records = data.get('records', [])
            all_records.extend(records)

            # 收集元数据
            if 'month' in data:
                months.add(data['month'])
            if 'source' in data:
                sources.add(data['source'])

            print(f'  读取 {filepath.name}: {len(records)} 条记录')

        except Exception as e:
            print(f'  警告: 无法读取 {filepath.name}: {e}')

    print(f'\n总共合并 {len(all_records)} 条记录')

    return {
        'records': all_records,
        'months': sorted(months),
        'sources': sorted(sources)
    }


def build_summary(records: List[Dict]) -> Dict[str, Any]:
    """
    构建数据摘要统计

    Args:
        records: 所有记录

    Returns:
        摘要统计字典
    """
    total_amount = 0
    total_records = len(records)

    by_dept = defaultdict(lambda: {'amount': 0, 'count': 0})
    by_type = defaultdict(lambda: {'amount': 0, 'count': 0})
    by_month = defaultdict(lambda: {'amount': 0, 'count': 0})
    by_employee = defaultdict(lambda: {'amount': 0, 'count': 0})
    by_source = defaultdict(lambda: {'amount': 0, 'count': 0})

    for record in records:
        amount = parse_amount(record)
        total_amount += amount

        dept = record.get('deptLevel1', '未知部门')
        record_type = record.get('type', 'unknown')
        date_str = parse_date_from_record(record)
        month = date_str[:7] if len(date_str) >= 7 else '未知月份'
        employee = get_employee_name(record) or '未知员工'
        source = record.get('source', '未知来源')

        by_dept[dept]['amount'] += amount
        by_dept[dept]['count'] += 1

        by_type[record_type]['amount'] += amount
        by_type[record_type]['count'] += 1

        by_month[month]['amount'] += amount
        by_month[month]['count'] += 1

        by_employee[employee]['amount'] += amount
        by_employee[employee]['count'] += 1

        by_source[source]['amount'] += amount
        by_source[source]['count'] += 1

    return {
        'totalAmount': round(total_amount, 2),
        'totalRecords': total_records,
        'byDept': {k: {'amount': round(v['amount'], 2), 'count': v['count']}
                   for k, v in sorted(by_dept.items(), key=lambda x: x[1]['amount'], reverse=True)},
        'byType': {k: {'amount': round(v['amount'], 2), 'count': v['count']}
                   for k, v in sorted(by_type.items(), key=lambda x: x[1]['amount'], reverse=True)},
        'byMonth': {k: {'amount': round(v['amount'], 2), 'count': v['count']}
                    for k, v in sorted(by_month.items())},
        'byEmployee': {k: {'amount': round(v['amount'], 2), 'count': v['count']}
                       for k, v in sorted(by_employee.items(), key=lambda x: x[1]['amount'], reverse=True)[:100]},
        'bySource': {k: {'amount': round(v['amount'], 2), 'count': v['count']}
                     for k, v in sorted(by_source.items(), key=lambda x: x[1]['amount'], reverse=True)}
    }


def build_indexes(records: List[Dict]) -> Dict[str, Any]:
    """
    构建数据索引

    Args:
        records: 所有记录

    Returns:
        索引字典
    """
    dept_index = defaultdict(list)
    type_index = defaultdict(list)
    month_index = defaultdict(list)
    employee_index = defaultdict(list)
    source_index = defaultdict(list)

    for i, record in enumerate(records):
        dept = record.get('deptLevel1', '未知部门')
        record_type = record.get('type', 'unknown')
        date_str = parse_date_from_record(record)
        month = date_str[:7] if len(date_str) >= 7 else '未知月份'
        employee = get_employee_name(record) or '未知员工'
        source = record.get('source', '未知来源')

        dept_index[dept].append(i)
        type_index[record_type].append(i)
        month_index[month].append(i)
        employee_index[employee].append(i)
        source_index[source].append(i)

    return {
        'byDept': {k: v for k, v in sorted(dept_index.items())},
        'byType': {k: v for k, v in sorted(type_index.items())},
        'byMonth': {k: v for k, v in sorted(month_index.items())},
        'byEmployee': {k: v for k, v in sorted(employee_index.items())},
        'bySource': {k: v for k, v in sorted(source_index.items())}
    }


def merge_data(
    by_month_dir: Path,
    output_path: Path,
    roster_index_path: Path
) -> bool:
    """
    合并数据并生成完整的数据文件

    Args:
        by_month_dir: 按月分片数据目录
        output_path: 输出文件路径
        roster_index_path: 花名册索引文件路径

    Returns:
        是否成功
    """
    print('=' * 70)
    print('合并差旅数据')
    print('=' * 70)

    if not by_month_dir.exists():
        print(f'错误: 数据目录不存在: {by_month_dir}')
        return False

    # 合并按月数据
    print('\n扫描按月分片数据...')
    merged_data = merge_monthly_data(by_month_dir)

    if not merged_data['records']:
        print('警告: 没有找到任何记录')
        return False

    # 读取花名册索引
    roster_data = {}
    if roster_index_path.exists():
        try:
            with open(roster_index_path, 'r', encoding='utf-8') as f:
                roster_data = json.load(f)
            print(f'\n读取花名册索引: {len(roster_data.get("allEmployees", {}))} 名员工')
        except Exception as e:
            print(f'警告: 无法读取花名册索引: {e}')

    # 构建摘要
    print('\n构建统计摘要...')
    summary = build_summary(merged_data['records'])

    # 构建索引
    print('\n构建数据索引...')
    indexes = build_indexes(merged_data['records'])

    # 组装最终数据
    output_data = {
        'lastUpdate': datetime.now().isoformat(),
        'months': merged_data['months'],
        'sources': merged_data['sources'],
        'records': merged_data['records'],
        'summary': summary,
        'indexes': indexes,
        'roster': roster_data
    }

    # 保存
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f'\n保存合并数据到: {output_path}')
    print(f'  总记录数: {summary["totalRecords"]}')
    print(f'  总金额: ¥{summary["totalAmount"]:,.2f}')
    print(f'  月份数: {len(merged_data["months"])}')
    print(f'  部门数: {len(summary["byDept"])}')
    print(f'  数据源: {", ".join(merged_data["sources"])}')

    return True


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='合并差旅数据')
    parser.add_argument('-i', '--input', default='data/processed/by-month', help='按月分片数据目录')
    parser.add_argument('-o', '--output', default='data/processed/travel-data.json', help='输出文件路径')
    parser.add_argument('-r', '--roster', default='data/processed/roster_index.json', help='花名册索引文件')

    args = parser.parse_args()

    by_month_dir = Path(args.input)
    output_path = Path(args.output)
    roster_index_path = Path(args.roster)

    success = merge_data(by_month_dir, output_path, roster_index_path)

    if not success:
        sys.exit(1)
