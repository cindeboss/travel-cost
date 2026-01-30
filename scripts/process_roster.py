#!/usr/bin/env python3
"""
花名册处理脚本

读取花名册Excel文件，提取在职员工信息，并关联部门信息。
"""

import sys
import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime

# 添加父目录到路径以导入utils
sys.path.insert(0, str(Path(__file__).parent))
from utils import extract_roster_month


def process_roster_file(filepath: Path) -> List[Dict]:
    """
    处理单个花名册文件

    Args:
        filepath: 花名册Excel文件路径

    Returns:
        员工记录列表
    """
    # 读取"原表"工作表
    try:
        df = pd.read_excel(filepath, sheet_name='原表')
    except Exception as e:
        print(f'错误: 无法读取工作表"原表": {e}')
        # 尝试读取第一个工作表
        try:
            df = pd.read_excel(filepath, sheet_name=0)
            print(f'警告: 使用第一个工作表代替"原表"')
        except Exception as e2:
            print(f'错误: 无法读取Excel文件: {e2}')
            return []

    # 标准化列名
    df.columns = df.columns.str.strip()

    # 检查必需的列
    required_columns = ['姓名', '一级部门', '在职状态']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print(f'警告: 缺少必需的列: {missing_columns}')
        print(f'可用的列: {list(df.columns)}')
        return []

    # 过滤在职员工
    # 常见的在职状态值: 在职, 试用期, 正式, 等
    active_statuses = ['在职', '试用期', '正式', '实习', 'contractor']
    df_active = df[df['在职状态'].isin(active_statuses) | df['在职状态'].str.contains('在职', na=False)]

    if len(df_active) == 0:
        print('警告: 没有找到在职员工')
        return []

    # 提取字段
    records = []
    for _, row in df_active.iterrows():
        record = {
            'name': str(row.get('姓名', '')).strip(),
            'englishName': str(row.get('英文名', '')).strip() if pd.notna(row.get('英文名')) else '',
            'deptLevel1': str(row.get('一级部门', '')).strip(),
            'deptLevel2': str(row.get('二级部门', '')).strip() if pd.notna(row.get('二级部门')) else '',
            'deptLevel3': str(row.get('三级部门', '')).strip() if pd.notna(row.get('三级部门')) else '',
            'position': str(row.get('岗位', '')).strip() if pd.notna(row.get('岗位')) else '',
            'status': str(row.get('在职状态', '')).strip()
        }

        # 只保留有姓名的记录
        if record['name'] and record['name'] != 'nan':
            records.append(record)

    return records


def build_employee_index(records: List[Dict]) -> Dict[str, Dict]:
    """
    构建员工索引

    Args:
        records: 员工记录列表

    Returns:
        员工索引字典 {姓名: 部门信息}
    """
    index = {}
    for record in records:
        name = record['name']
        index[name] = {
            'deptLevel1': record['deptLevel1'],
            'deptLevel2': record['deptLevel2'],
            'deptLevel3': record['deptLevel3'],
            'position': record['position'],
            'status': record['status']
        }
    return index


def process_roster(
    filepath: Path,
    output_dir: Path,
    roster_index_path: Path
) -> Optional[str]:
    """
    处理花名册文件并保存结果

    Args:
        filepath: 花名册Excel文件路径
        output_dir: 输出目录
        roster_index_path: 花名册索引文件路径

    Returns:
        处理的月份 (YYYY-MM格式)，如果失败返回None
    """
    print(f'\n处理花名册: {filepath.name}')

    # 提取月份
    month = extract_roster_month(filepath.name)
    if not month:
        print(f'错误: 无法从文件名提取月份: {filepath.name}')
        return None

    # 处理花名册
    records = process_roster_file(filepath)
    if not records:
        print(f'警告: 花名册没有有效记录: {filepath.name}')
        return month

    print(f'  提取了 {len(records)} 条在职员工记录')

    # 构建员工索引
    employee_index = build_employee_index(records)

    # 保存按月分片的数据
    output_dir.mkdir(parents=True, exist_ok=True)
    month_file = output_dir / f'{month}.json'

    month_data = {
        'month': month,
        'rosterFile': filepath.name,
        'processedAt': datetime.now().isoformat(),
        'employees': employee_index,
        'count': len(employee_index)
    }

    with open(month_file, 'w', encoding='utf-8') as f:
        json.dump(month_data, f, ensure_ascii=False, indent=2)

    print(f'  保存到: {month_file}')

    # 更新花名册索引
    update_roster_index(roster_index_path, month, filepath.name, employee_index)

    return month


def update_roster_index(
    index_path: Path,
    month: str,
    filename: str,
    employee_index: Dict[str, Dict]
):
    """
    更新花名册索引

    Args:
        index_path: 索引文件路径
        month: 月份
        filename: 文件名
        employee_index: 员工索引
    """
    # 读取现有索引
    index = {}
    if index_path.exists():
        try:
            with open(index_path, 'r', encoding='utf-8') as f:
                index = json.load(f)
        except Exception as e:
            print(f'警告: 无法读取现有索引，将创建新索引: {e}')

    # 更新月份信息
    if 'months' not in index:
        index['months'] = {}

    index['months'][month] = {
        'file': filename,
        'processedAt': datetime.now().isoformat(),
        'count': len(employee_index)
    }

    # 更新全局员工索引
    if 'allEmployees' not in index:
        index['allEmployees'] = {}

    for name, info in employee_index.items():
        if name not in index['allEmployees']:
            index['allEmployees'][name] = {
                'deptLevel1': info['deptLevel1'],
                'deptLevel2': info['deptLevel2'],
                'deptLevel3': info['deptLevel3'],
                'latestRecord': month
            }
        else:
            # 更新最新记录
            if month >= index['allEmployees'][name].get('latestRecord', ''):
                index['allEmployees'][name]['latestRecord'] = month
                # 也更新部门信息（可能发生调岗）
                index['allEmployees'][name]['deptLevel1'] = info['deptLevel1']
                index['allEmployees'][name]['deptLevel2'] = info['deptLevel2']
                index['allEmployees'][name]['deptLevel3'] = info['deptLevel3']

    # 保存索引
    index_path.parent.mkdir(parents=True, exist_ok=True)
    with open(index_path, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)

    print(f'  更新索引: {index_path}')


def load_employee_index(roster_index_path: Path, month: Optional[str] = None) -> Dict[str, Dict]:
    """
    加载员工索引

    Args:
        roster_index_path: 花名册索引文件路径
        month: 指定月份，如果为None则使用最新月份

    Returns:
        员工索引字典
    """
    if not roster_index_path.exists():
        return {}

    with open(roster_index_path, 'r', encoding='utf-8') as f:
        index = json.load(f)

    # 如果指定了月份，从月度文件加载
    if month:
        month_file = roster_index_path.parent / 'by-month' / f'{month}.json'
        if month_file.exists():
            with open(month_file, 'r', encoding='utf-8') as f:
                month_data = json.load(f)
            return month_data.get('employees', {})

    # 否则返回全局索引
    return index.get('allEmployees', {})


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='处理花名册文件')
    parser.add_argument('input', help='输入的Excel文件路径')
    parser.add_argument('-o', '--output', default='data/processed', help='输出目录')

    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)
    roster_index_path = output_dir / 'roster_index.json'

    process_roster(input_path, output_dir / 'by-month', roster_index_path)
