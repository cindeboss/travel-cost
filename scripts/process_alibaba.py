#!/usr/bin/env python3
"""
阿里商旅数据处理脚本

读取阿里商旅Excel文件，提取机票、酒店、火车、用车等数据。

阿里商旅Excel格式特点：
- 第1-2行：标题行
- 第3行：列名（表头）
- 第4行：合计行（需跳过）
- 第5行起：实际数据
"""

import sys
import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime

# 添加父目录到路径以导入utils
sys.path.insert(0, str(Path(__file__).parent))
from utils import TravelFileInfo, load_employee_index


# 工作表名称映射
SHEET_MAPPING = {
    'flight': ['本期国内机票交易明细', '国际/中国港澳台机票交易明细', '本期国际|中国港澳台机票交易明细'],
    'hotel': ['国内酒店对账单', '国际|中国港澳台酒店对账单'],
    'train': ['本期国内商旅火车票交易明细'],
    'car': ['国内用车对账单', '国际|中国港澳台用车对账单']
}


def read_alibaba_sheet(filepath: Path, sheet_name: str) -> Optional[pd.DataFrame]:
    """
    读取阿里商旅工作表，跳过标题行和合计行

    Args:
        filepath: Excel文件路径
        sheet_name: 工作表名称

    Returns:
        处理后的DataFrame
    """
    try:
        # 读取原始数据，header=0表示第3行（索引2）作为列名
        # skiprows=[3]表示跳过第4行（索引3，即"合计"行）
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=2, skiprows=[3])
        return df
    except Exception as e:
        print(f'    警告: 无法读取工作表 {sheet_name}: {e}')
        return None


def extract_flight_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取机票记录

    阿里商旅机票列索引（基于header=2的索引）:
    3: 预订人, 5: 出行人, 14: 起飞日期, 15: 起飞时间, 18: 出发城市, 19: 到达城市,
    23: 航空公司, 24: 航班号, 25: 舱位, 26: 舱等, 35: 订单金额
    """
    # 使用索引获取数据
    if len(row) <= 35:
        return None

    # 出行人（索引5）或预订人（索引3）
    passenger = row.iloc[5] if len(row) > 5 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        passenger = row.iloc[3] if len(row) > 3 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        return None

    passenger = str(passenger).strip()

    # 数据验证：跳过无效记录
    # 1. 乘机人不能是日期格式
    if passenger.startswith('202') or passenger.startswith('20') or passenger.count('-') >= 2:
        return None
    # 2. 乘机人不能是"小计"、"合计"等关键词
    if passenger in ['小计', '合计', '总计', '汇总', '平均值']:
        return None
    # 3. 乘机人不能是纯数字
    if passenger.isdigit():
        return None

    dept_info = roster_index.get(passenger, {})

    # 起飞日期和 时间
    depart_date = str(row.iloc[14]) if len(row) > 14 and pd.notna(row.iloc[14]) else ''
    depart_time = str(row.iloc[15]) if len(row) > 15 and pd.notna(row.iloc[15]) else ''

    # 航班号
    flight_no = str(row.iloc[24]) if len(row) > 24 and pd.notna(row.iloc[24]) else ''

    # 出发地和目的地
    from_city = str(row.iloc[18]) if len(row) > 18 and pd.notna(row.iloc[18]) else ''
    to_city = str(row.iloc[19]) if len(row) > 19 and pd.notna(row.iloc[19]) else ''

    # 数据验证：必须有航班号、起飞时间、出发地和目的地
    if not flight_no or flight_no == 'nan' or not depart_date or depart_date == 'nan':
        return None
    if not from_city or from_city == 'nan' or not to_city or to_city == 'nan':
        return None

    # 订单金额（索引35）
    price = 0
    if len(row) > 35 and pd.notna(row.iloc[35]):
        try:
            price = float(row.iloc[35])
        except (ValueError, TypeError):
            pass

    record = {
        'source': '阿里商旅',
        'type': 'flight',
        'passenger': passenger,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'bookTime': str(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else depart_date,
        'flightNo': flight_no,
        'departTime': f'{depart_date} {depart_time}'.strip(),
        'fromCity': from_city,
        'toCity': to_city,
        'price': price,
        'cabinClass': str(row.iloc[26]) if len(row) > 26 and pd.notna(row.iloc[26]) else '',
        'airline': str(row.iloc[23]) if len(row) > 23 and pd.notna(row.iloc[23]) else ''
    }

    # 订单号（索引1）
    if len(row) > 1 and pd.notna(row.iloc[1]):
        record['orderNo'] = str(int(row.iloc[1])) if isinstance(row.iloc[1], (int, float)) else str(row.iloc[1])

    return record


def extract_hotel_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取酒店记录
    注意：阿里商旅的酒店数据格式可能不同，需要根据实际文件调整
    """
    # 酒店数据暂时跳过，因为需要查看实际数据结构
    return None


def extract_train_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取火车记录

    阿里商旅火车列索引（基于header=2的索引）:
    2: 预订人, 3: 出行人, 10: 发车日期, 11: 发车时间, 12: 到达日期, 13: 到达时间,
    14: 出发城市, 15: 到达城市, 16: 车次, 18: 座席, 24: 订单金额
    """
    if len(row) <= 24:
        return None

    # 出行人（索引3）或预订人（索引2）
    passenger = row.iloc[3] if len(row) > 3 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        passenger = row.iloc[2] if len(row) > 2 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        return None

    passenger = str(passenger).strip()
    dept_info = roster_index.get(passenger, {})

    # 发车日期和时间
    depart_date = str(row.iloc[10]) if len(row) > 10 and pd.notna(row.iloc[10]) else ''
    depart_time = str(row.iloc[11]) if len(row) > 11 and pd.notna(row.iloc[11]) else ''

    # 订单金额（索引24）
    price = 0
    if len(row) > 24 and pd.notna(row.iloc[24]):
        try:
            price = float(row.iloc[24])
        except (ValueError, TypeError):
            pass

    record = {
        'source': '阿里商旅',
        'type': 'train',
        'employee': passenger,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'trainNo': str(row.iloc[16]) if len(row) > 16 and pd.notna(row.iloc[16]) else '',
        'seat': str(row.iloc[18]) if len(row) > 18 and pd.notna(row.iloc[18]) else '',
        'departTime': f'{depart_date} {depart_time}'.strip(),
        'fromCity': str(row.iloc[14]) if len(row) > 14 and pd.notna(row.iloc[14]) else '',
        'toCity': str(row.iloc[15]) if len(row) > 15 and pd.notna(row.iloc[15]) else '',
        'price': price
    }

    # 订单号（索引1）
    if len(row) > 1 and pd.notna(row.iloc[1]):
        record['orderNo'] = str(int(row.iloc[1])) if isinstance(row.iloc[1], (int, float)) else str(row.iloc[1])

    return record


def extract_car_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取用车记录

    阿里商旅用车列索引（header=2后）:
    3: 预订人, 6: 出行人, 14: 出发日期, 15: 出发时间, 16: 到达日期, 17: 到达时间,
    18: 出发城市, 19: 出发地, 21: 到达城市, 22: 到达地, 25: 实际行驶公里数,
    32: 结算金额, 41: 服务方, 42: 供应商车型, 43: 平台车型
    """
    if len(row) <= 43:
        return None

    # 出行人（索引6）或预订人（索引3）
    passenger = row.iloc[6] if len(row) > 6 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        passenger = row.iloc[3] if len(row) > 3 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        return None

    # 跳过员工ID
    if str(passenger).startswith('EMP'):
        passenger = row.iloc[3] if len(row) > 3 else None
        if pd.isna(passenger) or str(passenger).strip() == '':
            return None

    passenger = str(passenger).strip()
    dept_info = roster_index.get(passenger, {})

    # 上车时间（出发日期+出发时间）
    pickup_date = str(row.iloc[14]) if len(row) > 14 and pd.notna(row.iloc[14]) else ''
    pickup_time = str(row.iloc[15]) if len(row) > 15 and pd.notna(row.iloc[15]) else ''

    # 下车时间（到达日期+到达时间）
    dropoff_date = str(row.iloc[16]) if len(row) > 16 and pd.notna(row.iloc[16]) else ''
    dropoff_time = str(row.iloc[17]) if len(row) > 17 and pd.notna(row.iloc[17]) else ''

    # 金额（索引32是结算金额）
    amount = 0
    if len(row) > 32 and pd.notna(row.iloc[32]):
        try:
            amount = float(row.iloc[32])
        except (ValueError, TypeError):
            pass

    # 里程（索引25）
    distance = 0
    if len(row) > 25 and pd.notna(row.iloc[25]):
        try:
            distance = float(row.iloc[25])
        except (ValueError, TypeError):
            pass

    # 用车类型（索引43平台车型，索引42供应商车型）
    car_type = str(row.iloc[43]) if len(row) > 43 and pd.notna(row.iloc[43]) else ''
    if not car_type:
        car_type = str(row.iloc[42]) if len(row) > 42 and pd.notna(row.iloc[42]) else ''

    record = {
        'source': '阿里商旅',
        'type': 'car',
        'passenger': passenger,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'pickupTime': f'{pickup_date} {pickup_time}'.strip(),
        'dropoffTime': f'{dropoff_date} {dropoff_time}'.strip(),
        'carType': car_type,
        'provider': str(row.iloc[41]) if len(row) > 41 and pd.notna(row.iloc[41]) else '',
        'origin': {
            'city': str(row.iloc[18]) if len(row) > 18 and pd.notna(row.iloc[18]) else '',
            'address': str(row.iloc[19]) if len(row) > 19 and pd.notna(row.iloc[19]) else ''
        },
        'destination': {
            'city': str(row.iloc[21]) if len(row) > 21 and pd.notna(row.iloc[21]) else '',
            'address': str(row.iloc[22]) if len(row) > 22 and pd.notna(row.iloc[22]) else ''
        },
        'distance': distance,
        'totalAmount': amount
    }

    # 订单号（索引1）
    if len(row) > 1 and pd.notna(row.iloc[1]):
        record['orderNo'] = str(int(row.iloc[1])) if isinstance(row.iloc[1], (int, float)) else str(row.iloc[1])

    return record


def process_alibaba_file(
    filepath: Path,
    roster_index: Dict[str, Dict]
) -> List[Dict]:
    """
    处理阿里商旅Excel文件

    Args:
        filepath: Excel文件路径
        roster_index: 员工索引

    Returns:
        处理后的记录列表
    """
    print(f'\n处理阿里商旅文件: {filepath.name}')

    all_records = []

    try:
        excel_file = pd.ExcelFile(filepath)

        # 处理机票
        for sheet_pattern in SHEET_MAPPING['flight']:
            if sheet_pattern in excel_file.sheet_names:
                df = read_alibaba_sheet(filepath, sheet_pattern)
                if df is not None and len(df) > 0:
                    print(f'  处理机票数据 ({sheet_pattern}): {len(df)} 条')
                    for _, row in df.iterrows():
                        record = extract_flight_record(row, roster_index)
                        if record:
                            all_records.append(record)

        # 处理火车
        for sheet_pattern in SHEET_MAPPING['train']:
            if sheet_pattern in excel_file.sheet_names:
                df = read_alibaba_sheet(filepath, sheet_pattern)
                if df is not None and len(df) > 0:
                    print(f'  处理火车数据 ({sheet_pattern}): {len(df)} 条')
                    for _, row in df.iterrows():
                        record = extract_train_record(row, roster_index)
                        if record:
                            all_records.append(record)

        # 处理用车
        for sheet_pattern in SHEET_MAPPING['car']:
            if sheet_pattern in excel_file.sheet_names:
                df = read_alibaba_sheet(filepath, sheet_pattern)
                if df is not None and len(df) > 0:
                    print(f'  处理用车数据 ({sheet_pattern}): {len(df)} 条')
                    for _, row in df.iterrows():
                        record = extract_car_record(row, roster_index)
                        if record:
                            all_records.append(record)

    except Exception as e:
        print(f'  错误: 无法读取Excel文件: {e}')
        import traceback
        traceback.print_exc()
        return []

    print(f'  总共提取 {len(all_records)} 条记录')
    return all_records


def process_alibaba(
    filepath: Path,
    output_dir: Path,
    roster_index_path: Path,
    file_info: Optional[TravelFileInfo] = None
) -> Optional[str]:
    """
    处理阿里商旅文件并保存结果
    """
    # 确定月份
    month = None
    if file_info and file_info.target_month:
        month = file_info.target_month
    else:
        from utils import parse_date_range_from_filename
        date_range = parse_date_range_from_filename(filepath.name)
        if date_range:
            month = date_range.main_month

    if not month:
        print(f'警告: 无法确定文件 {filepath.name} 的归属月份')
        return None

    # 加载员工索引
    roster_index = load_employee_index(roster_index_path, month)
    if not roster_index:
        print(f'警告: 未找到月份 {month} 的花名册数据')

    # 处理文件
    records = process_alibaba_file(filepath, roster_index)
    if not records:
        return month

    # 保存结果
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f'alibaba_{month}.json'

    output_data = {
        'source': '阿里商旅',
        'month': month,
        'sourceFile': filepath.name,
        'processedAt': datetime.now().isoformat(),
        'records': records,
        'count': len(records)
    }

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f'  保存到: {output_file}')

    # 统计关联率
    matched_count = sum(1 for r in records if r.get('deptLevel1'))
    if records:
        match_rate = matched_count / len(records) * 100
        print(f'  部门关联率: {matched_count}/{len(records)} ({match_rate:.1f}%)')

    return month


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='处理阿里商旅文件')
    parser.add_argument('input', help='输入的Excel文件路径')
    parser.add_argument('-o', '--output', default='data/processed', help='输出目录')
    parser.add_argument('-r', '--roster', default='data/processed/roster_index.json', help='花名册索引文件')

    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)
    roster_index_path = Path(args.roster)

    process_alibaba(input_path, output_dir / 'by-month', roster_index_path)
