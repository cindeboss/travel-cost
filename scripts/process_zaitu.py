#!/usr/bin/env python3
"""
在途商旅数据处理脚本

读取在途商旅.xls文件，提取各类差旅数据。
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


# 在途工作表名称映射（可能需要根据实际文件调整）
ZAITU_SHEET_MAPPING = {
    'flight': ['机票', '机票交易', '机票明细'],
    'hotel': ['酒店', '酒店交易', '酒店明细', '住宿'],
    'train': ['火车', '火车票', '火车交易', '火车明细'],
    'car': ['用车', '用车交易', '用车明细', '用车服务', '增值服务']
}


def find_zaitu_worksheet(df_dict: Dict[str, pd.DataFrame], sheet_type: str) -> Optional[pd.DataFrame]:
    """
    根据类型查找在途工作表

    Args:
        df_dict: 所有工作表的字典
        sheet_type: 工作表类型

    Returns:
        找到的DataFrame，如果没找到返回None
    """
    possible_names = ZAITU_SHEET_MAPPING.get(sheet_type, [])

    for name in possible_names:
        if name in df_dict:
            return df_dict[name]

    # 模糊匹配
    for key, df in df_dict.items():
        for pattern in possible_names:
            if pattern in key or key in pattern:
                return df

    return None


def extract_zaitu_flight_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取在途机票记录

    在途商旅机票列索引:
    6: 预订人, 8: 航班号, 13: 起飞时间, 15: 航程(出发地-目的地), 17: 乘机人,
    23: 票价, 27: 小计, 31: 总额, 32: 服务费, 33: 合计
    正确的金额是 合计(33)
    正确的城市是 航程(15)，需要按"-"分割
    """
    if len(row) <= 33:
        return None

    # 乘机人（索引17）
    passenger = row.iloc[17] if len(row) > 17 else None
    if pd.isna(passenger) or str(passenger).strip() == '':
        passenger = row.iloc[6] if len(row) > 6 else None
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

    # 起飞时间（索引13）
    depart_time = str(row.iloc[13]) if len(row) > 13 and pd.notna(row.iloc[13]) else ''

    # 航班号（索引8）
    flight_no = str(row.iloc[8]) if len(row) > 8 and pd.notna(row.iloc[8]) else ''

    # 解析航程（索引15）获取出发地和目的地
    route = str(row.iloc[15]) if len(row) > 15 and pd.notna(row.iloc[15]) else ''
    from_city = ''
    to_city = ''
    if route and '-' in route:
        parts = route.split('-')
        if len(parts) >= 2:
            from_city = parts[0].strip()
            to_city = parts[1].strip()

    # 正确的金额是"合计"（索引33）
    price = 0
    if len(row) > 33 and pd.notna(row.iloc[33]):
        try:
            price = float(row.iloc[33])
        except (ValueError, TypeError):
            # 如果合计无效，尝试总额（索引31）
            if len(row) > 31 and pd.notna(row.iloc[31]):
                try:
                    price = float(row.iloc[31])
                except (ValueError, TypeError):
                    pass

    record = {
        'source': '在途商旅',
        'type': 'flight',
        'passenger': passenger,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'bookTime': str(row.iloc[5]) if len(row) > 5 and pd.notna(row.iloc[5]) else '',
        'flightNo': flight_no,
        'departTime': depart_time,
        'fromCity': from_city,
        'toCity': to_city,
        'price': price,
        'cabinClass': str(row.iloc[9]) if len(row) > 9 and pd.notna(row.iloc[9]) else '',
        'airline': str(row.iloc[7]) if len(row) > 7 and pd.notna(row.iloc[7]) else ''
    }

    # 订单号（索引1）
    if len(row) > 1 and pd.notna(row.iloc[1]):
        record['orderNo'] = str(row.iloc[1])

    return record


def extract_zaitu_hotel_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取在途酒店记录

    在途商旅酒店列索引:
    6: 预订人, 13: 入住人, 16: 总额, 17: 税额, 18: 服务费, 19: 合计
    正确的金额是 总额(16) 或 合计(19)
    """
    if len(row) <= 19:
        return None

    # 入住人（索引13）
    employee = row.iloc[13] if len(row) > 13 else None
    if pd.isna(employee) or str(employee).strip() == '':
        employee = row.iloc[6] if len(row) > 6 else None
    if pd.isna(employee) or str(employee).strip() == '':
        return None

    employee = str(employee).strip()
    dept_info = roster_index.get(employee, {})

    # 正确的金额是"总额"（索引16）或"合计"（索引19）
    price = 0
    if len(row) > 16 and pd.notna(row.iloc[16]):
        try:
            price = float(row.iloc[16])
        except (ValueError, TypeError):
            # 如果总额无效，尝试合计（索引19）
            if len(row) > 19 and pd.notna(row.iloc[19]):
                try:
                    price = float(row.iloc[19])
                except (ValueError, TypeError):
                    pass

    record = {
        'source': '在途商旅',
        'type': 'hotel',
        'employee': employee,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'checkInTime': str(row.iloc[10]) if len(row) > 10 and pd.notna(row.iloc[10]) else '',
        'checkOutTime': str(row.iloc[11]) if len(row) > 11 and pd.notna(row.iloc[11]) else '',
        'city': str(row.iloc[7]) if len(row) > 7 and pd.notna(row.iloc[7]) else '',
        'district': '',
        'hotelName': str(row.iloc[8]) if len(row) > 8 and pd.notna(row.iloc[8]) else '',
        'roomType': str(row.iloc[9]) if len(row) > 9 and pd.notna(row.iloc[9]) else '',
        'price': price,
        'isShared': False,
        'personalOverage': 0,
        'miscFee': 0
    }

    # 订单号（索引1）
    if len(row) > 1 and pd.notna(row.iloc[1]):
        record['orderNo'] = str(row.iloc[1])

    return record


def extract_zaitu_train_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取在途火车记录

    在途商旅火车列索引:
    6: 预订人, 7: 车次号, 8: 座席, 9: 出发城市, 10: 到达城市, 11: 发车时间
    19: 乘车人, 24: 小计, 26: 合计
    正确的金额是 小计(24) 或 合计(26)
    """
    if len(row) <= 26:
        return None

    # 乘车人（索引19）或预订人（索引6）
    employee = row.iloc[19] if len(row) > 19 else None
    if pd.isna(employee) or str(employee).strip() == '':
        employee = row.iloc[6] if len(row) > 6 else None
    if pd.isna(employee) or str(employee).strip() == '':
        return None

    employee = str(employee).strip()
    dept_info = roster_index.get(employee, {})

    # 正确的金额是"小计"（索引24）或"合计"（索引26）
    price = 0
    if len(row) > 24 and pd.notna(row.iloc[24]):
        try:
            price = float(row.iloc[24])
        except (ValueError, TypeError):
            # 如果小计无效，尝试合计（索引26）
            if len(row) > 26 and pd.notna(row.iloc[26]):
                try:
                    price = float(row.iloc[26])
                except (ValueError, TypeError):
                    pass

    # 过滤价格为0的记录
    if price == 0:
        return None

    record = {
        'source': '在途商旅',
        'type': 'train',
        'employee': employee,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'trainNo': str(row.iloc[7]) if len(row) > 7 and pd.notna(row.iloc[7]) else '',
        'seat': str(row.iloc[8]) if len(row) > 8 and pd.notna(row.iloc[8]) else '',
        'departTime': str(row.iloc[11]) if len(row) > 11 and pd.notna(row.iloc[11]) else '',
        'fromCity': str(row.iloc[9]) if len(row) > 9 and pd.notna(row.iloc[9]) else '',
        'toCity': str(row.iloc[10]) if len(row) > 10 and pd.notna(row.iloc[10]) else '',
        'price': price
    }

    # 订单号（索引1）
    if len(row) > 1 and pd.notna(row.iloc[1]):
        record['orderNo'] = str(row.iloc[1])

    return record


def extract_zaitu_car_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取在途用车记录

    在途商旅用车使用字典方式获取数据，但里程、服务方、金额使用特定索引
    """
    # 乘车人
    passenger = row.get('乘车人') or row.get('姓名') or row.get('员工姓名')
    if not passenger or pd.isna(passenger):
        return None

    passenger = str(passenger).strip()
    dept_info = roster_index.get(passenger, {})

    # 用索引获取的关键字段
    # 服务方（索引19）
    provider = ''
    if len(row) > 19 and pd.notna(row.iloc[19]):
        provider = str(row.iloc[19]).strip()

    # 行驶公里数（索引29）
    distance = 0
    if len(row) > 29 and pd.notna(row.iloc[29]):
        try:
            distance = float(row.iloc[29])
        except (ValueError, TypeError):
            distance = 0

    # 订单总金额（索引38）
    total_amount = 0
    if len(row) > 38 and pd.notna(row.iloc[38]):
        try:
            total_amount = float(row.iloc[38])
        except (ValueError, TypeError):
            total_amount = 0

    # 用字典方式获取其他字段
    # 获取时间字段
    pickup_time = str(row.get('上车时间', '') or row.get('开始时间', '') or row.get('用车开始时间', '')).strip() if pd.notna(row.get('上车时间') or row.get('开始时间') or row.get('用车开始时间')) else ''
    dropoff_time = str(row.get('下车时间', '') or row.get('结束时间', '') or row.get('用车结束时间', '')).strip() if pd.notna(row.get('下车时间') or row.get('结束时间') or row.get('用车结束时间')) else ''

    # 过滤时间为空的记录
    if not pickup_time:
        return None

    record = {
        'source': '在途商旅',
        'type': 'car',
        'passenger': passenger,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'pickupTime': pickup_time,
        'dropoffTime': dropoff_time,
        'carType': str(row.get('用车类型', '') or row.get('车型', '')).strip() if pd.notna(row.get('用车类型') or row.get('车型')) else '',
        'provider': provider,
        'origin': {
            'city': str(row.get('出发地（城市/区县/具体地址）', '') or row.get('出发地', '') or row.get('上车地点', '') or row.get('起点', '')).strip(),
        },
        'destination': {
            'city': str(row.get('目的地（城市/区县/具体地址）', '') or row.get('目的地', '') or row.get('下车地点', '') or row.get('终点', '')).strip(),
        },
        'distance': distance,
        'totalAmount': total_amount
    }

    # 解析出发地详细地址
    origin_raw = str(row.get('出发地（城市/区县/具体地址）', '') or row.get('出发地', '') or row.get('上车地点', '') or row.get('起点', '')).strip()
    if origin_raw:
        parts = origin_raw.split('/')
        if len(parts) >= 1:
            record['origin']['city'] = parts[0].strip()
        if len(parts) >= 2:
            record['origin']['district'] = parts[1].strip()
        if len(parts) >= 3:
            record['origin']['address'] = parts[2].strip()

    # 解析目的地详细地址
    dest_raw = str(row.get('目的地（城市/区县/具体地址）', '') or row.get('目的地', '') or row.get('下车地点', '') or row.get('终点', '')).strip()
    if dest_raw:
        parts = dest_raw.split('/')
        if len(parts) >= 1:
            record['destination']['city'] = parts[0].strip()
        if len(parts) >= 2:
            record['destination']['district'] = parts[1].strip()
        if len(parts) >= 3:
            record['destination']['address'] = parts[2].strip()

    if '订单号' in row and pd.notna(row['订单号']):
        record['orderNo'] = str(row['订单号'])

    return record


def process_zaitu_file(
    filepath: Path,
    roster_index: Dict[str, Dict]
) -> List[Dict]:
    """
    处理在途商旅Excel文件

    Args:
        filepath: Excel文件路径
        roster_index: 员工索引

    Returns:
        处理后的记录列表
    """
    print(f'\n处理在途商旅文件: {filepath.name}')

    all_records = []

    try:
        # 在途文件可能是.xls格式，使用xlrd引擎
        try:
            # 读取所有工作表
            excel_file = pd.ExcelFile(filepath, engine='xlrd')
        except Exception:
            # 如果xlrd失败，尝试openpyxl
            excel_file = pd.ExcelFile(filepath)

        sheet_dict = {name: pd.read_excel(filepath, sheet_name=name, engine='xlrd' if filepath.suffix == '.xls' else None)
                     for name in excel_file.sheet_names}

        # 处理机票
        flight_df = find_zaitu_worksheet(sheet_dict, 'flight')
        if flight_df is not None:
            print(f'  处理机票数据: {len(flight_df)} 条')
            for _, row in flight_df.iterrows():
                record = extract_zaitu_flight_record(row, roster_index)
                if record:
                    all_records.append(record)

        # 处理酒店
        hotel_df = find_zaitu_worksheet(sheet_dict, 'hotel')
        if hotel_df is not None:
            print(f'  处理酒店数据: {len(hotel_df)} 条')
            for _, row in hotel_df.iterrows():
                record = extract_zaitu_hotel_record(row, roster_index)
                if record:
                    all_records.append(record)

        # 处理火车
        train_df = find_zaitu_worksheet(sheet_dict, 'train')
        if train_df is not None:
            print(f'  处理火车数据: {len(train_df)} 条')
            for _, row in train_df.iterrows():
                record = extract_zaitu_train_record(row, roster_index)
                if record:
                    all_records.append(record)

        # 处理用车
        car_df = find_zaitu_worksheet(sheet_dict, 'car')
        if car_df is not None:
            print(f'  处理用车数据: {len(car_df)} 条')
            for _, row in car_df.iterrows():
                record = extract_zaitu_car_record(row, roster_index)
                if record:
                    all_records.append(record)

    except Exception as e:
        print(f'  错误: 无法读取Excel文件: {e}')
        import traceback
        traceback.print_exc()
        return []

    print(f'  总共提取 {len(all_records)} 条记录')

    return all_records


def process_zaitu(
    filepath: Path,
    output_dir: Path,
    roster_index_path: Path,
    file_info: Optional[TravelFileInfo] = None
) -> Optional[str]:
    """
    处理在途商旅文件并保存结果

    Args:
        filepath: Excel文件路径
        output_dir: 输出目录
        roster_index_path: 花名册索引文件路径
        file_info: 文件信息（如果已有）

    Returns:
        处理的月份 (YYYY-MM格式)，如果失败返回None
    """
    # 确定月份
    month = None
    if file_info and file_info.target_month:
        month = file_info.target_month
    else:
        # 从文件名提取月份
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
    records = process_zaitu_file(filepath, roster_index)

    if not records:
        return month

    # 保存结果
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f'zaitu_{month}.json'

    output_data = {
        'source': '在途商旅',
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

    parser = argparse.ArgumentParser(description='处理在途商旅文件')
    parser.add_argument('input', help='输入的Excel文件路径')
    parser.add_argument('-o', '--output', default='data/processed', help='输出目录')
    parser.add_argument('-r', '--roster', default='data/processed/roster_index.json', help='花名册索引文件')

    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)
    roster_index_path = Path(args.roster)

    process_zaitu(input_path, output_dir / 'by-month', roster_index_path)
