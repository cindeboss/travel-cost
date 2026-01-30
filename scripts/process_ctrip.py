#!/usr/bin/env python3
"""
携程商旅数据处理脚本

读取携程商旅Excel文件，提取机票、酒店、用车等数据。

携程商旅Excel格式特点：
- 第1-4行：标题行
- 第5行：中文列名（表头）
- 第6行：英文列名（表头）
- 第7行起：实际数据
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


# 航空公司代码映射（从航班号前缀推断）
AIRLINE_CODE_MAP = {
    'CA': '中国国航',
    'MU': '中国东方航空',
    'CZ': '中国南方航空',
    '3U': '四川航空',
    'ZH': '深圳航空',
    'HU': '海南航空',
    'FM': '上海航空',
    'KN': '中国联合航空',
    'JD': '金鹏航空',
    'MF': '厦门航空',
    'SC': '山东航空',
    'HO': '吉祥航空',
    'OQ': '重庆航空',
    'TV': '西藏航空',
    '8L': '祥鹏航空',
    'AQ': '九元航空',
    'GS': '天津航空',
    'NS': '河北航空',
    'BK': '奥凯航空',
    'EU': '成都航空',
    'GY': '多彩贵州航空',
    'DZ': '深圳东海航空',
    'G5': '华夏航空',
    'PN': '西部航空',
    'QW': '青岛航空',
    'RY': '江西航空',
    'UQ': '乌鲁木齐航空',
    'Y8': '扬子江航空',
    'GY': '桂林航空',
    'DR': '瑞丽航空',
    'A6': '红土航空',
    'GT': '大新华航空',
    'VN': '越南航空',
    'GJ': '长龙航空',
    '9C': '春秋航空',
    'CX': '国泰航空',
    'KA': '国泰港龙航空',
    'HX': '香港航空',
    'UO': '香港快运航空'
}


# 舱位代码映射
CABIN_CODE_MAP = {
    'Y': '经济舱', 'B': '经济舱', 'M': '经济舱', 'H': '经济舱',
    'K': '经济舱', 'L': '经济舱', 'Q': '经济舱', 'E': '经济舱',
    'C': '商务舱', 'D': '商务舱', 'J': '商务舱',
    'F': '头等舱', 'A': '头等舱', 'P': '经济舱',
    'W': '经济舱', 'S': '经济舱', 'T': '经济舱'
}


def infer_airline_from_flight_no(flight_no: str) -> str:
    """从航班号推断航空公司"""
    if not flight_no or len(flight_no) < 2:
        return ''
    code = flight_no[:2].upper()
    return AIRLINE_CODE_MAP.get(code, '')


def convert_cabin_code(cabin_code: str) -> str:
    """将舱位代码转换为中文"""
    if not cabin_code:
        return ''
    code = cabin_code.strip().upper()
    return CABIN_CODE_MAP.get(code, cabin_code)


# 携程工作表名称映射
CTRIP_SHEET_MAPPING = {
    'flight': ['预存机票'],
    'hotel': ['预存会员酒店', '预存协议酒店'],
    'car': ['预存增值']
}


def read_ctrip_sheet(filepath: Path, sheet_name: str) -> Optional[pd.DataFrame]:
    """
    读取携程商旅工作表，跳过标题行

    Args:
        filepath: Excel文件路径
        sheet_name: 工作表名称

    Returns:
        处理后的DataFrame
    """
    try:
        # 携程文件有两行表头（中文+英文），使用header=5
        # Row 0-3: 标题行
        # Row 4: 中文列名
        # Row 5: 英文列名
        # Row 6+: 实际数据
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=5)
        # 过滤掉明显是列名的行（第一列包含中文订单号或英文OrderID）
        if len(df) > 0:
            df = df[~df.iloc[:, 0].astype(str).str.contains('订单号|OrderID', na=False)]
        return df
    except Exception as e:
        print(f'    警告: 无法读取工作表 {sheet_name}: {e}')
        return None


def extract_ctrip_flight_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取携程机票记录

    携程机票列索引（基于header=5，英文列名行）:
    0: 订单号(OrderID), 5: 乘机人(PassengerName), 6: 预订日期(OrderDate), 7: 起飞时间(TakeOffTime),
    11: 航程(OrderDesc), 12: 航班号(Flight), 13: 舱等(Class), 14: 价格(price)
    """
    if len(row) <= 14:
        return None

    # 乘机人（索引5）
    passenger = row.iloc[5]
    if pd.isna(passenger) or str(passenger).strip() == '':
        return None

    passenger = str(passenger).strip()

    # 数据验证：跳过无效记录
    # 1. 乘机人不能是日期格式
    if passenger.startswith('202') or passenger.startswith('20') or passenger.count('-') >= 2:
        return None
    # 2. 乘机人不能是"小计"、"合计"等关键词
    if passenger in ['小计', '合计', '总计', '汇总', '平均值', 'ExchangeTime', 'ETD', 'PassengerName']:
        return None
    # 3. 乘机人不能是纯数字
    if passenger.isdigit():
        return None

    dept_info = roster_index.get(passenger, {})

    # 起飞时间（索引7）
    depart_time = str(row.iloc[7]) if len(row) > 7 and pd.notna(row.iloc[7]) else ''

    # 航程（索引11，格式如"深圳-北京"）
    route = str(row.iloc[11]) if len(row) > 11 and pd.notna(row.iloc[11]) else ''
    from_city = ''
    to_city = ''
    if '-' in route:
        cities = route.split('-')
        from_city = cities[0].strip() if len(cities) > 0 else ''
        to_city = cities[1].strip() if len(cities) > 1 else ''

    # 航班号（索引12）
    flight_no = str(row.iloc[12]) if len(row) > 12 and pd.notna(row.iloc[12]) else ''

    # 数据验证：必须有起飞时间、航班号和航程
    if not depart_time or depart_time == 'nan' or not flight_no or flight_no == 'nan':
        return None

    # 从航班号推断航空公司
    airline = infer_airline_from_flight_no(flight_no)

    # 价格（索引14）
    price = 0
    if len(row) > 14 and pd.notna(row.iloc[14]):
        try:
            price = float(row.iloc[14])
        except (ValueError, TypeError):
            pass

    # 过滤价格为0的记录（可能是异常数据）
    if price == 0:
        return None

    # 舱位（索引13），转换代码为中文
    cabin_raw = str(row.iloc[13]) if len(row) > 13 and pd.notna(row.iloc[13]) else ''
    cabin_class = convert_cabin_code(cabin_raw)

    record = {
        'source': '携程商旅',
        'type': 'flight',
        'passenger': passenger,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'bookTime': str(row.iloc[6]) if len(row) > 6 and pd.notna(row.iloc[6]) else depart_time,
        'flightNo': flight_no,
        'departTime': depart_time,
        'fromCity': from_city,
        'toCity': to_city,
        'price': price,  # 保留原始价格，退票为负数
        'cabinClass': cabin_class,
        'airline': airline
    }

    # 订单号（索引0）
    if len(row) > 0 and pd.notna(row.iloc[0]):
        record['orderNo'] = str(row.iloc[0])

    return record


def extract_ctrip_hotel_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取携程酒店记录

    携程酒店列索引（基于header=5，英文列名行）:
    0: 订单号(OrderID), 4: 入住人(clients), 6: 预订日期(OrderDate), 7: 入住日期(ETA), 8: 离店日期(ETD),
    9: 酒店城市(city), 10: 酒店名称(HotelName), 12: 房型(roomname), 14: 单价(Price), 18: 金额(Amount)
    """
    if len(row) <= 18:
        return None

    # 入住人（索引4）
    employee = row.iloc[4]
    if pd.isna(employee) or str(employee).strip() == '':
        return None

    employee = str(employee).strip()

    # 数据验证：跳过无效记录
    if employee in ['clients', '入住人', '小计', '合计', '总计']:
        return None

    dept_info = roster_index.get(employee, {})

    # 金额（索引18）
    price = 0
    if len(row) > 18 and pd.notna(row.iloc[18]):
        try:
            price = float(row.iloc[18])
        except (ValueError, TypeError):
            pass

    record = {
        'source': '携程商旅',
        'type': 'hotel',
        'employee': employee,
        'deptLevel1': dept_info.get('deptLevel1', ''),
        'deptLevel2': dept_info.get('deptLevel2', ''),
        'checkInTime': str(row.iloc[7]) if len(row) > 7 and pd.notna(row.iloc[7]) else '',
        'checkOutTime': str(row.iloc[8]) if len(row) > 8 and pd.notna(row.iloc[8]) else '',
        'city': str(row.iloc[9]) if len(row) > 9 and pd.notna(row.iloc[9]) else '',
        'district': '',
        'hotelName': str(row.iloc[10]) if len(row) > 10 and pd.notna(row.iloc[10]) else '',
        'roomType': str(row.iloc[12]) if len(row) > 12 and pd.notna(row.iloc[12]) else '',
        'price': abs(price),
        'isShared': False,
        'personalOverage': 0,
        'miscFee': 0
    }

    # 订单号（索引0）
    if len(row) > 0 and pd.notna(row.iloc[0]):
        record['orderNo'] = str(row.iloc[0])

    return record


def extract_ctrip_car_record(row: pd.Series, roster_index: Dict[str, Dict]) -> Optional[Dict]:
    """
    提取携程用车记录（预存增值）
    注意：携程的用车数据可能在"预存增值"中，需要根据实际数据结构调整
    """
    # 用车数据暂时跳过，因为需要查看实际数据结构
    return None


def process_ctrip_file(
    filepath: Path,
    roster_index: Dict[str, Dict]
) -> List[Dict]:
    """
    处理携程商旅Excel文件

    Args:
        filepath: Excel文件路径
        roster_index: 员工索引

    Returns:
        处理后的记录列表
    """
    print(f'\n处理携程商旅文件: {filepath.name}')

    all_records = []

    try:
        excel_file = pd.ExcelFile(filepath)

        # 处理机票
        for sheet_pattern in CTRIP_SHEET_MAPPING['flight']:
            if sheet_pattern in excel_file.sheet_names:
                df = read_ctrip_sheet(filepath, sheet_pattern)
                if df is not None and len(df) > 0:
                    print(f'  处理机票数据 ({sheet_pattern}): {len(df)} 条')
                    for _, row in df.iterrows():
                        record = extract_ctrip_flight_record(row, roster_index)
                        if record:
                            all_records.append(record)

        # 处理酒店
        for sheet_pattern in CTRIP_SHEET_MAPPING['hotel']:
            if sheet_pattern in excel_file.sheet_names:
                df = read_ctrip_sheet(filepath, sheet_pattern)
                if df is not None and len(df) > 0:
                    print(f'  处理酒店数据 ({sheet_pattern}): {len(df)} 条')
                    for _, row in df.iterrows():
                        record = extract_ctrip_hotel_record(row, roster_index)
                        if record:
                            all_records.append(record)

    except Exception as e:
        print(f'  错误: 无法读取Excel文件: {e}')
        import traceback
        traceback.print_exc()
        return []

    print(f'  总共提取 {len(all_records)} 条记录')
    return all_records


def process_ctrip(
    filepath: Path,
    output_dir: Path,
    roster_index_path: Path,
    file_info: Optional[TravelFileInfo] = None
) -> Optional[str]:
    """
    处理携程商旅文件并保存结果

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
    records = process_ctrip_file(filepath, roster_index)

    if not records:
        return month

    # 保存结果
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f'ctrip_{month}.json'

    output_data = {
        'source': '携程商旅',
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

    parser = argparse.ArgumentParser(description='处理携程商旅文件')
    parser.add_argument('input', help='输入的Excel文件路径')
    parser.add_argument('-o', '--output', default='data/processed', help='输出目录')
    parser.add_argument('-r', '--roster', default='data/processed/roster_index.json', help='花名册索引文件')

    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)
    roster_index_path = Path(args.roster)

    process_ctrip(input_path, output_dir / 'by-month', roster_index_path)
