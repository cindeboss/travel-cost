#!/usr/bin/env python3
"""
日期匹配工具模块

用于从文件名中提取日期范围，并计算主要归属月份。
"""

import re
from datetime import datetime, date, timedelta
from typing import Tuple, Optional, Dict
from dataclasses import dataclass


@dataclass
class DateRange:
    """日期范围数据类"""
    start_date: date
    end_date: date
    days_in_month: Dict[str, int]  # 月份 -> 天数

    @property
    def total_days(self) -> int:
        """总天数"""
        return (self.end_date - self.start_date).days + 1

    @property
    def main_month(self) -> str:
        """主要归属月份 (YYYY-MM格式)"""
        if not self.days_in_month:
            return None

        # 返回天数最多的月份
        main_month = max(self.days_in_month.items(), key=lambda x: x[1])
        return main_month[0]

    @property
    def main_month_ratio(self) -> float:
        """主要月份的天数占比"""
        if not self.days_in_month:
            return 0
        return self.days_in_month.get(self.main_month, 0) / self.total_days


def extract_date_from_string(text: str, year_hint: Optional[int] = None) -> Optional[date]:
    """
    从字符串中提取日期

    支持格式：
    - 20251125 或 2025-11-25
    - 1125 (需要year_hint)

    Args:
        text: 包含日期的字符串
        year_hint: 年份提示（当文本中只有月日时使用）

    Returns:
        date对象或None
    """
    # 尝试匹配 YYYYMMDD 格式
    match = re.search(r'(\d{4})(\d{2})(\d{2})', text)
    if match:
        year, month, day = map(int, match.groups())
        try:
            return date(year, month, day)
        except ValueError:
            pass

    # 尝试匹配 YYYY-MM-DD 格式
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', text)
    if match:
        year, month, day = map(int, match.groups())
        try:
            return date(year, month, day)
        except ValueError:
            pass

    # 尝试匹配 MMDD 格式（需要year_hint）
    if year_hint:
        match = re.search(r'(\d{2})(\d{2})', text)
        if match:
            month, day = map(int, match.groups())
            try:
                return date(year_hint, month, day)
            except ValueError:
                pass

    return None


def parse_date_range_from_filename(filename: str) -> Optional[DateRange]:
    """
    从文件名中解析日期范围

    支持格式：
    - 阿里20251125-20251224.xlsx
    - 携程20251126-20251225.xlsx
    - 在途202510-202511.xls

    Args:
        filename: 文件名

    Returns:
        DateRange对象或None
    """
    # 尝试匹配 YYYYMMDD-YYYYMMDD 格式
    match = re.search(r'(\d{8})-(\d{8})', filename)
    if match:
        start_str, end_str = match.groups()
        start_date = extract_date_from_string(start_str)
        end_date = extract_date_from_string(end_str)

        if start_date and end_date:
            return _calculate_days_by_month(start_date, end_date)

    # 尝试匹配 YYYYMM-YYYYMM 格式（月份范围）
    match = re.search(r'(\d{6})-(\d{6})', filename)
    if match:
        start_str, end_str = match.groups()
        # 假设月初到月末
        year1, month1 = int(start_str[:4]), int(start_str[4:6])
        year2, month2 = int(end_str[:4]), int(end_str[4:6])

        try:
            start_date = date(year1, month1, 1)
            # 计算月末
            if month2 == 12:
                end_date = date(year2, 12, 31)
            else:
                end_date = date(year2, month2 + 1, 1) - timedelta(days=1)
            return _calculate_days_by_month(start_date, end_date)
        except ValueError:
            pass

    # 尝试匹配单个 YYYYMM 格式
    match = re.search(r'(\d{6})', filename)
    if match:
        date_str = match.group(1)
        year, month = int(date_str[:4]), int(date_str[4:6])
        try:
            # 假设整月
            start_date = date(year, month, 1)
            if month == 12:
                end_date = date(year, 12, 31)
            else:
                end_date = date(year, month + 1, 1) - timedelta(days=1)
            return _calculate_days_by_month(start_date, end_date)
        except ValueError:
            pass

    return None


def _calculate_days_by_month(start_date: date, end_date: date) -> DateRange:
    """
    计算日期范围内每个月的天数

    Args:
        start_date: 开始日期
        end_date: 结束日期

    Returns:
        DateRange对象
    """
    days_by_month = {}
    current = start_date

    while current <= end_date:
        month_key = current.strftime('%Y-%m')
        days_by_month[month_key] = days_by_month.get(month_key, 0) + 1
        current += timedelta(days=1)

    return DateRange(
        start_date=start_date,
        end_date=end_date,
        days_in_month=days_by_month
    )


def extract_roster_month(filename: str) -> Optional[str]:
    """
    从花名册文件名中提取月份

    支持格式：
    - 2025年12月花名册.xlsx
    - 2025年11月花名册.xlsx
    - 花名册202512.xlsx

    Args:
        filename: 文件名

    Returns:
        月份字符串 (YYYY-MM格式) 或None
    """
    # 尝试匹配 YYYY年MM月 格式
    match = re.search(r'(\d{4})年(\d{1,2})月', filename)
    if match:
        year, month = int(match.group(1)), int(match.group(2))
        return f'{year}-{month:02d}'

    # 尝试匹配 YYYYMM 格式
    match = re.search(r'(\d{6})', filename)
    if match:
        date_str = match.group(1)
        year, month = int(date_str[:4]), int(date_str[4:6])
        return f'{year}-{month:02d}'

    return None


def find_matching_roster_file(
    travel_month: str,
    available_rosters: Dict[str, str]
) -> Optional[str]:
    """
    根据商旅数据的归属月份，查找匹配的花名册文件

    Args:
        travel_month: 商旅数据归属月份 (YYYY-MM)
        available_rosters: 可用的花名册字典 {月份: 文件名}

    Returns:
        匹配的花名册文件名，如果没有找到则返回None
    """
    # 精确匹配
    if travel_month in available_rosters:
        return available_rosters[travel_month]

    # 尝试找最近的月份（优先找当月或前一个月）
    travel_year, travel_month_int = map(int, travel_month.split('-'))

    # 优先找当月
    if travel_month in available_rosters:
        return available_rosters[travel_month]

    # 找前一个月
    if travel_month_int > 1:
        prev_month = f'{travel_year}-{travel_month_int - 1:02d}'
        if prev_month in available_rosters:
            return available_rosters[prev_month]

    # 找后一个月
    if travel_month_int < 12:
        next_month = f'{travel_year}-{travel_month_int + 1:02d}'
        if next_month in available_rosters:
            return available_rosters[next_month]

    # 找最近的可用的月份
    if available_rosters:
        # 按月份排序，找最近的
        sorted_months = sorted(available_rosters.keys())
        for m in sorted_months:
            if m <= travel_month:
                return available_rosters[m]

        # 如果都比目标月份早，返回最新的
        return available_rosters[sorted_months[-1]]

    return None


if __name__ == '__main__':
    # 测试代码
    test_filenames = [
        '阿里20251125-20251224.xlsx',
        '携程20251126-20251225.xlsx',
        '在途202511-202512.xls',
        '2025年12月花名册.xlsx',
        '2025年11月花名册.xlsx',
    ]

    print('日期范围解析测试:')
    print('-' * 60)

    for filename in test_filenames:
        print(f'\n文件: {filename}')

        if '花名册' in filename:
            month = extract_roster_month(filename)
            print(f'  花名册月份: {month}')
        else:
            date_range = parse_date_range_from_filename(filename)
            if date_range:
                print(f'  日期范围: {date_range.start_date} 至 {date_range.end_date}')
                print(f'  总天数: {date_range.total_days}')
                print(f'  每月天数: {date_range.days_in_month}')
                print(f'  主要归属: {date_range.main_month} ({date_range.main_month_ratio:.1%})')
            else:
                print('  无法解析日期范围')
