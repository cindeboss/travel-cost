#!/usr/bin/env python3
"""
文件扫描工具模块

自动扫描、分类和匹配差旅数据文件。
"""

import os
from pathlib import Path
from typing import Dict, List, Optional, NamedTuple
from dataclasses import dataclass, field
from datetime import datetime

from .date_matcher import (
    parse_date_range_from_filename,
    extract_roster_month,
    DateRange,
    find_matching_roster_file
)


@dataclass
class TravelFileInfo:
    """商旅文件信息"""
    filename: str
    filepath: Path
    source: str  # 'alibaba', 'ctrip', 'zaitu'
    date_range: Optional[DateRange] = None
    target_month: Optional[str] = None  # 主要归属月份
    matching_roster: Optional[str] = None  # 匹配的花名册文件名


@dataclass
class RosterFileInfo:
    """花名册文件信息"""
    filename: str
    filepath: Path
    month: str  # YYYY-MM 格式


@dataclass
class ScanResult:
    """扫描结果"""
    rosters: Dict[str, RosterFileInfo] = field(default_factory=dict)  # month -> info
    alibaba: List[TravelFileInfo] = field(default_factory=list)
    ctrip: List[TravelFileInfo] = field(default_factory=list)
    zaitu: List[TravelFileInfo] = field(default_factory=list)

    @property
    def all_travel_files(self) -> List[TravelFileInfo]:
        """所有商旅文件"""
        return self.alibaba + self.ctrip + self.zaitu

    @property
    def roster_months(self) -> List[str]:
        """所有可用的花名册月份"""
        return sorted(self.rosters.keys())

    def get_roster_map(self) -> Dict[str, str]:
        """获取月份到花名册文件名的映射"""
        return {month: info.filename for month, info in self.rosters.items()}


def scan_excel_files(directory: Path) -> List[Path]:
    """
    扫描目录中的所有Excel文件

    Args:
        directory: 目录路径

    Returns:
        Excel文件路径列表
    """
    excel_extensions = {'.xlsx', '.xls'}
    files = []

    if not directory.exists():
        return files

    for filepath in directory.iterdir():
        if filepath.is_file() and filepath.suffix.lower() in excel_extensions:
            files.append(filepath)

    return sorted(files, key=lambda x: x.name)


def classify_file(filename: str) -> Optional[str]:
    """
    根据文件名分类文件类型

    Args:
        filename: 文件名

    Returns:
        文件类型: 'roster', 'alibaba', 'ctrip', 'zaitu' 或 None
    """
    name = filename.lower()

    # 花名册：匹配 "花名册" 关键字
    if '花名册' in name:
        return 'roster'

    # 阿里商旅：以 "阿里" 开头
    if name.startswith('阿里'):
        return 'alibaba'

    # 携程商旅：以 "携程" 开头
    if name.startswith('携程'):
        return 'ctrip'

    # 在途商旅：以 "在途" 开头
    if name.startswith('在途'):
        return 'zaitu'

    return None


def scan_and_classify_files(raw_dir: Path) -> ScanResult:
    """
    扫描并分类原始数据目录中的所有文件

    Args:
        raw_dir: 原始数据目录路径

    Returns:
        ScanResult 对象
    """
    result = ScanResult()
    files = scan_excel_files(raw_dir)

    for filepath in files:
        filename = filepath.name
        file_type = classify_file(filename)

        if file_type == 'roster':
            # 处理花名册
            month = extract_roster_month(filename)
            if month:
                result.rosters[month] = RosterFileInfo(
                    filename=filename,
                    filepath=filepath,
                    month=month
                )
            else:
                print(f'警告: 无法从花名册文件名提取月份: {filename}')

        elif file_type in ('alibaba', 'ctrip', 'zaitu'):
            # 处理商旅数据
            date_range = parse_date_range_from_filename(filename)
            target_month = date_range.main_month if date_range else None

            travel_info = TravelFileInfo(
                filename=filename,
                filepath=filepath,
                source=file_type,
                date_range=date_range,
                target_month=target_month
            )

            getattr(result, file_type).append(travel_info)

    # 匹配花名册
    roster_map = result.get_roster_map()
    for travel_info in result.all_travel_files:
        if travel_info.target_month:
            travel_info.matching_roster = find_matching_roster_file(
                travel_info.target_month,
                roster_map
            )

    return result


def print_scan_summary(result: ScanResult):
    """
    打印扫描结果摘要

    Args:
        result: ScanResult 对象
    """
    print('=' * 70)
    print('文件扫描结果')
    print('=' * 70)

    print(f'\n花名册文件 ({len(result.rosters)} 个):')
    if result.rosters:
        for month in sorted(result.rosters.keys()):
            info = result.rosters[month]
            print(f'  {month}: {info.filename}')
    else:
        print('  (无)')

    print(f'\n阿里商旅文件 ({len(result.alibaba)} 个):')
    for info in result.alibaba:
        roster_info = f' -> {info.matching_roster}' if info.matching_roster else ' (无匹配花名册)'
        month_info = f' [{info.target_month}]' if info.target_month else ' (无月份)'
        print(f'  {info.filename}{month_info}{roster_info}')

    print(f'\n携程商旅文件 ({len(result.ctrip)} 个):')
    for info in result.ctrip:
        roster_info = f' -> {info.matching_roster}' if info.matching_roster else ' (无匹配花名册)'
        month_info = f' [{info.target_month}]' if info.target_month else ' (无月份)'
        print(f'  {info.filename}{month_info}{roster_info}')

    print(f'\n在途商旅文件 ({len(result.zaitu)} 个):')
    for info in result.zaitu:
        roster_info = f' -> {info.matching_roster}' if info.matching_roster else ' (无匹配花名册)'
        month_info = f' [{info.target_month}]' if info.target_month else ' (无月份)'
        print(f'  {info.filename}{month_info}{roster_info}')

    print('\n' + '=' * 70)


def get_files_to_process(
    raw_dir: Path,
    processed_dir: Path,
    force: bool = False
) -> List[Path]:
    """
    获取需要处理的文件列表

    根据文件的修改时间和已处理记录，判断哪些文件需要重新处理

    Args:
        raw_dir: 原始文件目录
        processed_dir: 处理后文件目录
        force: 是否强制重新处理所有文件

    Returns:
        需要处理的文件路径列表
    """
    if force:
        return scan_excel_files(raw_dir)

    # 检查原始文件的修改时间
    files_to_process = []
    raw_files = scan_excel_files(raw_dir)

    # 获取已处理文件的元数据（如果存在）
    processed_metadata = processed_dir / '.processed.json'
    processed_times = {}

    if processed_metadata.exists():
        import json
        try:
            with open(processed_metadata, 'r', encoding='utf-8') as f:
                processed_times = json.load(f)
        except Exception:
            pass

    for raw_file in raw_files:
        raw_mtime = raw_file.stat().st_mtime
        processed_mtime = processed_times.get(raw_file.name, 0)

        if raw_mtime > processed_mtime:
            files_to_process.append(raw_file)

    return files_to_process


def update_processed_metadata(raw_dir: Path, processed_dir: Path):
    """
    更新已处理文件的元数据

    Args:
        raw_dir: 原始文件目录
        processed_dir: 处理后文件目录
    """
    import json
    from datetime import datetime

    raw_files = scan_excel_files(raw_dir)
    processed_times = {}

    for raw_file in raw_files:
        processed_times[raw_file.name] = raw_file.stat().st_mtime

    metadata_path = processed_dir / '.processed.json'
    with open(metadata_path, 'w', encoding='utf-8') as f:
        json.dump(processed_times, f, indent=2)


if __name__ == '__main__':
    # 测试代码
    import sys

    # 默认使用项目目录
    if len(sys.argv) > 1:
        test_dir = Path(sys.argv[1])
    else:
        test_dir = Path(__file__).parent.parent.parent / 'data' / 'raw'

    print(f'扫描目录: {test_dir}')
    print()

    result = scan_and_classify_files(test_dir)
    print_scan_summary(result)
