#!/usr/bin/env python3
"""
一键处理所有数据脚本

自动检测、分类并处理所有差旅数据文件。
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime

# 添加父目录到路径以导入utils和处理器
sys.path.insert(0, str(Path(__file__).parent))

from utils import scan_and_classify_files, print_scan_summary, update_processed_metadata
from process_roster import process_roster
from process_alibaba import process_alibaba
from process_ctrip import process_ctrip
from process_zaitu import process_zaitu
from merge_data import merge_data


def process_all_files(
    raw_dir: Path,
    output_dir: Path,
    force: bool = False
) -> bool:
    """
    处理所有数据文件

    Args:
        raw_dir: 原始数据目录
        output_dir: 输出目录
        force: 是否强制重新处理所有文件

    Returns:
        是否成功
    """
    print('=' * 70)
    print('差旅数据处理工具')
    print('=' * 70)
    print(f'开始时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print()

    # 检查目录
    if not raw_dir.exists():
        print(f'错误: 原始数据目录不存在: {raw_dir}')
        print(f'请将Excel文件放入以下目录:')
        print(f'  {raw_dir}')
        return False

    by_month_dir = output_dir / 'by-month'
    roster_index_path = output_dir / 'roster_index.json'

    # 扫描并分类文件
    print('扫描原始数据文件...')
    scan_result = scan_and_classify_files(raw_dir)
    print_scan_summary(scan_result)

    # 统计
    total_files = (
        len(scan_result.rosters) +
        len(scan_result.alibaba) +
        len(scan_result.ctrip) +
        len(scan_result.zaitu)
    )

    if total_files == 0:
        print('\n警告: 没有找到任何可处理的文件')
        print(f'请确保文件命名符合以下格式:')
        print(f'  花名册: *花名册*.xlsx (如: 2025年12月花名册.xlsx)')
        print(f'  阿里商旅: 阿里*.xlsx (如: 阿里20251125-20251224.xlsx)')
        print(f'  携程商旅: 携程*.xlsx (如: 携程20251126-20251225.xlsx)')
        print(f'  在途商旅: 在途*.xls (如: 在途20251126-20251225.xls)')
        return False

    # Phase 1: 处理花名册（必须先处理，因为其他数据需要关联部门信息）
    print('\n' + '=' * 70)
    print('Phase 1: 处理花名册')
    print('=' * 70)

    processed_months = set()
    for month, roster_info in sorted(scan_result.rosters.items()):
        month = process_roster(
            roster_info.filepath,
            by_month_dir,
            roster_index_path
        )
        if month:
            processed_months.add(month)

    if not processed_months:
        print('警告: 没有成功处理任何花名册文件')
        print('其他数据将无法关联部门信息')
    else:
        print(f'\n成功处理 {len(processed_months)} 个月份的花名册')

    # Phase 2: 处理商旅数据
    print('\n' + '=' * 70)
    print('Phase 2: 处理商旅数据')
    print('=' * 70)

    # 处理阿里商旅
    print('\n--- 阿里商旅 ---')
    for file_info in scan_result.alibaba:
        process_alibaba(
            file_info.filepath,
            by_month_dir,
            roster_index_path,
            file_info
        )

    # 处理携程商旅
    print('\n--- 携程商旅 ---')
    for file_info in scan_result.ctrip:
        process_ctrip(
            file_info.filepath,
            by_month_dir,
            roster_index_path,
            file_info
        )

    # 处理在途商旅
    print('\n--- 在途商旅 ---')
    for file_info in scan_result.zaitu:
        process_zaitu(
            file_info.filepath,
            by_month_dir,
            roster_index_path,
            file_info
        )

    # Phase 3: 合并数据
    print('\n' + '=' * 70)
    print('Phase 3: 合并数据')
    print('=' * 70)

    travel_data_path = output_dir / 'travel-data.json'
    success = merge_data(by_month_dir, travel_data_path, roster_index_path)

    if success:
        # 更新处理元数据
        update_processed_metadata(raw_dir, output_dir)
        print('\n处理元数据已更新')

    print('\n' + '=' * 70)
    print('处理完成')
    print('=' * 70)
    print(f'结束时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')

    if success:
        print(f'\n数据文件已生成:')
        print(f'  花名册索引: {roster_index_path}')
        print(f'  差旅数据: {travel_data_path}')
        print(f'\n下一步: 运行 generate_html.py 生成HTML文件')

    return success


def main():
    parser = argparse.ArgumentParser(
        description='差旅数据处理工具 - 一键处理所有Excel数据',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例用法:
  python process_all.py                    # 使用默认目录
  python process_all.py -f                 # 强制重新处理所有文件
  python process_all.py -i data/raw -o data/processed

输出文件:
  data/processed/roster_index.json        # 花名册索引
  data/processed/by-month/*.json           # 按月分片的数据
  data/processed/travel-data.json          # 合并后的完整数据
        '''
    )

    parser.add_argument(
        '-i', '--input',
        default='data/raw',
        help='原始数据目录 (默认: data/raw)'
    )
    parser.add_argument(
        '-o', '--output',
        default='data/processed',
        help='输出目录 (默认: data/processed)'
    )
    parser.add_argument(
        '-f', '--force',
        action='store_true',
        help='强制重新处理所有文件（忽略修改时间检查）'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='显示详细信息'
    )

    args = parser.parse_args()

    raw_dir = Path(args.input)
    output_dir = Path(args.output)

    success = process_all_files(raw_dir, output_dir, args.force)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
