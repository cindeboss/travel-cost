#!/usr/bin/env python3
"""
HTML生成脚本

读取处理后的数据，生成包含数据的独立HTML文件。
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime


def generate_html(
    data_path: Path,
    template_path: Path,
    output_path: Path
) -> bool:
    """
    生成包含数据的HTML文件（单文件，包含所有代码、库和数据）

    Args:
        data_path: 数据文件路径
        template_path: HTML模板路径
        output_path: 输出HTML文件路径

    Returns:
        是否成功
    """
    print('=' * 70)
    print('生成HTML文件（完全单文件，包含所有库）')
    print('=' * 70)

    # 读取数据
    if not data_path.exists():
        print(f'错误: 数据文件不存在: {data_path}')
        print('请先运行 process_all.py 处理数据')
        return False

    print(f'读取数据文件: {data_path}')
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    print(f'  记录数: {len(data.get("records", []))}')
    print(f'  总金额: ¥{data.get("summary", {}).get("totalAmount", 0):,.2f}')

    # 读取模板
    if not template_path.exists():
        print(f'错误: HTML模板不存在: {template_path}')
        return False

    print(f'读取HTML模板: {template_path}')
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    # 读取app.js
    app_js_path = template_path.parent / 'app.js'
    if not app_js_path.exists():
        print(f'错误: app.js不存在: {app_js_path}')
        return False

    print(f'读取app.js: {app_js_path}')
    with open(app_js_path, 'r', encoding='utf-8') as f:
        app_js_content = f.read()

    # 读取第三方库（如果本地缓存不存在则下载）
    import urllib.request
    import os

    echarts_path = Path('/tmp/echarts.min.js')
    dayjs_path = Path('/tmp/dayjs.min.js')

    # 下载echarts（如果需要）
    if not echarts_path.exists():
        print('下载echarts库...')
        try:
            urllib.request.urlretrieve(
                'https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js',
                echarts_path
            )
            print(f'  echarts: {echarts_path.stat().st_size / 1024:.1f} KB')
        except Exception as e:
            print(f'  警告: 无法下载echarts: {e}')
            return False

    # 下载dayjs（如果需要）
    if not dayjs_path.exists():
        print('下载dayjs库...')
        try:
            urllib.request.urlretrieve(
                'https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js',
                dayjs_path
            )
            print(f'  dayjs: {dayjs_path.stat().st_size / 1024:.1f} KB')
        except Exception as e:
            print(f'  警告: 无法下载dayjs: {e}')
            return False

    # 读取库文件
    with open(echarts_path, 'r', encoding='utf-8') as f:
        echarts_content = f.read()

    with open(dayjs_path, 'r', encoding='utf-8') as f:
        dayjs_content = f.read()

    print(f'  echarts: {len(echarts_content):,} 字节 ({len(echarts_content) / 1024:.1f} KB)')
    print(f'  dayjs: {len(dayjs_content):,} 字节 ({len(dayjs_content) / 1024:.1f} KB)')

    # 嵌入数据
    data_json = json.dumps(data, ensure_ascii=False, indent=2)

    # 构建内嵌脚本（包含库）
    embedded_scripts = f'''    <script>
{dayjs_content}
    </script>
    <script>
{echarts_content}
    </script>
    <script>
        const TRAVEL_DATA = {data_json};
    </script>
    <script>
{app_js_content}
    </script>'''

    # 替换CDN引用为内嵌
    html_content = template.replace(
        '    <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>\n    <script src="https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js"></script>\n    <script src="app.js"></script>',
        embedded_scripts
    )

    # 添加生成时间戳
    html_content = html_content.replace(
        'GENERATION_TIMESTAMP',
        datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )

    # 保存HTML文件
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f'保存HTML文件: {output_path}')
    print(f'  文件大小: {len(html_content):,} 字节 ({len(html_content) / 1024 / 1024:.2f} MB)')

    print('\n生成完成!')
    print(f'请在浏览器中打开: {output_path}')
    print('注意: 文件包含所有库和数据，可以离线使用')

    return True


def main():
    parser = argparse.ArgumentParser(
        description='生成差旅数据分析HTML文件',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例用法:
  python generate_html.py                              # 使用默认路径
  python generate_html.py -d data.json -t template.html -o output.html

注意事项:
  - 确保已运行 process_all.py 生成数据文件
  - HTML文件包含所有数据，可直接在浏览器中打开
        '''
    )

    parser.add_argument(
        '-d', '--data',
        default='data/processed/travel-data.json',
        help='数据文件路径 (默认: data/processed/travel-data.json)'
    )
    parser.add_argument(
        '-t', '--template',
        default='templates/travel-analysis.html',
        help='HTML模板路径 (默认: templates/travel-analysis.html)'
    )
    parser.add_argument(
        '-o', '--output',
        default='output/travel-analysis.html',
        help='输出HTML文件路径 (默认: output/travel-analysis.html)'
    )

    args = parser.parse_args()

    data_path = Path(args.data)
    template_path = Path(args.template)
    output_path = Path(args.output)

    success = generate_html(data_path, template_path, output_path)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
