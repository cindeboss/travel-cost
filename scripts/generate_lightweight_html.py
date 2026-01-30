#!/usr/bin/env python3
"""
轻量级HTML生成脚本

通过数据抽样生成较小的HTML文件，适合移动端使用。
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime
from collections import defaultdict


def sample_data(data: dict, max_records: int = 500) -> dict:
    """
    抽样数据以减小文件大小

    保留全部统计信息，但限制明细记录数量
    """
    records = data.get('records', [])

    # 按类型分组
    by_type = defaultdict(list)
    for r in records:
        by_type[r['type']].append(r)

    # 每种类型抽样最近的记录
    sampled_records = []
    for record_type, type_records in by_type.items():
        # 按日期排序（如果有日期字段）
        def get_date(r):
            if r.get('departTime'):
                return r['departTime']
            elif r.get('checkInTime'):
                return r['checkInTime']
            elif r.get('pickupTime'):
                return r['pickupTime']
            return ''

        sorted_records = sorted(type_records, key=get_date, reverse=True)
        sampled_records.extend(sorted_records[:max_records])

    # 重新构建summary
    summary = data.get('summary', {})

    return {
        'lastUpdate': data.get('lastUpdate'),
        'months': data.get('months', []),
        'summary': summary,
        'records': sampled_records,
        'isSample': True,
        'totalRecords': len(records),
        'sampledRecords': len(sampled_records)
    }


def generate_lightweight_html(
    data_path: Path,
    template_path: Path,
    output_path: Path,
    max_records: int = 200
) -> bool:
    """
    生成轻量级HTML文件（数据抽样）
    """
    print('=' * 70)
    print('生成轻量级HTML文件（数据抽样）')
    print('=' * 70)

    # 读取数据
    if not data_path.exists():
        print(f'错误: 数据文件不存在: {data_path}')
        print('请先运行 process_all.py 处理数据')
        return False

    print(f'读取数据文件: {data_path}')
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    total_records = len(data.get('records', []))
    print(f'  原始记录数: {total_records}')

    # 抽样数据
    sampled_data = sample_data(data, max_records=200)
    sampled_count = len(sampled_data['records'])
    print(f'  抽样记录数: {sampled_count} (每种类型最多200条)')

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

    # 读取第三方库（使用缓存）
    import urllib.request

    echarts_path = Path('/tmp/echarts.min.js')
    dayjs_path = Path('/tmp/dayjs.min.js')

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

    with open(echarts_path, 'r', encoding='utf-8') as f:
        echarts_content = f.read()

    with open(dayjs_path, 'r', encoding='utf-8') as f:
        dayjs_content = f.read()

    print(f'  echarts: {len(echarts_content):,} 字节 ({len(echarts_content) / 1024:.1f} KB)')
    print(f'  dayjs: {len(dayjs_content):,} 字节 ({len(dayjs_content) / 1024:.1f} KB)')

    # 嵌入抽样数据
    data_json = json.dumps(sampled_data, ensure_ascii=False, indent=2)

    # 添加抽样提示到app.js
    sampled_notice = '''
    // 抽样数据提示
    const isSample = true;
    const totalRecords = ''' + str(total_records) + ''';
    const sampledRecords = ''' + str(sampled_count) + ''';
    '''

    modified_app_js = sampled_notice + app_js_content

    # 构建内嵌脚本
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
{modified_app_js}
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

    # 添加抽样提示到页面
    if 'recordSummary' in html_content or 'recordSource' in html_content:
        html_content = html_content.replace(
            'id="recordSummary"',
            'id="recordSummary_sample"'
        )
        html_content = html_content.replace(
            'id="recordSource"',
            'id="recordSource_sample"'
        )
        # 在页面加载后显示抽样提示
        html_content = html_content.replace(
            '</body>',
            '''    <script>
        window.addEventListener('load', function() {
            const summary = document.getElementById('recordSummary_sample');
            const source = document.getElementById('recordSource_sample');
            if (summary) {
                summary.innerHTML = `共 ${sampledRecords.toLocaleString()} 条记录 (抽样显示，总计 ${totalRecords.toLocaleString()} 条)`;
            }
            if (source) {
                source.innerHTML = `抽样数据 | 共 ${totalRecords.toLocaleString()} 条记录`;
            }
            // 添加顶部提示条
            const header = document.querySelector('.header-left');
            if (header) {
                const notice = document.createElement('div');
                notice.style.cssText = 'font-size: 0.75rem; color: #f59e0b; margin-top: 0.25rem;';
                notice.textContent = '📱 移动版：显示每种类型最近200条记录';
                header.appendChild(notice);
            }
        });
    </script>
</body>'''
        )

    # 保存HTML文件
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    file_size = output_path.stat().st_size
    print(f'保存HTML文件: {output_path}')
    print(f'  文件大小: {file_size:,} 字节 ({file_size / 1024:.1f} KB)')

    print('\n轻量级HTML生成完成!')
    print(f'请在浏览器中打开: {output_path}')
    print('\n说明:')
    print(f'- 原始数据: {total_records} 条记录')
    print(f'- 抽样显示: {sampled_count} 条记录 (每种类型最多200条)')
    print('- 概览统计基于全部数据')
    print('- 明细表格显示抽样数据')
    print('- 文件较小，适合移动端和分享')

    return True


def main():
    parser = argparse.ArgumentParser(
        description='生成轻量级差旅数据分析HTML文件',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例用法:
  python generate_lightweight_html.py                        # 使用默认路径
  python generate_lightweight_html.py -o output-light.html  # 指定输出文件

注意事项:
  - 通过数据抽样减小文件大小，适合移动端
  - 概览统计基于全部数据
  - 明细表格显示抽样数据
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
        default='output/travel-analysis-light.html',
        help='输出HTML文件路径 (默认: output/travel-analysis-light.html)'
    )
    parser.add_argument(
        '--max-records',
        type=int,
        default=200,
        help='每种类型最多保留的记录数 (默认: 200)'
    )

    args = parser.parse_args()

    data_path = Path(args.data)
    template_path = Path(args.template)
    output_path = Path(args.output)

    success = generate_lightweight_html(data_path, template_path, output_path, args.max_records)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
