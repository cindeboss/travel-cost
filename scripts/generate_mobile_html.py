#!/usr/bin/env python3
"""
移动端HTML生成脚本

生成移动端友好的HTML文件，数据外置到外部JSON文件。
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime
import shutil


def generate_mobile_html(
    data_path: Path,
    template_path: Path,
    output_dir: Path
) -> bool:
    """
    生成移动端HTML文件（数据外置）

    Args:
        data_path: 数据文件路径
        template_path: HTML模板路径
        output_dir: 输出目录

    Returns:
        是否成功
    """
    print('=' * 70)
    print('生成移动端HTML文件（数据外置）')
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

    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)

    # 复制数据文件到输出目录
    data_output_path = output_dir / 'travel-data.json'
    print(f'复制数据文件: {data_output_path}')
    shutil.copy(data_path, data_output_path)

    data_size = data_output_path.stat().st_size
    print(f'  数据文件大小: {data_size:,} 字节 ({data_size / 1024 / 1024:.2f} MB)')

    # 嵌入app.js（修改为异步加载数据）
    modified_app_js = app_js_content.replace(
        'this.data = TRAVEL_DATA;',
        '''
        // 移动端：异步加载数据
        fetch('travel-data.json')
            .then(response => response.json())
            .then(data => {
                this.data = data;
                this.filteredData = [...data.records];
                this.initUI();
                this.bindEvents();
                this.applyFilters();
                // 隐藏加载界面
                const loadingScreen = document.getElementById('loadingScreen');
                if (loadingScreen) loadingScreen.classList.add('hidden');
            })
            .catch(error => {
                console.error('加载数据失败:', error);
                const loadingScreen = document.getElementById('loadingScreen');
                const loadingText = document.getElementById('loadingText');
                if (loadingScreen) {
                    loadingScreen.classList.add('error');
                }
                if (loadingText) {
                    loadingText.textContent = '加载数据失败: ' + error.message;
                }
            });
        '''
    )

    # 构建脚本（不包含数据，使用CDN）
    embedded_scripts = f'''    <script src="https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
    <script>
{modified_app_js}
    </script>'''

    # 替换模板中的脚本引用
    html_content = template.replace(
        '    <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>\n    <script src="https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js"></script>\n    <script src="app.js"></script>',
        embedded_scripts
    )

    # 添加生成时间戳
    html_content = html_content.replace(
        'GENERATION_TIMESTAMP',
        datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )

    # 添加移动端提示
    html_content = html_content.replace(
        '</body>',
        '''    <div class="mobile-footer">
            <p>💡 移动端提示：确保data.json和此HTML在同一目录下</p>
        </div>
</body>'''
    )

    # 添加移动端样式
    html_content = html_content.replace(
        '</style>',
        '''
        .mobile-footer {
            text-align: center;
            padding: 1rem;
            background: #f1f5f9;
            color: #64748b;
            font-size: 0.875rem;
            border-top: 1px solid #e2e8f0;
            margin-top: 2rem;
        }
</style>'''
    )

    # 保存HTML文件
    html_output_path = output_dir / 'travel-analysis-mobile.html'
    with open(html_output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    html_size = html_output_path.stat().st_size
    print(f'保存HTML文件: {html_output_path}')
    print(f'  文件大小: {html_size:,} 字节 ({html_size / 1024:.1f} KB)')

    print('\n移动端HTML生成完成!')
    print(f'文件位置: {html_output_path}')
    print(f'数据文件: {data_output_path}')
    print('\n使用说明:')
    print('1. 将以下两个文件放在同一目录下:')
    print(f'   - {html_output_path.name}')
    print(f'   - {data_output_path.name}')
    print('2. 用企业微信或其他移动浏览器打开HTML文件')
    print('3. 需要网络连接加载图表库')

    return True


def main():
    parser = argparse.ArgumentParser(
        description='生成移动端差旅数据分析HTML文件',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例用法:
  python generate_mobile_html.py                          # 使用默认路径
  python generate_mobile_html.py -o output/mobile

注意事项:
  - 移动端版本数据外置，需要和HTML文件在同一目录
  - 需要网络连接加载图表库（CDN）
  - HTML文件较小，适合移动浏览器
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
        default='output/mobile',
        help='输出目录 (默认: output/mobile)'
    )

    args = parser.parse_args()

    data_path = Path(args.data)
    template_path = Path(args.template)
    output_dir = Path(args.output)

    success = generate_mobile_html(data_path, template_path, output_dir)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
