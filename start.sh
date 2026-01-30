#!/bin/bash
# 差旅数据分析工具 - 启动脚本
# 功能：检测数据更新、处理数据、生成HTML、打开浏览器

set -e  # 遇到错误时退出

# 项目目录
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}  差旅数据分析工具${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}错误: 未找到 Python3${NC}"
    echo "请先安装 Python3"
    exit 1
fi

# 检查依赖
echo -e "${BLUE}检查依赖...${NC}"
if ! python3 -c "import pandas, openpyxl" &> /dev/null; then
    echo -e "${YELLOW}安装依赖中...${NC}"
    pip3 install --quiet pandas openpyxl xlrd || {
        echo -e "${RED}依赖安装失败${NC}"
        exit 1
    }
fi
echo -e "${GREEN}✓ 依赖检查完成${NC}"
echo ""

# 检查数据目录
DATA_DIR="$PROJECT_DIR/data/raw"
if [ ! -d "$DATA_DIR" ]; then
    echo -e "${YELLOW}创建数据目录: $DATA_DIR${NC}"
    mkdir -p "$DATA_DIR"
fi

# 检查是否有Excel文件
EXCEL_COUNT=$(find "$DATA_DIR" -maxdepth 1 -type f \( -name "*.xlsx" -o -name "*.xls" \) 2>/dev/null | wc -l)

if [ "$EXCEL_COUNT" -eq 0 ]; then
    echo -e "${YELLOW}数据目录为空: $DATA_DIR${NC}"
    echo "请将以下格式的Excel文件放入该目录："
    echo "  - 花名册: *花名册*.xlsx"
    echo "  - 阿里商旅: 阿里*.xlsx"
    echo "  - 携程商旅: 携程*.xlsx"
    echo "  - 在途商旅: 在途*.xls"
    echo ""
    read -p "按回车键继续，或按 Ctrl+C 退出..."
fi

# 运行数据处理
echo -e "${BLUE}处理数据...${NC}"
python3 scripts/process_all.py

if [ $? -ne 0 ]; then
    echo -e "${RED}数据处理失败${NC}"
    read -p "按回车键退出..."
    exit 1
fi

echo ""

# 生成HTML
echo -e "${BLUE}生成HTML...${NC}"
python3 scripts/generate_html.py

if [ $? -ne 0 ]; then
    echo -e "${RED}HTML生成失败${NC}"
    read -p "按回车键退出..."
    exit 1
fi

HTML_FILE="$PROJECT_DIR/output/travel-analysis.html"
echo ""

# 检查HTML文件
if [ ! -f "$HTML_FILE" ]; then
    echo -e "${RED}HTML文件未生成: $HTML_FILE${NC}"
    read -p "按回车键退出..."
    exit 1
fi

echo -e "${GREEN}✓ 处理完成！${NC}"
echo ""

# 在浏览器中打开
case "$(uname -s)" in
    Darwin*)    # macOS
        echo -e "${BLUE}在浏览器中打开...${NC}"
        open "$HTML_FILE"
        ;;
    Linux*)     # Linux
        if command -v xdg-open &> /dev/null; then
            echo -e "${BLUE}在浏览器中打开...${NC}"
            xdg-open "$HTML_FILE" &> /dev/null &
        elif command -v gnome-open &> /dev/null; then
            echo -e "${BLUE}在浏览器中打开...${NC}"
            gnome-open "$HTML_FILE" &> /dev/null &
        else
            echo -e "${YELLOW}请手动在浏览器中打开: $HTML_FILE${NC}"
        fi
        ;;
    MINGW*|MSYS*|CYGWIN*)  # Windows (Git Bash, MSYS, Cygwin)
        echo -e "${BLUE}在浏览器中打开...${NC}"
        start "$HTML_FILE"
        ;;
    *)
        echo -e "${YELLOW}请手动在浏览器中打开: $HTML_FILE${NC}"
        ;;
esac

echo ""
echo -e "${GREEN}完成！${NC}"
