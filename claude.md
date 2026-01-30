# 差旅数据分析系统

一个纯前端的差旅数据处理、分析和查询工具，用于帮助业务部门负责人掌握销售人员的差旅情况。

---

## 项目概述

### 核心功能
- **数据源**：4个Excel文件（花名册 + 阿里/携程/在途商旅数据）
- **功能**：按一级部门筛选，查看员工差旅明细，费用统计汇总，数据筛选
- **技术方案**：纯前端（数据嵌入HTML），直接在浏览器打开使用
- **展示风格**：商务风格，深蓝色主色调 (#1a56db)

### 安全特性
- **默认部门**：教培业务中心
- **切换密码**：201212
- **部门隔离**：默认只显示教培业务中心数据，切换部门需要密码验证

---

## 项目结构

```
/Users/wanghui/Travel-cost/
├── data/                           # 数据目录
│   ├── raw/                        # 原始Excel文件（用户放置）
│   │   ├── 2025年12月花名册.xlsx
│   │   ├── 阿里20251125-20251224.xlsx
│   │   ├── 携程20251126-20251225.xlsx
│   │   ├── 在途20251126-20251225.xls
│   │   └── ...（所有Excel文件放在这里）
│   └── processed/                  # 处理后的JSON（自动生成）
│       ├── by-month/               # 按月分片的数据
│       ├── roster_index.json       # 所有月份花名册索引
│       └── travel-data.json        # 合并后的完整数据
├── scripts/                        # Python脚本
│   ├── utils/                      # 工具模块
│   │   ├── __init__.py
│   │   └── file_scanner.py         # 文件自动扫描和分类
│   ├── process_roster.py          # 处理花名册
│   ├── process_alibaba.py         # 处理阿里商旅
│   ├── process_ctrip.py           # 处理携程商旅
│   ├── process_zaitu.py           # 处理在途商旅
│   ├── merge_data.py              # 合并所有数据
│   ├── process_all.py             # 一键处理所有数据
│   └── generate_html.py           # 生成HTML文件
├── templates/                      # HTML模板
│   ├── travel-analysis.html       # HTML模板
│   └── app.js                     # 前端逻辑
├── output/                         # 输出目录
│   └── travel-analysis.html       # 最终生成的HTML（含数据）
├── start.sh                       # 启动脚本
└── claude.md                      # 本文件
```

---

## 快速开始

### 1. 安装依赖
```bash
pip install pandas openpyxl xlrd
```

### 2. 准备数据文件
将Excel文件复制到 `data/raw/` 目录：
- 花名册文件（命名格式：`YYYY年MM月花名册.xlsx`）
- 阿里商旅文件（命名格式：`阿里YYYYMMDD-YYYYMMDD.xlsx`）
- 携程商旅文件（命名格式：`携程YYYYMMDD-YYYYMMDD.xlsx`）
- 在途商旅文件（命名格式：`在途YYYYMMDD-YYYYMMDD.xls`）

### 3. 处理数据
```bash
# 方式一：一键处理（推荐）
python3 scripts/process_all.py

# 方式二：使用启动脚本
./start.sh
```

### 4. 生成HTML
```bash
python3 scripts/generate_html.py
```

### 5. 打开使用
在浏览器中打开 `output/travel-analysis.html`

---

## 前端功能

### 数据筛选
- **部门筛选**：选择一级部门，默认显示教培业务中心
- **时间范围**：全部、本月、本季度、本年
- **来源筛选**：在途商旅、携程商旅、阿里商旅
- **员工搜索**：模糊匹配员工姓名
- **列筛选**：明细表格每列都有独立的筛选输入框

### 数据展示
- **概览卡片**：总金额、各类型费用统计、订单数量
- **图表展示**：趋势图、类型分布、员工排名、部门对比
- **明细表格**：分类型Tab（机票/酒店/火车/用车）
  - 点击列头排序
  - 每列独立筛选
  - 第一列冻结（横向滚动时）
  - 分页显示

### 数据导出
- 导出当前Tab的数据为CSV

---

## 数据源格式

### 花名册文件
- **工作表**："原表"
- **关键字段**：姓名、英文名、一级部门、二级部门、三级部门、岗位、在职状态

### 阿里商旅文件
- **工作表**：
  - 本期国内机票交易明细 → `flight`
  - 国际/中国港澳台机票交易明细 → `flight`
  - 国内酒店、国际酒店 → `hotel`
  - 火车票交易明细 → `train`
  - 国内用车对账单 → `car`

### 携程商旅文件
- **工作表**：
  - 预存机票 → `flight`
  - 预存会员酒店 → `hotel`
  - 预存协议酒店 → `hotel`
  - 预存增值 → `car`

### 在途商旅文件
- **工作表**：
  - 机票/机票交易/机票明细 → `flight`
  - 酒店/酒店交易/酒店明细 → `hotel`
  - 火车/火车票/火车交易 → `train`
  - 用车/用车交易/用车明细 → `car`

---

## 前端技术栈

- **框架**：纯JavaScript（无框架依赖）
- **图表**：ECharts 5
- **日期处理**：Day.js
- **样式**：原生CSS，深蓝色商务风格

### 主要类
- `DeptSecurity`：部门安全控制（密码验证）
- `TravelAnalysisApp`：主应用类（数据处理、UI渲染）

---

## 开发指南

### 修改数据处理逻辑
1. 编辑 `scripts/process_*.py` 文件
2. 运行 `python3 scripts/process_all.py` 重新处理数据
3. 运行 `python3 scripts/generate_html.py` 重新生成HTML

### 修改前端界面
1. 编辑 `templates/travel-analysis.html`（HTML/CSS）
2. 编辑 `templates/app.js`（JavaScript逻辑）
3. 运行 `python3 scripts/generate_html.py` 重新生成HTML

### 调试技巧
- 在浏览器控制台查看错误信息
- 使用 `window.app` 访问应用实例
- 使用 `console.log()` 调试代码

---

## 数据处理脚本说明

| 脚本 | 功能 |
|------|------|
| `process_all.py` | 一键处理所有Excel数据 |
| `process_roster.py` | 处理花名册，生成员工索引 |
| `process_alibaba.py` | 处理阿里商旅数据 |
| `process_ctrip.py` | 处理携程商旅数据 |
| `process_zaitu.py` | 处理在途商旅数据 |
| `merge_data.py` | 合并所有数据源 |
| `generate_html.py` | 生成最终的HTML文件 |

---

## 常见问题

### Q: 为什么有些记录金额为负数？
A: 这是退票记录，属于正常的业务数据。

### Q: 为什么航空公司有些为空？
A: 携程商旅的原始数据没有航空公司字段，系统会从航班号前缀推断，少数无法推断的会留空。

### Q: 如何查看其他部门的数据？
A: 在部门下拉框中选择其他部门，输入密码 `201212` 即可切换。

### Q: HTML文件可以分发给其他人吗？
A: 可以，HTML文件包含所有数据、代码和库，可以离线使用。

---

# 附录：数据验证规范

## 一、Excel文件结构验证

### 1.1 携程商旅文件结构
```
行0-3: 标题行（公司信息、日期范围等）
行4: 中文列名
行5: 英文列名
行6+: 实际数据
```

**关键点**：
- 使用 `header=5` 读取数据（跳过前6行）
- 必须过滤掉第5行（英文列名行），它会被当作第一条数据
- 验证方法：检查第一列是否包含"订单号"或"OrderID"

### 1.2 阿里商旅文件结构
```
行0-1: 标题行
行2: 中文列名（表头）
行3: 合计行（需跳过）
行4+: 实际数据
```

**关键点**：
- 使用 `header=2, skiprows=[3]` 读取数据
- 注意某些字段可能为空（如航空公司）

### 1.3 在途商旅文件结构
```
.xls格式，使用xlrd引擎读取
结构类似阿里商旅，但列索引不同
```

---

## 二、数据有效性验证规则

### 2.1 乘机人/员工姓名验证

**必须过滤的无效值**：
```python
# 1. 日期格式
if passenger.startswith('202') or passenger.startswith('20') or passenger.count('-') >= 2:
    return None

# 2. 汇总关键词
if passenger in ['小计', '合计', '总计', '汇总', '平均值', 'ExchangeTime', 'ETD', 'PassengerName']:
    return None

# 3. 纯数字
if passenger.isdigit():
    return None

# 4. 列名（表头被误识别）
if passenger in ['clients', '入住人', '出行人', '乘机人']:
    return None
```

### 2.2 机票记录必填字段验证

**必须同时满足的条件**：
- `passenger` (乘机人)：有效姓名
- `flightNo` (航班号)：非空且不是"nan"
- `departTime` (起飞时间)：非空且不是"nan"
- `fromCity` (出发地)：非空且不是"nan"
- `toCity` (目的地)：非空且不是"nan"

**示例代码**：
```python
if not flight_no or flight_no == 'nan':
    return None
if not from_city or from_city == 'nan':
    return None
```

### 2.3 价格字段验证

**正常情况**：
- 价格应 > 0（正常交易）
- 价格 < 0（退票，正常业务数据）
- 价格 = 0（需要核查，可能是数据问题）

**处理原则**：
- 保留负数价格（退票记录）
- 0值价格需要检查原始数据

---

## 三、字段提取规范

### 3.1 机票数据

| 字段 | 携程索引 | 阿里索引 | 在途索引 | 补充逻辑 |
|------|---------|---------|---------|---------|
| 乘机人 | 5 | 5 | 17 | 必填验证 |
| 航班号 | 12 | 24 | 8 | 必填验证 |
| 起飞时间 | 7 | 14 | 13 | 必填验证 |
| 出发地 | 11(航程拆分) | 18 | 15 | 必填验证 |
| 目的地 | 11(航程拆分) | 19 | 15(航程拆分) | 必填验证 |
| 舱位 | 13 | 26 | - | 可选 |
| 价格 | 14 | 35 | 33 | 可为负数 |
| 航空公司 | - | 23 | - | **从航班号推断** |

**航空公司推断规则**：
```python
AIRLINE_CODE_MAP = {
    'CA': '中国国航', 'MU': '中国东方航空', 'CZ': '中国南方航空',
    '3U': '四川航空', 'ZH': '深圳航空', 'HU': '海南航空',
    'FM': '上海航空', 'KN': '中国联合航空', 'JD': '金鹏航空',
    'MF': '厦门航空', 'SC': '山东航空', 'HO': '吉祥航空'
    # ... 更多映射
}
```

### 3.2 酒店数据

| 字段 | 携程索引 | 在途索引 | 备注 |
|------|---------|---------|------|
| 员工 | 4 | 13 | 必填验证 |
| 入住时间 | 7 | 10 | - |
| 离开时间 | 8 | 11 | - |
| 城市 | 9 | 7 | - |
| 酒店名称 | 10 | 8 | - |
| 房型 | 12 | 9 | - |
| 价格 | 18 | 16或19 | - |

### 3.3 用车数据

| 字段 | 阿里索引 | 在途索引 | 备注 |
|------|---------|---------|------|
| 乘车人 | 6 | 19或6 | 必填验证 |
| 上车时间 | 14+15 | 28或- | 日期+时间拼接 |
| 下车时间 | 16+17 | 29或- | 日期+时间拼接 |
| 出发地 | 18 | 28或- | 字典结构city/address |
| 目的地 | 21 | 29或- | 字典结构city/address |
| 用车类型 | 43(平台车型) | - | **阿里有此字段** |
| 服务方 | 41 | 19 | - |
| 里程 | 25 | 29 | - |
| 金额 | 32 | 38 | - |

**阿里用车重要**：必须提取索引43的"平台车型"字段

---

## 四、常见问题及解决方案

### 4.1 表头被识别为数据行

**症状**：
- 乘机人显示为"PassengerName"、"ExchangeTime"
- 员工显示为"clients"、"入住人"

**解决方案**：
1. 检查header参数是否正确
2. 添加过滤逻辑过滤列名行
3. 使用以下代码验证：
```python
df = df[~df.iloc[:, 0].astype(str).str.contains('订单号|OrderID', na=False)]
```

### 4.2 航程字段包含出发地和目的地

**症状**：
- fromCity和toCity为空
- route字段包含"深圳-北京"格式

**解决方案**：
```python
route = str(row.iloc[11])  # 航程字段
if '-' in route:
    cities = route.split('-')
    from_city = cities[0].strip()
    to_city = cities[1].strip() if len(cities) > 1 else ''
```

### 4.3 日期被当作姓名提取

**症状**：
- 乘机人显示为"2025-11-28"或"11-28"

**解决方案**：
```python
if passenger.startswith('202') or passenger.count('-') >= 2:
    return None  # 跳过此记录
```

### 4.4 员工ID被当作姓名

**症状**：
- 乘机人显示为"EMP2203181HRL87EJ"

**解决方案**：
```python
if str(passenger).startswith('EMP'):
    passenger = row.iloc[3]  # 尝试使用预订人
```

---

## 五、数据验证检查清单

处理完每个数据源后，必须进行以下检查：

### 5.1 机票数据检查
```python
# 检查空值
empty_flight_no = [r for r in flight if not r.get('flightNo') or r.get('flightNo') == '']
empty_from_city = [r for r in flight if not r.get('fromCity') or r.get('fromCity') == '']
empty_to_city = [r for r in flight if not r.get('toCity') or r.get('toCity') == '']
empty_airline = [r for r in flight if not r.get('airline') or r.get('airline') == '']

# 检查负数价格（退票）
neg_price = [r for r in flight if r.get('price', 0) < 0]

# 预期结果：
# - 航班号、出发地、目的地空值应为0
# - 航空公司空值应<10（少数无法推断的航班）
# - 负数价格为正常退票数据
```

### 5.2 酒店数据检查
```python
empty_hotel_name = [r for r in hotel if not r.get('hotelName') or r.get('hotelName') == '']
empty_room_type = [r for r in hotel if not r.get('roomType') or r.get('roomType') == '']

# 预期结果：空值应为0
```

### 5.3 用车数据检查
```python
empty_car_type = [r for r in car if not r.get('carType') or r.get('carType') == '']
empty_provider = [r for r in car if not r.get('provider') or r.get('provider') == '']
zero_distance = [r for r in car if r.get('distance', 0) == 0]

# 预期结果：
# - 阿里商旅：用车类型空值应为0
# - 服务方空值应为0
# - 里程为0的记录应<10
```

### 5.4 火车数据检查
```python
empty_train_no = [r for r in train if not r.get('trainNo') or r.get('trainNo') == '']
empty_seat = [r for r in train if not r.get('seat') or r.get('seat') == '']

# 预期结果：空值应为0
```

---

## 六、数据处理最佳实践

### 6.1 读取数据时
```python
# 携程文件
df = pd.read_excel(filepath, sheet_name=sheet_name, header=5)
df = df[~df.iloc[:, 0].astype(str).str.contains('订单号|OrderID', na=False)]

# 阿里文件
df = pd.read_excel(filepath, sheet_name=sheet_name, header=2, skiprows=[3])

# 在途文件
df = pd.read_excel(filepath, sheet_name=sheet_name, engine='xlrd')
```

### 6.2 提取记录时
```python
# 1. 先检查行长度
if len(row) <= required_index:
    return None

# 2. 提取姓名并进行验证
passenger = str(row.iloc[index]).strip()
# 添加所有验证规则...

# 3. 提取必填字段并验证
if not required_field or required_field == 'nan':
    return None

# 4. 提取可选字段
optional_field = str(row.iloc[index]) if len(row) > index and pd.notna(row.iloc[index]) else ''
```

### 6.3 处理完成后
```python
# 运行数据验证检查清单
# 打印各类型记录的空值统计
# 确认异常数据在合理范围内
```

---

## 七、版本历史

| 日期 | 版本 | 更新内容 |
|------|------|---------|
| 2026-01-30 | 1.0 | 初始版本，总结数据处理验证规范 |
| 2026-01-30 | 1.1 | 添加项目概述、快速开始、开发指南 |
