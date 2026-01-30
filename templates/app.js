/**
 * 差旅数据分析系统 - 前端应用逻辑
 *
 * 功能特性：
 * - 数据嵌入HTML（单文件部署）
 * - 加载进度显示
 * - 部门安全控制（默认：教培业务中心，密码：201212）
 * - 部门数据隔离
 * - 数据筛选（时间范围、类型、员工搜索）
 * - 统计汇总
 * - 图表展示（趋势、分布、排名）
 * - 分类型表格（机票/酒店/火车/用车）
 * - 表格排序、筛选、分页
 * - CSV导出
 */

// ========================================
// 部门安全控制类
// ========================================

class DeptSecurity {
    constructor() {
        this.defaultDept = '教培业务中心';
        this.password = '201212';
        this.currentDept = this.defaultDept;
        this.isUnlocked = false;
        this.pendingDept = null;
    }

    /**
     * 尝试切换部门
     */
    switchDept(deptName) {
        // 切换到默认部门不需要密码
        if (deptName === this.defaultDept) {
            this.currentDept = deptName;
            return { success: true, dept: deptName };
        }

        // 切换到当前部门
        if (deptName === this.currentDept) {
            return { success: true, dept: deptName };
        }

        // 其他部门需要密码验证
        if (!this.isUnlocked) {
            this.pendingDept = deptName;
            return { success: false, needPassword: true };
        }

        this.currentDept = deptName;
        return { success: true, dept: deptName };
    }

    /**
     * 验证密码
     */
    verifyPassword(input) {
        if (input === this.password) {
            this.isUnlocked = true;
            if (this.pendingDept) {
                this.currentDept = this.pendingDept;
                this.pendingDept = null;
            }
            return true;
        }
        return false;
    }

    /**
     * 获取当前部门显示标签
     */
    getDeptLabel() {
        if (this.currentDept === '全部') {
            return '全部数据 (已解锁)';
        }
        const suffix = this.currentDept === this.defaultDept ? ' [默认]' : ' [已切换]';
        return `${this.currentDept}${suffix}`;
    }

    /**
     * 重置到默认部门
     */
    reset() {
        this.currentDept = this.defaultDept;
        this.isUnlocked = false;
        this.pendingDept = null;
    }
}

// ========================================
// 差旅数据分析主应用类
// ========================================

class TravelAnalysisApp {
    constructor() {
        this.data = null;
        this.filteredData = [];
        this.security = new DeptSecurity();
        this.currentType = 'flight';
        this.currentChart = 'trend';
        this.currentPage = 1;
        this.pageSize = 25;
        this.sortColumn = null;
        this.sortDirection = 'asc';
        this.chartInstance = null;
        this.columnFilters = {}; // 列筛选值

        // 表格列定义
        this.tableColumns = {
            flight: [
                { key: 'passenger', label: '乘机人', sortable: true },
                { key: 'bookTime', label: '预定时间', sortable: true },
                { key: 'flightNo', label: '航班号', sortable: true },
                { key: 'departTime', label: '起飞时间', sortable: true },
                { key: 'fromCity', label: '出发地', sortable: true },
                { key: 'toCity', label: '目的地', sortable: true },
                { key: 'price', label: '票价', sortable: true, format: 'currency' },
                { key: 'cabinClass', label: '舱位类型', sortable: true },
                { key: 'airline', label: '航空公司', sortable: true },
                { key: 'source', label: '来源', sortable: true }
            ],
            hotel: [
                { key: 'employee', label: '员工', sortable: true },
                { key: 'checkInTime', label: '入住时间', sortable: true },
                { key: 'checkOutTime', label: '离开时间', sortable: true },
                { key: 'city', label: '城市', sortable: true },
                { key: 'hotelName', label: '酒店名称', sortable: true },
                { key: 'roomType', label: '房型', sortable: true },
                { key: 'price', label: '房价', sortable: true, format: 'currency' },
                { key: 'isShared', label: '是否合住', sortable: true, format: 'boolean' },
                { key: 'source', label: '来源', sortable: true }
            ],
            train: [
                { key: 'employee', label: '员工', sortable: true },
                { key: 'trainNo', label: '车次', sortable: true },
                { key: 'seat', label: '座席', sortable: true },
                { key: 'departTime', label: '发车时间', sortable: true },
                { key: 'fromCity', label: '出发城市', sortable: true },
                { key: 'toCity', label: '到达城市', sortable: true },
                { key: 'price', label: '票价', sortable: true, format: 'currency' },
                { key: 'source', label: '来源', sortable: true }
            ],
            car: [
                { key: 'passenger', label: '乘车人', sortable: true },
                { key: 'pickupTime', label: '上车时间', sortable: true },
                { key: 'dropoffTime', label: '下车时间', sortable: true },
                { key: 'carType', label: '用车类型', sortable: true },
                { key: 'origin', label: '出发地', sortable: true, format: 'address' },
                { key: 'destination', label: '目的地', sortable: true, format: 'address' },
                { key: 'distance', label: '里程(km)', sortable: true, format: 'number' },
                { key: 'totalAmount', label: '金额', sortable: true, format: 'currency' },
                { key: 'source', label: '来源', sortable: true },
                { key: 'provider', label: '服务方', sortable: true }
            ]
        };

        this.init();
    }

    /**
     * 初始化应用（带进度条）
     */
    async init() {
        const loadingScreen = document.getElementById('loadingScreen');
        const loadingProgress = document.getElementById('loadingProgress');
        const loadingText = document.getElementById('loadingText');

        try {
            // 使用嵌入的数据，但显示处理进度
            const updateProgress = (percent, message) => {
                if (loadingProgress) {
                    loadingProgress.classList.remove('animating');
                    loadingProgress.style.width = percent + '%';
                }
                if (loadingText) {
                    loadingText.textContent = message;
                }
            };

            // 模拟数据处理进度（增加延迟让用户看到进度）
            updateProgress(10, '正在读取数据...');
            await new Promise(resolve => setTimeout(resolve, 100));

            // 检查数据是否存在
            if (typeof TRAVEL_DATA === 'undefined') {
                throw new Error('数据未加载，请确保数据已正确嵌入');
            }

            updateProgress(30, '正在解析数据...');
            await new Promise(resolve => setTimeout(resolve, 150));

            updateProgress(60, '正在构建索引...');
            this.data = TRAVEL_DATA;
            this.filteredData = [...this.data.records];
            await new Promise(resolve => setTimeout(resolve, 150));

            updateProgress(85, '正在初始化界面...');
            this.initUI();
            this.bindEvents();
            await new Promise(resolve => setTimeout(resolve, 200));

            updateProgress(100, '完成');
            await new Promise(resolve => setTimeout(resolve, 300));

            // 隐藏加载界面
            if (loadingScreen) {
                loadingScreen.classList.add('hidden');
            }

            // 应用筛选
            this.applyFilters();

        } catch (error) {
            console.error('初始化失败:', error);
            if (loadingScreen) {
                loadingScreen.classList.add('error');
                if (loadingText) {
                    loadingText.textContent = `加载失败: ${error.message}`;
                }
            }
        }
    }

    /**
     * 初始化UI
     */
    initUI() {
        // 更新部门显示
        this.updateDeptDisplay();

        // 初始化部门选择器
        this.initDeptSelector();

        // 更新时间显示
        const updateTime = document.getElementById('updateTime');
        if (updateTime && this.data.lastUpdate) {
            updateTime.textContent = dayjs(this.data.lastUpdate).format('YYYY-MM-DD HH:mm:ss');
        }

        // 隐藏加载遮罩
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) {
            loadingOverlay.classList.remove('active');
        }
    }

    /**
     * 初始化部门选择器
     */
    initDeptSelector() {
        const deptSelect = document.getElementById('deptSelect');
        if (!deptSelect || !this.data.summary) return;

        // 获取所有部门
        const depts = Object.keys(this.data.summary.byDept || {}).sort();

        // 清空现有选项
        deptSelect.innerHTML = '<option value="">当前部门</option>';

        // 添加部门选项
        depts.forEach(dept => {
            const option = document.createElement('option');
            option.value = dept;
            option.textContent = dept;
            deptSelect.appendChild(option);
        });

        // 添加"全部"选项（需要密码）
        const allOption = document.createElement('option');
        allOption.value = '全部';
        allOption.textContent = '全部数据 (需密码)';
        deptSelect.appendChild(allOption);
    }

    /**
     * 绑定事件
     */
    bindEvents() {
        // 部门选择
        const deptSelect = document.getElementById('deptSelect');
        deptSelect?.addEventListener('change', (e) => this.handleDeptChange(e.target.value));

        // 时间范围选择
        const timeRange = document.getElementById('timeRange');
        timeRange?.addEventListener('change', () => this.applyFilters());

        // 来源范围选择
        const sourceRange = document.getElementById('sourceRange');
        sourceRange?.addEventListener('change', () => this.applyFilters());

        // 搜索输入
        const searchInput = document.getElementById('searchInput');
        let searchTimeout;
        searchInput?.addEventListener('input', (e) => {
            clearTimeout(searchTimeout);
            searchTimeout = setTimeout(() => this.applyFilters(), 300);
        });

        // 图表Tab切换
        document.querySelectorAll('.chart-tab').forEach(tab => {
            tab.addEventListener('click', () => this.switchChart(tab.dataset.chart));
        });

        // 表格Tab切换
        document.querySelectorAll('.table-tab').forEach(tab => {
            tab.addEventListener('click', () => this.switchTableType(tab.dataset.type));
        });

        // 导出按钮
        const exportBtn = document.getElementById('exportBtn');
        exportBtn?.addEventListener('click', () => this.exportData());

        // 密码验证
        const passwordConfirm = document.getElementById('passwordConfirm');
        const passwordInput = document.getElementById('passwordInput');
        passwordConfirm?.addEventListener('click', () => {
            const input = passwordInput.value;
            if (this.security.verifyPassword(input)) {
                this.closeModal('passwordModal');
                // 密码验证成功后，更新部门显示并刷新数据
                this.updateDeptDisplay();
                this.applyFilters();
                // 重置选择器
                const deptSelect = document.getElementById('deptSelect');
                if (deptSelect) deptSelect.value = '';
            } else {
                alert('密码错误');
            }
        });

        // 模态框关闭
        document.querySelectorAll('[data-close]').forEach(btn => {
            btn.addEventListener('click', () => this.closeModal(btn.dataset.close));
        });

        // 模态框遮罩关闭
        document.querySelectorAll('.modal-overlay').forEach(overlay => {
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) {
                    const modal = overlay.closest('.modal');
                    if (modal) this.closeModal(modal.id);
                }
            });
        });

        // 回车键确认密码
        passwordInput?.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                passwordConfirm.click();
            }
        });
    }

    /**
     * 处理部门切换
     */
    handleDeptChange(deptName) {
        if (!deptName) return;

        const result = this.security.switchDept(deptName);

        if (result.success) {
            this.updateDeptDisplay();
            this.applyFilters();
        } else if (result.needPassword) {
            this.showModal('passwordModal');
        }

        // 重置选择器
        const deptSelect = document.getElementById('deptSelect');
        if (deptSelect) deptSelect.value = '';
    }

    /**
     * 更新部门显示
     */
    updateDeptDisplay() {
        const currentDept = document.getElementById('currentDept');
        if (currentDept) {
            currentDept.textContent = this.security.getDeptLabel();
        }
    }

    /**
     * 应用筛选条件
     */
    applyFilters() {
        let records = [...this.data.records];

        // 部门筛选
        const currentDept = this.security.currentDept;
        if (currentDept !== '全部') {
            records = records.filter(r => r.deptLevel1 === currentDept);
        }

        // 时间范围筛选
        const timeRange = document.getElementById('timeRange')?.value;
        if (timeRange && timeRange !== 'all') {
            const now = dayjs();
            records = records.filter(r => {
                const dateStr = this.parseRecordDate(r);
                if (!dateStr) return false;
                const date = dayjs(dateStr);

                switch (timeRange) {
                    case 'month':
                        return date.isSame(now, 'month');
                    case 'quarter':
                        return date.isSame(now, 'quarter');
                    case 'year':
                        return date.isSame(now, 'year');
                    default:
                        return true;
                }
            });
        }

        // 员工搜索
        const searchTerm = document.getElementById('searchInput')?.value.toLowerCase();
        if (searchTerm) {
            records = records.filter(r => {
                const name = r.type === 'flight' || r.type === 'car'
                    ? r.passenger
                    : r.employee;
                return name && name.toLowerCase().includes(searchTerm);
            });
        }

        // 来源筛选
        const sourceRange = document.getElementById('sourceRange')?.value;
        if (sourceRange && sourceRange !== 'all') {
            records = records.filter(r => r.source === sourceRange);
        }

        this.filteredData = records;
        this.updateOverview();
        this.updateChart();
        this.updateTable();
    }

    /**
     * 解析记录日期
     */
    parseRecordDate(record) {
        switch (record.type) {
            case 'flight':
                return record.departTime?.split(' ')[0];
            case 'hotel':
                return record.checkInTime?.split(' ')[0];
            case 'train':
                return record.departTime?.split(' ')[0];
            case 'car':
                return record.pickupTime?.split(' ')[0];
            default:
                return '';
        }
    }

    /**
     * 更新概览卡片
     */
    updateOverview() {
        const summary = {
            totalAmount: 0,
            flight: { amount: 0, count: 0 },
            hotel: { amount: 0, count: 0 },
            train: { amount: 0, count: 0 },
            car: { amount: 0, count: 0 },
            totalRecords: this.filteredData.length
        };

        const sources = new Set();

        this.filteredData.forEach(r => {
            const amount = this.parseRecordAmount(r);
            summary.totalAmount += amount;
            sources.add(r.source);

            if (summary[r.type]) {
                summary[r.type].amount += amount;
                summary[r.type].count += 1;
            }
        });

        // 更新显示
        this.updateStat('totalAmount', summary.totalAmount, 'currency');
        this.updateStat('flightAmount', summary.flight.amount, 'currency');
        this.updateStat('flightCount', summary.flight.count, 'number');
        this.updateStat('hotelAmount', summary.hotel.amount, 'currency');
        this.updateStat('hotelCount', summary.hotel.count, 'number');
        this.updateStat('trainAmount', summary.train.amount, 'currency');
        this.updateStat('trainCount', summary.train.count, 'number');
        this.updateStat('carAmount', summary.car.amount, 'currency');
        this.updateStat('carCount', summary.car.count, 'number');
        this.updateStat('totalRecords', summary.totalRecords, 'number');

        const recordSource = document.getElementById('recordSource');
        if (recordSource) {
            recordSource.textContent = `${sources.size} 个数据源`;
        }

        const recordSummary = document.getElementById('recordSummary');
        if (recordSummary) {
            recordSummary.textContent = `共 ${summary.totalRecords} 条记录`;
        }
    }

    /**
     * 解析记录金额
     */
    parseRecordAmount(record) {
        switch (record.type) {
            case 'flight':
                return record.price || 0;
            case 'hotel':
                return record.price || 0;
            case 'train':
                return record.price || 0;
            case 'car':
                return record.totalAmount || 0;
            default:
                return 0;
        }
    }

    /**
     * 更新统计卡片
     */
    updateStat(id, value, format) {
        const el = document.getElementById(id);
        if (!el) return;

        switch (format) {
            case 'currency':
                el.textContent = `¥${value.toLocaleString('zh-CN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
                break;
            case 'number':
                el.textContent = value.toLocaleString('zh-CN');
                break;
            default:
                el.textContent = value;
        }
    }

    /**
     * 切换图表
     */
    switchChart(chartType) {
        this.currentChart = chartType;

        // 更新Tab状态
        document.querySelectorAll('.chart-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.chart === chartType);
        });

        this.updateChart();
    }

    /**
     * 更新图表
     */
    updateChart() {
        if (!this.chartInstance) {
            const chartDom = document.getElementById('mainChart');
            if (chartDom) {
                this.chartInstance = echarts.init(chartDom);
            }
        }

        const option = this.getChartOption();
        if (this.chartInstance && option) {
            this.chartInstance.setOption(option, true);
        }
    }

    /**
     * 获取图表配置
     */
    getChartOption() {
        const records = this.filteredData;
        const colorPalette = ['#1a56db', '#06b6d4', '#7c3aed', '#10b981', '#f59e0b'];

        switch (this.currentChart) {
            case 'trend':
                return this.getTrendChartOption(records, colorPalette);
            case 'distribution':
                return this.getDistributionChartOption(records, colorPalette);
            case 'ranking':
                return this.getRankingChartOption(records, colorPalette);
            case 'department':
                return this.getDepartmentChartOption(records, colorPalette);
            default:
                return {};
        }
    }

    /**
     * 格式化金额为整数
     */
    formatCurrency(value) {
        return '¥' + Number(value).toLocaleString('zh-CN', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    }

    /**
     * 趋势图
     */
    getTrendChartOption(records, colors) {
        // 按月统计
        const monthlyData = {};
        records.forEach(r => {
            const dateStr = this.parseRecordDate(r);
            // 验证日期格式，只处理有效的日期（YYYY-MM或YYYY-MM-DD格式）
            if (!dateStr || !dateStr.match(/^\d{4}-\d{2}/)) {
                return; // 跳过无效日期
            }
            const month = dateStr.substring(0, 7);
            if (!monthlyData[month]) {
                monthlyData[month] = { flight: 0, hotel: 0, train: 0, car: 0 };
            }
            monthlyData[month][r.type] += this.parseRecordAmount(r);
        });

        // 过滤掉无效的月份键
        const months = Object.keys(monthlyData).filter(m => m.match(/^\d{4}-\d{2}$/)).sort();

        return {
            tooltip: {
                trigger: 'axis',
                axisPointer: { type: 'shadow' },
                formatter: (params) => {
                    let result = params[0].name + '<br/>';
                    params.forEach(item => {
                        result += `${item.seriesName}: ${this.formatCurrency(item.value)}<br/>`;
                    });
                    return result;
                }
            },
            legend: {
                data: ['机票', '酒店', '火车', '用车'],
                bottom: 0
            },
            grid: {
                left: '3%',
                right: '4%',
                bottom: '15%',
                top: '10%',
                containLabel: true
            },
            xAxis: {
                type: 'category',
                data: months
            },
            yAxis: {
                type: 'value',
                axisLabel: {
                    formatter: (value) => this.formatCurrency(value)
                }
            },
            series: [
                { name: '机票', type: 'bar', stack: 'total', data: months.map(m => monthlyData[m].flight), itemStyle: { color: colors[0] } },
                { name: '酒店', type: 'bar', stack: 'total', data: months.map(m => monthlyData[m].hotel), itemStyle: { color: colors[2] } },
                { name: '火车', type: 'bar', stack: 'total', data: months.map(m => monthlyData[m].train), itemStyle: { color: colors[3] } },
                { name: '用车', type: 'bar', stack: 'total', data: months.map(m => monthlyData[m].car), itemStyle: { color: colors[4] } }
            ]
        };
    }

    /**
     * 类型分布图
     */
    getDistributionChartOption(records, colors) {
        const typeData = { flight: 0, hotel: 0, train: 0, car: 0 };
        records.forEach(r => {
            typeData[r.type] += this.parseRecordAmount(r);
        });

        return {
            tooltip: {
                trigger: 'item',
                formatter: (params) => `${params.name}: ${this.formatCurrency(params.value)} (${params.percent}%)`
            },
            legend: {
                orient: 'vertical',
                right: '10%',
                top: 'center',
                data: ['机票', '酒店', '火车', '用车']
            },
            series: [{
                type: 'pie',
                radius: ['40%', '70%'],
                center: ['35%', '50%'],
                data: [
                    { value: typeData.flight, name: '机票', itemStyle: { color: colors[0] } },
                    { value: typeData.hotel, name: '酒店', itemStyle: { color: colors[2] } },
                    { value: typeData.train, name: '火车', itemStyle: { color: colors[3] } },
                    { value: typeData.car, name: '用车', itemStyle: { color: colors[4] } }
                ],
                label: {
                    formatter: (params) => this.formatCurrency(params.value)
                }
            }]
        };
    }

    /**
     * 员工排名图
     */
    getRankingChartOption(records, colors) {
        const employeeData = {};
        records.forEach(r => {
            const name = r.type === 'flight' || r.type === 'car' ? r.passenger : r.employee;
            if (!name) return;
            if (!employeeData[name]) employeeData[name] = 0;
            employeeData[name] += this.parseRecordAmount(r);
        });

        const sorted = Object.entries(employeeData)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15);

        return {
            tooltip: {
                trigger: 'axis',
                axisPointer: { type: 'shadow' },
                formatter: (params) => `${params[0].name}: ${this.formatCurrency(params[0].value)}`
            },
            grid: {
                left: '20%',
                right: '5%',
                top: '5%',
                bottom: '5%',
                containLabel: true
            },
            xAxis: {
                type: 'value',
                axisLabel: {
                    formatter: (value) => this.formatCurrency(value)
                }
            },
            yAxis: {
                type: 'category',
                data: sorted.map(s => s[0]).reverse()
            },
            series: [{
                type: 'bar',
                data: sorted.map(s => s[1]).reverse(),
                itemStyle: {
                    color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
                        { offset: 0, color: colors[0] },
                        { offset: 1, color: colors[1] }
                    ])
                }
            }]
        };
    }

    /**
     * 部门对比图
     */
    getDepartmentChartOption(records, colors) {
        const deptData = {};
        records.forEach(r => {
            const dept = r.deptLevel1 || '未知部门';
            if (!deptData[dept]) deptData[dept] = 0;
            deptData[dept] += this.parseRecordAmount(r);
        });

        const sorted = Object.entries(deptData).sort((a, b) => b[1] - a[1]);

        return {
            tooltip: {
                trigger: 'axis',
                axisPointer: { type: 'shadow' },
                formatter: (params) => `${params[0].name}: ${this.formatCurrency(params[0].value)}`
            },
            grid: {
                left: '3%',
                right: '4%',
                bottom: '10%',
                top: '5%',
                containLabel: true
            },
            xAxis: {
                type: 'category',
                data: sorted.map(s => s[0]),
                axisLabel: {
                    rotate: 45
                }
            },
            yAxis: {
                type: 'value',
                axisLabel: {
                    formatter: (value) => this.formatCurrency(value)
                }
            },
            series: [{
                type: 'bar',
                data: sorted.map(s => s[1]),
                itemStyle: {
                    color: new echarts.graphic.LinearGradient(0, 1, 0, 0, [
                        { offset: 0, color: colors[0] },
                        { offset: 1, color: colors[1] }
                    ])
                }
            }]
        };
    }

    /**
     * 切换表格类型
     */
    switchTableType(type) {
        this.currentType = type;
        this.currentPage = 1;
        this.sortColumn = null;
        this.sortDirection = 'asc';
        this.columnFilters = {}; // 清空列筛选

        // 更新Tab状态
        document.querySelectorAll('.table-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.type === type);
        });

        this.updateTable();
    }

    /**
     * 更新表格
     */
    updateTable() {
        const typeRecords = this.filteredData.filter(r => r.type === this.currentType);
        const columns = this.tableColumns[this.currentType] || [];

        // 应用列筛选
        let displayRecords = [...typeRecords];
        Object.entries(this.columnFilters).forEach(([key, value]) => {
            if (value && value.trim()) {
                const filter = value.trim().toLowerCase();
                displayRecords = displayRecords.filter(r => {
                    const cellValue = this.getCellValue(r, key);
                    return String(cellValue).toLowerCase().includes(filter);
                });
            }
        });

        // 排序
        if (this.sortColumn) {
            displayRecords.sort((a, b) => {
                let aVal = this.getCellValue(a, this.sortColumn);
                let bVal = this.getCellValue(b, this.sortColumn);

                if (typeof aVal === 'number' && typeof bVal === 'number') {
                    return this.sortDirection === 'asc' ? aVal - bVal : bVal - aVal;
                }
                return this.sortDirection === 'asc'
                    ? String(aVal).localeCompare(String(bVal))
                    : String(bVal).localeCompare(String(aVal));
            });
        }

        // 分页
        const totalPages = Math.ceil(displayRecords.length / this.pageSize) || 1;
        const start = (this.currentPage - 1) * this.pageSize;
        const end = Math.min(start + this.pageSize, displayRecords.length);
        const pageRecords = displayRecords.slice(start, end);

        // 渲染表头
        this.renderTableHeader(columns);

        // 渲染表体
        this.renderTableBody(pageRecords, columns);

        // 更新分页信息
        this.updatePagination(displayRecords.length, start, end, totalPages);
    }

    /**
     * 渲染表头
     */
    renderTableHeader(columns) {
        const thead = document.getElementById('tableHeader');
        if (!thead) return;

        // 生成表头HTML（包含列筛选）
        thead.innerHTML = columns.map(col => `
            <th class="${col.sortable ? 'sortable' : ''} ${this.sortColumn === col.key ? `sort-${this.sortDirection}` : ''}"
                data-column="${col.key}">
                <div class="th-content">
                    <span class="th-label">${col.label}</span>
                    ${this.sortColumn === col.key ? `<span class="sort-indicator">${this.sortDirection === 'asc' ? '↑' : '↓'}</span>` : ''}
                </div>
                <select class="column-filter" data-column="${col.key}">
                    <option value="">全部</option>
                </select>
            </th>
        `).join('');

        // 绑定排序事件
        thead.querySelectorAll('th.sortable').forEach(th => {
            const label = th.querySelector('.th-label');
            label.addEventListener('click', () => {
                const column = th.dataset.column;
                if (this.sortColumn === column) {
                    this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    this.sortColumn = column;
                    this.sortDirection = 'asc';
                }
                this.updateTable();
            });
        });

        // 绑定列筛选事件
        thead.querySelectorAll('.column-filter').forEach(select => {
            select.addEventListener('change', (e) => {
                const column = e.target.dataset.column;
                this.columnFilters[column] = e.target.value;
                this.currentPage = 1; // 重置到第一页
                this.updateTable();
            });

            // 阻止筛选选择框的点击事件冒泡到排序
            select.addEventListener('click', (e) => {
                e.stopPropagation();
            });
        });

        // 填充下拉选项
        this.populateColumnFilters(columns);
    }

    /**
     * 填充列筛选下拉选项
     */
    populateColumnFilters(columns) {
        const typeRecords = this.filteredData.filter(r => r.type === this.currentType);

        columns.forEach(col => {
            const select = document.querySelector(`.column-filter[data-column="${col.key}"]`);
            if (!select) return;

            // 收集该列的所有唯一值
            const values = new Set();
            typeRecords.forEach(r => {
                const value = this.getCellValue(r, col.key);
                if (value !== undefined && value !== null && value !== '') {
                    values.add(String(value));
                }
            });

            // 按字母顺序排序
            const sortedValues = Array.from(values).sort((a, b) => a.localeCompare(b, 'zh-CN'));

            // 保留"全部"选项，添加其他选项
            const currentValue = this.columnFilters[col.key] || '';
            sortedValues.forEach(value => {
                const option = document.createElement('option');
                option.value = value;
                option.textContent = value;
                if (value === currentValue) {
                    option.selected = true;
                }
                select.appendChild(option);
            });
        });
    }

    /**
     * 渲染表体
     */
    renderTableBody(records, columns) {
        const tbody = document.getElementById('tableBody');
        if (!tbody) return;

        if (records.length === 0) {
            tbody.innerHTML = '<tr><td colspan="' + columns.length + '" style="text-align:center;color:#94a3b8;">暂无数据</td></tr>';
            return;
        }

        tbody.innerHTML = records.map(record => `
            <tr data-record='${JSON.stringify(record).replace(/'/g, '&#39;')}'>
                ${columns.map(col => `
                    <td class="${col.key === 'passenger' || col.key === 'employee' ? 'cell-clickable' : ''}">
                        ${this.formatCellValue(record, col)}
                    </td>
                `).join('')}
            </tr>
        `).join('');

        // 绑定行点击事件
        tbody.querySelectorAll('tr[data-record]').forEach(tr => {
            tr.addEventListener('click', (e) => {
                if (e.target.classList.contains('cell-clickable')) {
                    const record = JSON.parse(tr.dataset.record);
                    this.showDetailModal(record);
                }
            });
        });
    }

    /**
     * 获取单元格值
     */
    getCellValue(record, key) {
        if (key === 'origin' || key === 'destination') {
            const addr = record[key] || {};
            return addr.city || '';
        }
        return record[key] || '';
    }

    /**
     * 格式化单元格值
     */
    formatCellValue(record, col) {
        const value = this.getCellValue(record, col.key);

        switch (col.format) {
            case 'currency':
                return `¥${Number(value).toLocaleString('zh-CN', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
            case 'boolean':
                return value ? '是' : '否';
            case 'address':
                const addr = record[col.key] || {};
                return [addr.city, addr.district, addr.address].filter(Boolean).join(' / ');
            case 'number':
                return Number(value).toLocaleString('zh-CN');
            default:
                return value;
        }
    }

    /**
     * 更新分页
     */
    updatePagination(total, start, end, totalPages) {
        // 更新显示信息
        document.getElementById('showStart').textContent = total > 0 ? start + 1 : 0;
        document.getElementById('showEnd').textContent = end;
        document.getElementById('totalItems').textContent = total;

        // 渲染分页按钮
        const pagination = document.getElementById('pagination');
        if (!pagination) return;

        let buttons = '';

        // 上一页
        buttons += `<button ${this.currentPage === 1 ? 'disabled' : ''} data-page="${this.currentPage - 1}">‹</button>`;

        // 页码
        const maxButtons = 5;
        let startPage = Math.max(1, this.currentPage - Math.floor(maxButtons / 2));
        let endPage = Math.min(totalPages, startPage + maxButtons - 1);

        if (endPage - startPage < maxButtons - 1) {
            startPage = Math.max(1, endPage - maxButtons + 1);
        }

        if (startPage > 1) {
            buttons += `<button data-page="1">1</button>`;
            if (startPage > 2) buttons += '<span>...</span>';
        }

        for (let i = startPage; i <= endPage; i++) {
            buttons += `<button class="${i === this.currentPage ? 'active' : ''}" data-page="${i}">${i}</button>`;
        }

        if (endPage < totalPages) {
            if (endPage < totalPages - 1) buttons += '<span>...</span>';
            buttons += `<button data-page="${totalPages}">${totalPages}</button>`;
        }

        // 下一页
        buttons += `<button ${this.currentPage === totalPages ? 'disabled' : ''} data-page="${this.currentPage + 1}">›</button>`;

        pagination.innerHTML = buttons;

        // 绑定分页事件
        pagination.querySelectorAll('button[data-page]').forEach(btn => {
            btn.addEventListener('click', () => {
                const page = parseInt(btn.dataset.page);
                if (page >= 1 && page <= totalPages) {
                    this.currentPage = page;
                    this.updateTable();
                }
            });
        });
    }

    /**
     * 显示详情模态框
     */
    showDetailModal(record) {
        const modal = document.getElementById('detailModal');
        const title = document.getElementById('detailTitle');
        const content = document.getElementById('detailContent');

        if (!modal || !content) return;

        const typeLabels = {
            flight: '机票详情',
            hotel: '酒店详情',
            train: '火车详情',
            car: '用车详情'
        };

        title.textContent = typeLabels[record.type] || '详情';

        let html = '';
        for (const [key, value] of Object.entries(record)) {
            if (key === 'source' || key === 'type') continue;

            const label = this.getFieldLabel(record.type, key);
            if (!label) continue;

            let displayValue = value;
            if (typeof value === 'object') {
                displayValue = [value.city, value.district, value.address].filter(Boolean).join(' / ');
            }

            html += `
                <div class="detail-row">
                    <div class="detail-label">${label}</div>
                    <div class="detail-value ${key === 'price' || key === 'totalAmount' ? 'highlight' : ''}">${displayValue || '-'}</div>
                </div>
            `;
        }

        content.innerHTML = html;
        this.showModal('detailModal');
    }

    /**
     * 获取字段标签
     */
    getFieldLabel(type, key) {
        const columns = this.tableColumns[type] || [];
        const col = columns.find(c => c.key === key);
        return col?.label || null;
    }

    /**
     * 显示模态框
     */
    showModal(modalId) {
        const modal = document.getElementById(modalId);
        if (modal) modal.classList.add('active');
    }

    /**
     * 关闭模态框
     */
    closeModal(modalId) {
        const modal = document.getElementById(modalId);
        if (modal) modal.classList.remove('active');
    }

    /**
     * 导出数据
     */
    exportData() {
        const typeRecords = this.filteredData.filter(r => r.type === this.currentType);
        const columns = this.tableColumns[this.currentType] || [];

        if (typeRecords.length === 0) {
            alert('当前没有数据可导出');
            return;
        }

        // 生成CSV
        const headers = columns.map(c => c.label).join(',');
        const rows = typeRecords.map(r => {
            return columns.map(c => {
                const value = this.getCellValue(r, c.key);
                return `"${String(value).replace(/"/g, '""')}"`;
            }).join(',');
        });

        const csv = [headers, ...rows].join('\n');

        // 添加BOM以支持Excel正确显示中文
        const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `差旅数据_${this.currentType}_${dayjs().format('YYYYMMDD')}.csv`;
        link.click();
        URL.revokeObjectURL(url);
    }

    /**
     * 显示错误
     */
    showError(message) {
        const loadingScreen = document.getElementById('loadingScreen');
        const loadingText = document.getElementById('loadingText');
        if (loadingScreen) {
            loadingScreen.classList.add('error');
        }
        if (loadingText) {
            loadingText.textContent = message;
        }
    }
}

// 响应式图表
window.addEventListener('resize', () => {
    if (window.app && window.app.chartInstance) {
        window.app.chartInstance.resize();
    }
});
