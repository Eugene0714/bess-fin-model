/**
 * 德国独立储能电站投资测算系统 - 财务模型计算引擎
 * @description 包含所有财务计算逻辑、报表生成和敏感性分析
 * @version 1.0
 */

// ==================== 全局变量 ====================

/** @type {Object} 存储所有计算结果 */
let calculationResults = {
    capex: {},
    opex: [],
    revenue: [],
    depreciation: [],
    loan: [],
    incomeStatement: [],
    balanceSheet: [],
    cashFlow: [],
    indicators: {}
};

/** @type {Object} 图表实例 */
let charts = {
    revenue: null,
    sensitivity: null
};

// ==================== 数字格式化工具函数 ====================

/**
 * 格式化数字为千分位显示
 * @param {number} num - 要格式化的数字
 * @param {number} decimals - 小数位数，默认2位
 * @returns {string} 格式化后的字符串
 */
function formatNumber(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) {
        return '-';
    }
    return num.toLocaleString('de-DE', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}

/**
 * 格式化数字为千分位显示（无小数）
 * @param {number} num - 要格式化的数字
 * @returns {string} 格式化后的字符串
 */
function formatInt(num) {
    if (num === null || num === undefined || isNaN(num)) {
        return '-';
    }
    return Math.round(num).toLocaleString('de-DE');
}

/**
 * 格式化百分比
 * @param {number} num - 要格式化的数字（已经是百分比形式，如15表示15%）
 * @param {number} decimals - 小数位数，默认2位
 * @returns {string} 格式化后的字符串
 */
function formatPercent(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) {
        return '-';
    }
    return num.toLocaleString('de-DE', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}

/**
 * 格式化货币（万EUR）
 * @param {number} num - 要格式化的数字
 * @param {number} decimals - 小数位数，默认2位
 * @returns {string} 格式化后的字符串
 */
function formatCurrency(num, decimals = 2) {
    if (num === null || num === undefined || isNaN(num)) {
        return '-';
    }
    return num.toLocaleString('de-DE', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}

// ==================== 初始化 ====================

/**
 * 页面加载完成后初始化
 */
document.addEventListener('DOMContentLoaded', function() {
    // 显示当前登录用户名
    const username = sessionStorage.getItem('username');
    const userNameElement = document.getElementById('currentUserName');
    if (userNameElement && username) {
        userNameElement.textContent = username;
    }
    
    initNavigation();
    initInputListeners();
    initSpotPriceTable();
    // 初始化衰减模式显示
    onDegradationModeChange();
    calculateAll();
});

/**
 * 用户登出
 */
function logout() {
    if (confirm('确定要退出登录吗？')) {
        sessionStorage.removeItem('isLoggedIn');
        sessionStorage.removeItem('username');
        window.location.href = 'login.html';
    }
}

/**
 * 显示帮助模态框
 */
function showHelp() {
    const modal = document.getElementById('helpModal');
    if (modal) {
        modal.style.display = 'flex';
        modal.classList.add('show');
        // 重置到第一个标签页
        switchHelpTab('quickstart');
        // 阻止背景滚动
        document.body.style.overflow = 'hidden';
    }
}

/**
 * 关闭帮助模态框
 */
function closeHelp() {
    const modal = document.getElementById('helpModal');
    if (modal) {
        modal.style.display = 'none';
        modal.classList.remove('show');
        // 恢复背景滚动
        document.body.style.overflow = '';
    }
}

/**
 * 切换帮助标签页
 * @param {string} tabName 标签页名称
 */
function switchHelpTab(tabName) {
    // 隐藏所有标签页内容
    const allContents = document.querySelectorAll('.help-tab-content');
    allContents.forEach(content => {
        content.classList.remove('active');
        content.style.display = 'none';
    });
    
    // 移除所有标签的active类
    const allTabs = document.querySelectorAll('.help-tab');
    allTabs.forEach(tab => {
        tab.classList.remove('active');
    });
    
    // 显示选中的标签页内容
    const targetContent = document.getElementById('help-' + tabName);
    if (targetContent) {
        targetContent.classList.add('active');
        targetContent.style.display = 'block';
    }
    
    // 激活选中的标签（通过查找包含tabName的onclick属性）
    allTabs.forEach(tab => {
        const onclickAttr = tab.getAttribute('onclick');
        if (onclickAttr && onclickAttr.includes(tabName)) {
            tab.classList.add('active');
        }
    });
}

// 点击模态框外部关闭
document.addEventListener('DOMContentLoaded', function() {
    const modal = document.getElementById('helpModal');
    if (modal) {
        modal.addEventListener('click', function(e) {
            if (e.target === modal) {
                closeHelp();
            }
        });
        
        // ESC键关闭
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && modal.classList.contains('show')) {
                closeHelp();
            }
        });
    }
});

/**
 * 初始化导航
 */
function initNavigation() {
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const tabId = this.getAttribute('data-tab');
            
            // 切换活动状态
            navLinks.forEach(l => l.classList.remove('active'));
            this.classList.add('active');
            
            // 切换内容
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.remove('active');
            });
            document.getElementById(tabId).classList.add('active');
        });
    });
}

/**
 * 防抖函数
 * @param {Function} func 要执行的函数
 * @param {number} wait 等待时间（毫秒）
 * @returns {Function} 防抖后的函数
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// 防抖后的计算函数（500ms延迟，避免频繁计算）
const debouncedCalculateAll = debounce(() => {
    try {
        calculateAll();
    } catch (e) {
        console.error('自动计算出错:', e);
    }
}, 500);

/**
 * 初始化输入监听器
 */
function initInputListeners() {
    // 功率和容量变化时更新储能时长
    document.getElementById('power_mw').addEventListener('input', updateDuration);
    document.getElementById('capacity_mwh').addEventListener('input', updateDuration);
    
    // 充放电效率变化时更新RTE
    document.getElementById('charge_efficiency').addEventListener('input', updateRTE);
    document.getElementById('discharge_efficiency').addEventListener('input', updateRTE);
    
    // 使用事件委托监听所有输入变化 - 确保动态和静态元素都能被监听
    document.body.addEventListener('input', function(event) {
        const target = event.target;
        // 只处理input和select元素
        if (target.tagName === 'INPUT' || target.tagName === 'SELECT') {
            // 检查是否在需要监听的区域内
            if (target.closest('#parameters') || target.closest('#opex') || target.closest('#revenue')) {
                debouncedCalculateAll();
            }
        }
    });
    
    document.body.addEventListener('change', function(event) {
        const target = event.target;
        if (target.tagName === 'INPUT' || target.tagName === 'SELECT') {
            if (target.closest('#parameters') || target.closest('#opex') || target.closest('#revenue')) {
                debouncedCalculateAll();
            }
        }
    });
    
    console.log('[财务模型] 已启用参数自动计算功能（事件委托模式）');
}

/**
 * 更新储能时长
 */
function updateDuration() {
    const power = parseFloat(document.getElementById('power_mw').value) || 1;
    const capacity = parseFloat(document.getElementById('capacity_mwh').value) || 0;
    document.getElementById('duration_hours').value = (capacity / power).toFixed(1);
}

/**
 * 更新系统综合效率
 */
function updateRTE() {
    const chargeEff = parseFloat(document.getElementById('charge_efficiency').value) || 95;
    const dischargeEff = parseFloat(document.getElementById('discharge_efficiency').value) || 95;
    const rte = (chargeEff * dischargeEff / 100).toFixed(2);
    document.getElementById('rte').value = rte;
}

/**
 * 电池衰减参数数据库
 * 根据电池厂家和型号提供默认衰减参数
 */
const BATTERY_DEGRADATION_DB = {
    // CATL 宁德时代
    'CATL_EnerOne_Plus': {
        mode: 'linear',
        degradation_rate: 2.2,
        degradation_first_year: 2.8,
        degradation_annual_decrease: 0.08,
        cycles_per_degradation: 1.8,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'CATL_EnerC_Plus': {
        mode: 'linear',
        degradation_rate: 2.0,
        degradation_first_year: 2.5,
        degradation_annual_decrease: 0.07,
        cycles_per_degradation: 1.6,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'CATL_EnerD': {
        mode: 'nonlinear',
        degradation_rate: 2.3,
        degradation_first_year: 3.0,
        degradation_annual_decrease: 0.1,
        cycles_per_degradation: 2.0,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'CATL_TENER': {
        mode: 'linear',
        degradation_rate: 1.8,
        degradation_first_year: 2.2,
        degradation_annual_decrease: 0.06,
        cycles_per_degradation: 1.5,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // BYD 比亚迪
    'BYD_MC_Cube': {
        mode: 'linear',
        degradation_rate: 2.4,
        degradation_first_year: 3.2,
        degradation_annual_decrease: 0.12,
        cycles_per_degradation: 2.1,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'BYD_Cube_Pro': {
        mode: 'linear',
        degradation_rate: 2.1,
        degradation_first_year: 2.7,
        degradation_annual_decrease: 0.09,
        cycles_per_degradation: 1.7,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'BYD_BatteryBox': {
        mode: 'nonlinear',
        degradation_rate: 2.5,
        degradation_first_year: 3.5,
        degradation_annual_decrease: 0.15,
        cycles_per_degradation: 2.2,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // EVE 亿纬锂能
    'EVE_LF560K': {
        mode: 'linear',
        degradation_rate: 2.3,
        degradation_first_year: 2.9,
        degradation_annual_decrease: 0.09,
        cycles_per_degradation: 1.9,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'EVE_LF280K': {
        mode: 'linear',
        degradation_rate: 2.4,
        degradation_first_year: 3.0,
        degradation_annual_decrease: 0.1,
        cycles_per_degradation: 2.0,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'EVE_LF314K': {
        mode: 'linear',
        degradation_rate: 2.2,
        degradation_first_year: 2.8,
        degradation_annual_decrease: 0.08,
        cycles_per_degradation: 1.8,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // REPT 瑞浦兰钧
    'REPT_320Ah': {
        mode: 'linear',
        degradation_rate: 2.5,
        degradation_first_year: 3.2,
        degradation_annual_decrease: 0.11,
        cycles_per_degradation: 2.1,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'REPT_345Ah': {
        mode: 'linear',
        degradation_rate: 2.4,
        degradation_first_year: 3.0,
        degradation_annual_decrease: 0.1,
        cycles_per_degradation: 2.0,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // HiTHIUM 海辰储能
    'HiTHIUM_314Ah': {
        mode: 'linear',
        degradation_rate: 2.3,
        degradation_first_year: 2.9,
        degradation_annual_decrease: 0.09,
        cycles_per_degradation: 1.9,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'HiTHIUM_560Ah': {
        mode: 'linear',
        degradation_rate: 2.2,
        degradation_first_year: 2.8,
        degradation_annual_decrease: 0.08,
        cycles_per_degradation: 1.8,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // Gotion 国轩高科
    'Gotion_280Ah': {
        mode: 'linear',
        degradation_rate: 2.5,
        degradation_first_year: 3.2,
        degradation_annual_decrease: 0.12,
        cycles_per_degradation: 2.1,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    'Gotion_314Ah': {
        mode: 'linear',
        degradation_rate: 2.4,
        degradation_first_year: 3.0,
        degradation_annual_decrease: 0.1,
        cycles_per_degradation: 2.0,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // Samsung 三星SDI
    'Samsung_E3': {
        mode: 'nonlinear',
        degradation_rate: 2.0,
        degradation_first_year: 2.5,
        degradation_annual_decrease: 0.08,
        cycles_per_degradation: 1.7,
        annual_cycles: 365,
        capacity_threshold: 80
    },
    // LGES LG新能源
    'LGES_RESU': {
        mode: 'nonlinear',
        degradation_rate: 1.9,
        degradation_first_year: 2.4,
        degradation_annual_decrease: 0.07,
        cycles_per_degradation: 1.6,
        annual_cycles: 365,
        capacity_threshold: 80
    }
};

/**
 * 更新电池柜规格（根据选择的型号）
 */
function updateBatterySpecs() {
    const select = document.getElementById('battery_model');
    const selectedOption = select.options[select.selectedIndex];
    
    if (selectedOption && selectedOption.value !== 'CUSTOM') {
        const capacity = selectedOption.getAttribute('data-capacity');
        const price = selectedOption.getAttribute('data-price');
        const model = selectedOption.value;
        
        if (capacity) {
            document.getElementById('battery_cabinet_capacity').value = capacity;
            // 同步更新容量下拉选择
            syncBatteryCapacitySelect(capacity);
        }
        if (price) {
            document.getElementById('battery_unit_price').value = price;
        }
        
        // 自动填充电池衰减参数
        updateDegradationParams(model);
        
        // 自动计算电池柜数量
        autoCalcBatteryCount();
    }
}

/**
 * 根据电池型号自动填充衰减参数
 * @param {string} model 电池型号
 */
function updateDegradationParams(model) {
    const params = BATTERY_DEGRADATION_DB[model];
    if (!params) {
        return; // 如果没有找到对应参数，不更新
    }
    
    // 更新衰减模式
    const modeSelect = document.getElementById('degradation_mode');
    if (modeSelect) {
        modeSelect.value = params.mode;
        // 触发模式切换以显示/隐藏相关输入框
        onDegradationModeChange();
    }
    
    // 更新年衰减率
    const degradationRateInput = document.getElementById('degradation_rate');
    if (degradationRateInput) {
        degradationRateInput.value = params.degradation_rate;
    }
    
    // 更新首年衰减率（非线性模式）
    const firstYearInput = document.getElementById('degradation_first_year');
    if (firstYearInput) {
        firstYearInput.value = params.degradation_first_year;
    }
    
    // 更新年衰减率递减幅度（非线性模式）
    const annualDecreaseInput = document.getElementById('degradation_annual_decrease');
    if (annualDecreaseInput) {
        annualDecreaseInput.value = params.degradation_annual_decrease;
    }
    
    // 更新每1000次循环衰减率（循环次数模式）
    const cyclesPerDegInput = document.getElementById('cycles_per_degradation');
    if (cyclesPerDegInput) {
        cyclesPerDegInput.value = params.cycles_per_degradation;
    }
    
    // 更新年运行循环次数
    const annualCyclesInput = document.getElementById('annual_cycles');
    if (annualCyclesInput) {
        annualCyclesInput.value = params.annual_cycles;
    }
    
    // 更新容量保持率阈值
    const thresholdInput = document.getElementById('capacity_threshold');
    if (thresholdInput) {
        thresholdInput.value = params.capacity_threshold;
    }
    
    console.log(`[电池衰减] 已自动填充 ${model} 的衰减参数`);
}

/**
 * 衰减模式切换时的处理
 */
function onDegradationModeChange() {
    const mode = document.getElementById('degradation_mode').value;
    const firstYearGroup = document.getElementById('degradation_first_year_group');
    const annualDecreaseGroup = document.getElementById('degradation_annual_decrease_group');
    const cyclesGroup = document.getElementById('cycles_per_degradation_group');
    
    // 重置所有组的显示状态
    if (firstYearGroup) firstYearGroup.style.display = 'none';
    if (annualDecreaseGroup) annualDecreaseGroup.style.display = 'none';
    if (cyclesGroup) cyclesGroup.style.display = 'none';
    
    // 根据选择的模式显示相应的输入框
    if (mode === 'nonlinear') {
        if (firstYearGroup) firstYearGroup.style.display = 'block';
        if (annualDecreaseGroup) annualDecreaseGroup.style.display = 'block';
    } else if (mode === 'cycle_based') {
        if (cyclesGroup) cyclesGroup.style.display = 'block';
    }
}

/**
 * 同步电池容量下拉选择
 * @param {string} capacity 容量值
 */
function syncBatteryCapacitySelect(capacity) {
    const select = document.getElementById('battery_capacity_select');
    const options = select.options;
    let found = false;
    
    for (let i = 0; i < options.length; i++) {
        if (options[i].value === capacity) {
            select.selectedIndex = i;
            found = true;
            break;
        }
    }
    
    if (!found) {
        // 设置为自定义
        for (let i = 0; i < options.length; i++) {
            if (options[i].value === 'custom') {
                select.selectedIndex = i;
                break;
            }
        }
    }
}

/**
 * 电池容量下拉选择变化
 */
function onBatteryCapacitySelect() {
    const select = document.getElementById('battery_capacity_select');
    const value = select.value;
    
    if (value !== 'custom') {
        document.getElementById('battery_cabinet_capacity').value = value;
        autoCalcBatteryCount();
    }
}

/**
 * 电池容量手动输入变化
 */
function onBatteryCapacityInput() {
    const value = document.getElementById('battery_cabinet_capacity').value;
    syncBatteryCapacitySelect(value);
}

/**
 * 自动计算电池柜数量
 */
function autoCalcBatteryCount() {
    const totalCapacity = parseFloat(document.getElementById('capacity_mwh').value) || 200;
    const cabinetCapacity = parseFloat(document.getElementById('battery_cabinet_capacity').value) || 3.44;
    
    if (cabinetCapacity > 0) {
        const count = Math.ceil(totalCapacity / cabinetCapacity);
        document.getElementById('battery_cabinet_count').value = count;
    }
}

/**
 * 更新PCS规格（根据选择的型号）
 */
function updatePCSSpecs() {
    const select = document.getElementById('pcs_model');
    const selectedOption = select.options[select.selectedIndex];
    
    if (selectedOption && selectedOption.value !== 'CUSTOM') {
        const power = selectedOption.getAttribute('data-power');
        const price = selectedOption.getAttribute('data-price');
        
        if (power) {
            document.getElementById('pcs_power').value = power;
            // 同步更新功率下拉选择
            syncPCSPowerSelect(power);
        }
        if (price) {
            document.getElementById('pcs_unit_price').value = price;
        }
        
        // 自动计算PCS数量
        autoCalcPCSCount();
    }
}

/**
 * 同步PCS功率下拉选择
 * @param {string} power 功率值
 */
function syncPCSPowerSelect(power) {
    const select = document.getElementById('pcs_power_select');
    const options = select.options;
    let found = false;
    
    for (let i = 0; i < options.length; i++) {
        if (options[i].value === power) {
            select.selectedIndex = i;
            found = true;
            break;
        }
    }
    
    if (!found) {
        for (let i = 0; i < options.length; i++) {
            if (options[i].value === 'custom') {
                select.selectedIndex = i;
                break;
            }
        }
    }
}

/**
 * PCS功率下拉选择变化
 */
function onPCSPowerSelect() {
    const select = document.getElementById('pcs_power_select');
    const value = select.value;
    
    if (value !== 'custom') {
        document.getElementById('pcs_power').value = value;
        autoCalcPCSCount();
    }
}

/**
 * PCS功率手动输入变化
 */
function onPCSPowerInput() {
    const value = document.getElementById('pcs_power').value;
    syncPCSPowerSelect(value);
}

/**
 * 自动计算PCS数量
 */
function autoCalcPCSCount() {
    const totalPower = parseFloat(document.getElementById('power_mw').value) || 100;
    const pcsPower = parseFloat(document.getElementById('pcs_power').value) || 3.45;
    
    if (pcsPower > 0) {
        const count = Math.ceil(totalPower / pcsPower);
        document.getElementById('pcs_count').value = count;
        // 同时更新中压变数量
        document.getElementById('mv_transformer_count').value = count;
    }
}

/**
 * 更新中压变压器规格（根据选择的型号）
 */
function updateMVTransformerSpecs() {
    const select = document.getElementById('mv_transformer_model');
    const selectedOption = select.options[select.selectedIndex];
    
    if (selectedOption && selectedOption.value !== 'CUSTOM') {
        const voltage = selectedOption.getAttribute('data-voltage');
        const capacity = selectedOption.getAttribute('data-capacity');
        const price = selectedOption.getAttribute('data-price');
        
        if (voltage) {
            document.getElementById('mv_transformer_voltage').value = voltage;
        }
        if (capacity) {
            document.getElementById('mv_transformer_capacity').value = capacity;
            syncMVCapacitySelect(capacity);
        }
        if (price) {
            document.getElementById('mv_transformer_price').value = price;
        }
    }
}

/**
 * 同步中压变容量下拉选择
 * @param {string} capacity 容量值
 */
function syncMVCapacitySelect(capacity) {
    const select = document.getElementById('mv_capacity_select');
    const options = select.options;
    let found = false;
    
    for (let i = 0; i < options.length; i++) {
        if (options[i].value === capacity) {
            select.selectedIndex = i;
            found = true;
            break;
        }
    }
    
    if (!found) {
        for (let i = 0; i < options.length; i++) {
            if (options[i].value === 'custom') {
                select.selectedIndex = i;
                break;
            }
        }
    }
}

/**
 * 中压变容量下拉选择变化
 */
function onMVCapacitySelect() {
    const select = document.getElementById('mv_capacity_select');
    const value = select.value;
    
    if (value !== 'custom') {
        document.getElementById('mv_transformer_capacity').value = value;
    }
}

/**
 * 中压变容量手动输入变化
 */
function onMVCapacityInput() {
    const value = document.getElementById('mv_transformer_capacity').value;
    syncMVCapacitySelect(value);
}

/**
 * 自动计算中压变数量（与PCS数量同步）
 */
function autoCalcMVCount() {
    const pcsCount = parseInt(document.getElementById('pcs_count').value) || 29;
    document.getElementById('mv_transformer_count').value = pcsCount;
}

/**
 * 更新升压变规格（根据选择的型号）
 */
function updateHVTransformerSpecs() {
    const select = document.getElementById('hv_transformer_model');
    const selectedOption = select.options[select.selectedIndex];
    
    if (selectedOption && selectedOption.value !== 'CUSTOM') {
        const capacity = selectedOption.getAttribute('data-capacity');
        const price = selectedOption.getAttribute('data-price');
        
        if (capacity) {
            document.getElementById('hv_transformer_capacity').value = capacity;
            syncHVCapacitySelect(capacity);
        }
        if (price) {
            document.getElementById('hv_transformer_price').value = price;
        }
    }
}

/**
 * 同步升压变容量下拉选择
 * @param {string} capacity 容量值
 */
function syncHVCapacitySelect(capacity) {
    const select = document.getElementById('hv_capacity_select');
    const options = select.options;
    let found = false;
    
    for (let i = 0; i < options.length; i++) {
        if (options[i].value === capacity) {
            select.selectedIndex = i;
            found = true;
            break;
        }
    }
    
    if (!found) {
        for (let i = 0; i < options.length; i++) {
            if (options[i].value === 'custom') {
                select.selectedIndex = i;
                break;
            }
        }
    }
}

/**
 * 升压变容量下拉选择变化
 */
function onHVCapacitySelect() {
    const select = document.getElementById('hv_capacity_select');
    const value = select.value;
    
    if (value !== 'custom') {
        document.getElementById('hv_transformer_capacity').value = value;
    }
}

/**
 * 升压变容量手动输入变化
 */
function onHVCapacityInput() {
    const value = document.getElementById('hv_transformer_capacity').value;
    syncHVCapacitySelect(value);
}

/**
 * 自动计算升压变数量（根据总功率和单台容量）
 */
function autoCalcHVCount() {
    const totalPower = parseFloat(document.getElementById('power_mw').value) || 100;
    const hvCapacity = parseFloat(document.getElementById('hv_transformer_capacity').value) || 120;
    
    if (hvCapacity > 0) {
        // 升压变容量一般要大于总功率，通常只需要1台
        const count = Math.ceil(totalPower / hvCapacity);
        document.getElementById('hv_transformer_count').value = Math.max(1, count);
    }
}

/**
 * 初始化现货价格表
 * @description 默认使用德国2025年现货市场套利预期收入
 */
function initSpotPriceTable() {
    const years = parseInt(document.getElementById('operation_years').value) || 20;
    const tbody = document.getElementById('spot_price_body');
    tbody.innerHTML = '';
    
    // 德国2025年现货市场套利预期: 30-45k€/MW/年，取35k作为基准
    const basePrice = 35000;
    // 年增长率1.5%
    const escalation = 0.015;
    const halfYears = Math.ceil(years / 2);
    
    for (let i = 0; i < halfYears; i++) {
        const row = document.createElement('tr');
        const year1 = i + 1;
        const year2 = i + 1 + halfYears;
        
        // 按年增长率计算各年价格
        const price1 = Math.round(basePrice * Math.pow(1 + escalation, year1 - 1));
        const price2 = Math.round(basePrice * Math.pow(1 + escalation, year2 - 1));
        
        row.innerHTML = `
            <td>第${year1}年</td>
            <td><input type="number" id="spot_price_${year1}" value="${price1}" min="0"></td>
            ${year2 <= years ? `
                <td>第${year2}年</td>
                <td><input type="number" id="spot_price_${year2}" value="${price2}" min="0"></td>
            ` : '<td></td><td></td>'}
        `;
        tbody.appendChild(row);
    }
}

/**
 * 快速填充现货价格
 */
function fillSpotPrices() {
    const basePrice = parseFloat(document.getElementById('spot_base_price').value) || 45000;  // 默认45000€/MW/年
    const escalation = parseFloat(document.getElementById('spot_escalation').value) || 0;
    const years = parseInt(document.getElementById('operation_years').value) || 20;
    
    for (let i = 1; i <= years; i++) {
        const input = document.getElementById(`spot_price_${i}`);
        if (input) {
            const price = basePrice * Math.pow(1 + escalation / 100, i - 1);
            input.value = Math.round(price);
        }
    }
}

// ==================== 参数获取 ====================

/**
 * 获取所有输入参数
 * @returns {Object} 参数对象
 */
/**
 * 获取所有输入参数
 * @description 默认值基于2025年德国储能市场主流数据
 * @returns {Object} 参数对象
 */
function getParameters() {
    return {
        // 基础参数 - 德国大型独立储能项目典型配置
        power_mw: parseFloat(document.getElementById('power_mw').value) || 100,
        capacity_mwh: parseFloat(document.getElementById('capacity_mwh').value) || 200,
        operation_years: parseInt(document.getElementById('operation_years').value) || 20,
        initial_capacity_pct: parseFloat(document.getElementById('initial_capacity_pct').value) || 100,
        
        // 融资参数 - 德国项目融资市场标准
        equity_ratio: parseFloat(document.getElementById('equity_ratio').value) / 100 || 0.25,
        loan_years: parseInt(document.getElementById('loan_years').value) || 12,
        loan_rate: parseFloat(document.getElementById('loan_rate').value) / 100 || 0.045,
        grace_period: parseInt(document.getElementById('grace_period').value) || 1,
        repayment_method: document.getElementById('repayment_method').value,
        
        // 建设期与通胀参数
        construction_period: parseFloat(document.getElementById('construction_period').value) || 1,
        construction_fund_usage: parseFloat(document.getElementById('construction_fund_usage').value) / 100 || 0.5,
        inflation_rate: parseFloat(document.getElementById('inflation_rate').value) / 100 || 0.02,
        
        // 折旧参数
        depreciation_years: parseInt(document.getElementById('depreciation_years').value) || 15,
        salvage_rate: parseFloat(document.getElementById('salvage_rate').value) / 100 || 0.05,
        depreciation_method: document.getElementById('depreciation_method').value,
        amortization_years: parseInt(document.getElementById('amortization_years').value) || 20,
        
        // 效率参数
        charge_efficiency: parseFloat(document.getElementById('charge_efficiency').value) / 100 || 0.95,
        discharge_efficiency: parseFloat(document.getElementById('discharge_efficiency').value) / 100 || 0.95,
        degradation_rate: parseFloat(document.getElementById('degradation_rate').value) / 100 || 0.025,
        annual_cycles: parseInt(document.getElementById('annual_cycles').value) || 365,
        
        // 税费参数 - 德国税制
        corporate_tax_rate: parseFloat(document.getElementById('corporate_tax_rate').value) / 100 || 0.15,
        solidarity_tax_rate: parseFloat(document.getElementById('solidarity_tax_rate').value) / 100 || 0.055,
        trade_tax_rate: parseFloat(document.getElementById('trade_tax_rate').value) / 100 || 0.14,
        vat_rate: parseFloat(document.getElementById('vat_rate').value) / 100 || 0.19,
        other_tax_rate: parseFloat(document.getElementById('other_tax_rate').value) / 100 || 0,
        
        // Tolling参数 - 德国储能Tolling市场
        tolling_years: parseInt(document.getElementById('tolling_years').value) || 10,
        tolling_ratio: parseFloat(document.getElementById('tolling_ratio').value) / 100 || 0.8,
        tolling_price: parseFloat(document.getElementById('tolling_price').value) || 95,
        tolling_escalation: parseFloat(document.getElementById('tolling_escalation').value) / 100 || 0.02,
        
        // 主设备参数 - 2025年德国市场主流价格
        battery_cabinet_capacity: parseFloat(document.getElementById('battery_cabinet_capacity').value) || 5.0,
        battery_cabinet_count: parseInt(document.getElementById('battery_cabinet_count').value) || 40,
        battery_unit_price: parseFloat(document.getElementById('battery_unit_price').value) || 75,
        pcs_power: parseFloat(document.getElementById('pcs_power').value) || 5.5,
        pcs_count: parseInt(document.getElementById('pcs_count').value) || 19,
        pcs_unit_price: parseFloat(document.getElementById('pcs_unit_price').value) || 28,
        mv_transformer_capacity: parseFloat(document.getElementById('mv_transformer_capacity').value) || 6300,
        mv_transformer_count: parseInt(document.getElementById('mv_transformer_count').value) || 19,
        mv_transformer_price: parseFloat(document.getElementById('mv_transformer_price').value) || 35000,
        hv_transformer_capacity: parseFloat(document.getElementById('hv_transformer_capacity').value) || 120,
        hv_transformer_count: parseInt(document.getElementById('hv_transformer_count').value) || 1,
        hv_transformer_price: parseFloat(document.getElementById('hv_transformer_price').value) || 750000,
        
        // 辅助设备参数
        ems_cost: parseFloat(document.getElementById('ems_cost').value) || 200000,
        scada_cost: parseFloat(document.getElementById('scada_cost').value) || 120000,
        switchgear_price: parseFloat(document.getElementById('switchgear_price').value) || 18000,
        switchgear_count: parseInt(document.getElementById('switchgear_count').value) || 25,
        collector_line_cost: parseFloat(document.getElementById('collector_line_cost').value) || 250000,
        thermal_cost: parseFloat(document.getElementById('thermal_cost').value) || 20,
        fire_protection_cost: parseFloat(document.getElementById('fire_protection_cost').value) || 12,
        
        // 电网接入参数 - 德国110kV电网接入
        substation_cost: parseFloat(document.getElementById('substation_cost').value) || 800000,
        grid_line_cost: parseFloat(document.getElementById('grid_line_cost').value) || 500000,
        grid_study_cost: parseFloat(document.getElementById('grid_study_cost').value) || 80000,
        metering_cost: parseFloat(document.getElementById('metering_cost').value) || 100000,
        
        // 土地与基建参数
        land_acquisition_cost: parseFloat(document.getElementById('land_acquisition_cost').value) || 50,
        concrete_cost: parseFloat(document.getElementById('concrete_cost').value) || 25,
        fence_cost: parseFloat(document.getElementById('fence_cost').value) || 80000,
        road_cost: parseFloat(document.getElementById('road_cost').value) || 120000,
        drainage_cost: parseFloat(document.getElementById('drainage_cost').value) || 50000,
        
        // 安装与施工参数
        installation_cost_pct: parseFloat(document.getElementById('installation_cost_pct').value) / 100 || 0.06,
        construction_mgmt_pct: parseFloat(document.getElementById('construction_mgmt_pct').value) / 100 || 0.025,
        commissioning_cost: parseFloat(document.getElementById('commissioning_cost').value) || 150000,
        
        // 建设期保险参数
        car_insurance_pct: parseFloat(document.getElementById('car_insurance_pct').value) / 100 || 0.003,
        ear_insurance_pct: parseFloat(document.getElementById('ear_insurance_pct').value) / 100 || 0.002,
        cargo_insurance_pct: parseFloat(document.getElementById('cargo_insurance_pct').value) / 100 || 0.0015,
        liability_insurance: parseFloat(document.getElementById('liability_insurance').value) || 50000,
        
        // 开发与业主费用参数
        spv_acquisition_cost: parseFloat(document.getElementById('spv_acquisition_cost').value) || 50000,
        permit_cost: parseFloat(document.getElementById('permit_cost').value) || 180000,
        environmental_cost: parseFloat(document.getElementById('environmental_cost').value) || 60000,
        project_mgmt_pct: parseFloat(document.getElementById('project_mgmt_pct').value) / 100 || 0.02,
        legal_cost: parseFloat(document.getElementById('legal_cost').value) || 100000,
        engineering_pct: parseFloat(document.getElementById('engineering_pct').value) / 100 || 0.025,
        contingency_pct: parseFloat(document.getElementById('contingency_pct').value) / 100 || 0.05,
        
        // 拆除准备金参数（按年摊销）
        decommissioning_total: parseFloat(document.getElementById('decommissioning_total').value) || 500000,
        
        // OPEX参数 - 德国运营市场价格
        opex_technical: parseFloat(document.getElementById('opex_technical').value) || 6,
        opex_technical_esc: parseFloat(document.getElementById('opex_technical_esc').value) / 100 || 0.02,
        opex_insurance: parseFloat(document.getElementById('opex_insurance').value) / 100 || 0.004,
        opex_insurance_esc: parseFloat(document.getElementById('opex_insurance_esc').value) / 100 || 0.015,
        opex_grid: parseFloat(document.getElementById('opex_grid').value) || 12000,
        opex_grid_esc: parseFloat(document.getElementById('opex_grid_esc').value) / 100 || 0.02,
        opex_land: parseFloat(document.getElementById('opex_land').value) || 60000,
        opex_land_esc: parseFloat(document.getElementById('opex_land_esc').value) / 100 || 0.02,
        opex_commercial: parseFloat(document.getElementById('opex_commercial').value) || 4000,
        opex_commercial_esc: parseFloat(document.getElementById('opex_commercial_esc').value) / 100 || 0.02,
        opex_other: parseFloat(document.getElementById('opex_other').value) || 1500,
        opex_other_esc: parseFloat(document.getElementById('opex_other_esc').value) / 100 || 0.02
    };
}

/**
 * 获取现货价格数组
 * @param {number} years 年数
 * @returns {number[]} 现货价格数组
 */
function getSpotPrices(years) {
    const prices = [];
    // 德国2025年现货市场套利预期基准: 35k€/MW/年
    const defaultBasePrice = 35000;
    for (let i = 1; i <= years; i++) {
        const input = document.getElementById(`spot_price_${i}`);
        prices.push(input ? parseFloat(input.value) || defaultBasePrice : defaultBasePrice);
    }
    return prices;
}

// ==================== CAPEX计算 ====================

/**
 * 计算CAPEX明细
 * @param {Object} params 参数
 * @returns {Object} CAPEX明细
 */
function calculateCapex(params) {
    const capex = {};
    
    // ========== 一、主设备费用 ==========
    // 电池系统
    capex.battery = params.capacity_mwh * 1000 * params.battery_unit_price / 10000; // 万EUR
    // PCS系统
    capex.pcs = params.power_mw * 1000 * params.pcs_unit_price / 10000;
    // 中压变压器
    capex.mv_transformer = params.mv_transformer_count * params.mv_transformer_price / 10000;
    // 升压变压器
    capex.hv_transformer = params.hv_transformer_count * params.hv_transformer_price / 10000;
    
    // ========== 二、辅助设备费用 ==========
    // EMS能量管理系统
    capex.ems = params.ems_cost / 10000;
    // SCADA监控系统
    capex.scada = params.scada_cost / 10000;
    // 开关柜
    capex.switchgear = params.switchgear_count * params.switchgear_price / 10000;
    // 集电线路
    capex.collector_line = params.collector_line_cost / 10000;
    // 热管理系统
    capex.thermal = params.capacity_mwh * 1000 * params.thermal_cost / 10000;
    // 消防系统
    capex.fire_protection = params.capacity_mwh * 1000 * params.fire_protection_cost / 10000;
    
    // 主设备小计
    const equipmentSubtotal = capex.battery + capex.pcs + capex.mv_transformer + capex.hv_transformer +
                              capex.ems + capex.scada + capex.switchgear + capex.collector_line +
                              capex.thermal + capex.fire_protection;
    
    // ========== 三、电网接入费用 ==========
    // 变电站建设/扩容
    capex.substation = params.substation_cost / 10000;
    // 接入线路
    capex.grid_line = params.grid_line_cost / 10000;
    // 并网申请与研究费
    capex.grid_study = params.grid_study_cost / 10000;
    // 计量与保护设备
    capex.metering = params.metering_cost / 10000;
    
    // 电网接入小计
    capex.grid_connection_subtotal = capex.substation + capex.grid_line + capex.grid_study + capex.metering;
    
    // ========== 四、土地与基础建设 ==========
    // 土地获取成本
    capex.land_acquisition = params.power_mw * 1000 * params.land_acquisition_cost / 10000;
    // 混凝土基础
    capex.concrete = params.power_mw * 1000 * params.concrete_cost / 10000;
    // 围栏与安防
    capex.fence = params.fence_cost / 10000;
    // 道路建设
    capex.road = params.road_cost / 10000;
    // 排水系统
    capex.drainage = params.drainage_cost / 10000;
    
    // 土地与基建小计
    capex.civil_subtotal = capex.land_acquisition + capex.concrete + capex.fence + capex.road + capex.drainage;
    
    // ========== 五、安装与施工费用 ==========
    // 机电安装
    capex.installation = equipmentSubtotal * params.installation_cost_pct;
    // 施工管理费
    capex.construction_mgmt = equipmentSubtotal * params.construction_mgmt_pct;
    // 调试费用
    capex.commissioning = params.commissioning_cost / 10000;
    
    // 安装施工小计
    capex.installation_subtotal = capex.installation + capex.construction_mgmt + capex.commissioning;
    
    // ========== 六、建设期保险费用 ==========
    // 建设工程一切险(CAR)
    capex.car_insurance = equipmentSubtotal * params.car_insurance_pct;
    // 安装工程一切险(EAR)
    capex.ear_insurance = equipmentSubtotal * params.ear_insurance_pct;
    // 货物运输保险
    capex.cargo_insurance = equipmentSubtotal * params.cargo_insurance_pct;
    // 第三方责任险
    capex.liability_insurance = params.liability_insurance / 10000;
    
    // 保险费小计
    capex.insurance_subtotal = capex.car_insurance + capex.ear_insurance + capex.cargo_insurance + capex.liability_insurance;
    
    // ========== 七、开发与业主费用 ==========
    // SPV公司收购成本
    capex.spv_acquisition = params.spv_acquisition_cost / 10000;
    // 许可与规划费
    capex.permit = params.permit_cost / 10000;
    // 环境咨询费
    capex.environmental = params.environmental_cost / 10000;
    // 法律咨询费
    capex.legal = params.legal_cost / 10000;
    // 工程设计费
    capex.engineering = equipmentSubtotal * params.engineering_pct;
    
    // 小计（用于计算项目管理费和不可预见费）
    const subtotalBeforeMgmt = equipmentSubtotal + capex.grid_connection_subtotal + 
                               capex.civil_subtotal + capex.installation_subtotal +
                               capex.insurance_subtotal +
                               capex.spv_acquisition + capex.permit + capex.environmental + capex.legal + capex.engineering;
    
    // 项目管理费（按总投资比例）
    capex.project_mgmt = subtotalBeforeMgmt * params.project_mgmt_pct;
    
    // 开发费用小计
    capex.dev_subtotal = capex.spv_acquisition + capex.permit + capex.environmental + capex.legal + capex.engineering + capex.project_mgmt;
    
    // ========== 八、不可预见费 ==========
    const subtotalBeforeContingency = subtotalBeforeMgmt + capex.project_mgmt;
    capex.contingency = subtotalBeforeContingency * params.contingency_pct;
    
    // ========== 总计 ==========
    // 拆除准备金改为逐年摊销，不计入CAPEX
    capex.total = subtotalBeforeContingency + capex.contingency;
    
    // 建设期利息计算
    // 建设期利息 = 贷款金额 × 贷款利率 × 建设期 × 资金占用比例
    const loanAmount = capex.total * (1 - params.equity_ratio);
    capex.construction_interest = loanAmount * params.loan_rate * params.construction_period * params.construction_fund_usage;
    
    // 动态投资总额
    capex.dynamic_total = capex.total + capex.construction_interest;
    
    // 保存各类小计用于显示
    capex.equipment_subtotal = equipmentSubtotal;
    
    // 添加别名字段，用于折旧、现金流、资产负债表计算
    // dev_cost 对应开发费用小计，land 对应土地获取成本
    capex.dev_cost = capex.dev_subtotal;
    capex.land = capex.land_acquisition;
    
    return capex;
}

/**
 * 更新CAPEX表格显示
 * @param {Object} capex CAPEX数据
 * @param {Object} params 参数
 */
function updateCapexTable(capex, params) {
    const tbody = document.getElementById('capex_breakdown');
    tbody.innerHTML = `
        <!-- 一、主设备 -->
        <tr class="section-header"><td colspan="4">一、主设备</td></tr>
        <tr>
            <td>电池系统</td>
            <td>${formatInt(params.battery_unit_price)} EUR/kWh</td>
            <td>${formatInt(params.capacity_mwh)} MWh</td>
            <td>${formatNumber(capex.battery)}</td>
        </tr>
        <tr>
            <td>PCS系统</td>
            <td>${formatInt(params.pcs_unit_price)} EUR/kW</td>
            <td>${formatInt(params.power_mw)} MW</td>
            <td>${formatNumber(capex.pcs)}</td>
        </tr>
        <tr>
            <td>中压变压器</td>
            <td>${formatInt(params.mv_transformer_price)} EUR/台</td>
            <td>${formatInt(params.mv_transformer_count)} 台</td>
            <td>${formatNumber(capex.mv_transformer)}</td>
        </tr>
        <tr>
            <td>升压变压器</td>
            <td>${formatInt(params.hv_transformer_price)} EUR/台</td>
            <td>${formatInt(params.hv_transformer_count)} 台</td>
            <td>${formatNumber(capex.hv_transformer)}</td>
        </tr>
        
        <!-- 二、辅助设备 -->
        <tr class="section-header"><td colspan="4">二、辅助设备</td></tr>
        <tr>
            <td>EMS能量管理系统</td>
            <td>-</td>
            <td>1 套</td>
            <td>${formatNumber(capex.ems)}</td>
        </tr>
        <tr>
            <td>SCADA监控系统</td>
            <td>-</td>
            <td>1 套</td>
            <td>${formatNumber(capex.scada)}</td>
        </tr>
        <tr>
            <td>开关柜</td>
            <td>${formatInt(params.switchgear_price)} EUR/面</td>
            <td>${formatInt(params.switchgear_count)} 面</td>
            <td>${formatNumber(capex.switchgear)}</td>
        </tr>
        <tr>
            <td>集电线路</td>
            <td>-</td>
            <td>1 套</td>
            <td>${formatNumber(capex.collector_line)}</td>
        </tr>
        <tr>
            <td>热管理系统</td>
            <td>${formatInt(params.thermal_cost)} EUR/kWh</td>
            <td>${formatInt(params.capacity_mwh)} MWh</td>
            <td>${formatNumber(capex.thermal)}</td>
        </tr>
        <tr>
            <td>消防系统</td>
            <td>${formatInt(params.fire_protection_cost)} EUR/kWh</td>
            <td>${formatInt(params.capacity_mwh)} MWh</td>
            <td>${formatNumber(capex.fire_protection)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">设备费小计</td>
            <td>${formatNumber(capex.equipment_subtotal)}</td>
        </tr>
        
        <!-- 三、电网接入 -->
        <tr class="section-header"><td colspan="4">三、电网接入</td></tr>
        <tr>
            <td>变电站建设/扩容</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.substation)}</td>
        </tr>
        <tr>
            <td>接入线路</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.grid_line)}</td>
        </tr>
        <tr>
            <td>并网申请与研究费</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.grid_study)}</td>
        </tr>
        <tr>
            <td>计量与保护设备</td>
            <td>-</td>
            <td>1 套</td>
            <td>${formatNumber(capex.metering)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">电网接入小计</td>
            <td>${formatNumber(capex.grid_connection_subtotal)}</td>
        </tr>
        
        <!-- 四、土地与基建 -->
        <tr class="section-header"><td colspan="4">四、土地与基础建设</td></tr>
        <tr>
            <td>土地获取成本</td>
            <td>${formatInt(params.land_acquisition_cost)} EUR/kW</td>
            <td>${formatInt(params.power_mw)} MW</td>
            <td>${formatNumber(capex.land_acquisition)}</td>
        </tr>
        <tr>
            <td>混凝土基础</td>
            <td>${formatInt(params.concrete_cost)} EUR/kW</td>
            <td>${formatInt(params.power_mw)} MW</td>
            <td>${formatNumber(capex.concrete)}</td>
        </tr>
        <tr>
            <td>围栏与安防</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.fence)}</td>
        </tr>
        <tr>
            <td>道路建设</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.road)}</td>
        </tr>
        <tr>
            <td>排水系统</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.drainage)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">土地与基建小计</td>
            <td>${formatNumber(capex.civil_subtotal)}</td>
        </tr>
        
        <!-- 五、安装与施工 -->
        <tr class="section-header"><td colspan="4">五、安装与施工</td></tr>
        <tr>
            <td>机电安装</td>
            <td>${formatNumber(params.installation_cost_pct * 100, 1)}% 设备费</td>
            <td>-</td>
            <td>${formatNumber(capex.installation)}</td>
        </tr>
        <tr>
            <td>施工管理费</td>
            <td>${formatNumber(params.construction_mgmt_pct * 100, 1)}% 设备费</td>
            <td>-</td>
            <td>${formatNumber(capex.construction_mgmt)}</td>
        </tr>
        <tr>
            <td>调试费用</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.commissioning)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">安装施工小计</td>
            <td>${formatNumber(capex.installation_subtotal)}</td>
        </tr>
        
        <!-- 六、建设期保险 -->
        <tr class="section-header"><td colspan="4">六、建设期保险</td></tr>
        <tr>
            <td>建设工程一切险(CAR)</td>
            <td>${formatNumber(params.car_insurance_pct * 100, 2)}% 设备费</td>
            <td>-</td>
            <td>${formatNumber(capex.car_insurance)}</td>
        </tr>
        <tr>
            <td>安装工程一切险(EAR)</td>
            <td>${formatNumber(params.ear_insurance_pct * 100, 2)}% 设备费</td>
            <td>-</td>
            <td>${formatNumber(capex.ear_insurance)}</td>
        </tr>
        <tr>
            <td>货物运输保险</td>
            <td>${formatNumber(params.cargo_insurance_pct * 100, 2)}% 设备费</td>
            <td>-</td>
            <td>${formatNumber(capex.cargo_insurance)}</td>
        </tr>
        <tr>
            <td>第三方责任险</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.liability_insurance)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">建设期保险小计</td>
            <td>${formatNumber(capex.insurance_subtotal)}</td>
        </tr>
        
        <!-- 七、开发与业主费用 -->
        <tr class="section-header"><td colspan="4">七、开发与业主费用</td></tr>
        <tr>
            <td>SPV公司收购成本</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.spv_acquisition)}</td>
        </tr>
        <tr>
            <td>许可与规划费</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.permit)}</td>
        </tr>
        <tr>
            <td>环境咨询费</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.environmental)}</td>
        </tr>
        <tr>
            <td>法律咨询费</td>
            <td>-</td>
            <td>1 项</td>
            <td>${formatNumber(capex.legal)}</td>
        </tr>
        <tr>
            <td>工程设计费</td>
            <td>${formatNumber(params.engineering_pct * 100, 1)}% 设备费</td>
            <td>-</td>
            <td>${formatNumber(capex.engineering)}</td>
        </tr>
        <tr>
            <td>项目管理费</td>
            <td>${formatNumber(params.project_mgmt_pct * 100, 1)}% 总投资</td>
            <td>-</td>
            <td>${formatNumber(capex.project_mgmt)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">开发费用小计</td>
            <td>${formatNumber(capex.dev_subtotal)}</td>
        </tr>
        
        <!-- 八、不可预见费 -->
        <tr class="section-header"><td colspan="4">八、其他</td></tr>
        <tr>
            <td>不可预见费</td>
            <td>${formatNumber(params.contingency_pct * 100, 1)}%</td>
            <td>-</td>
            <td>${formatNumber(capex.contingency)}</td>
        </tr>
        <tr class="subtotal">
            <td colspan="3">建设期利息（${formatNumber(params.construction_period, 1)}年，资金占用${formatNumber(params.construction_fund_usage * 100, 0)}%）</td>
            <td>${formatNumber(capex.construction_interest)}</td>
        </tr>
    `;
    
    document.getElementById('total_capex').textContent = formatNumber(capex.total) + ' 万EUR';
}

// ==================== OPEX计算 ====================

/**
 * 计算年度OPEX
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @returns {Object[]} 年度OPEX数组
 */
function calculateOpex(params, capex) {
    const opexData = [];
    
    // 如果选择逐年计提，计算每年应计提的拆除准备金
    const annualDecommissioning = params.decommissioning_total
        ? (params.decommissioning_total / params.operation_years / 10000)
        : 0;
    
    for (let year = 1; year <= params.operation_years; year++) {
        // 计算通胀因子：考虑通货膨胀对OPEX的影响
        // 通胀影响 = (1 + 通胀率)^(年数-1)
        const inflationFactor = Math.pow(1 + params.inflation_rate, year - 1);
        
        // 各项OPEX = 基础值 × (1 + 各自增长率)^(年数-1) × 通胀因子
        const yearData = {
            year: year,
            technical: params.opex_technical * params.power_mw * 1000 * 
                      Math.pow(1 + params.opex_technical_esc, year - 1) * inflationFactor / 10000,
            insurance: capex.total * params.opex_insurance * 
                      Math.pow(1 + params.opex_insurance_esc, year - 1) * inflationFactor,
            grid: params.opex_grid * params.power_mw * 
                  Math.pow(1 + params.opex_grid_esc, year - 1) * inflationFactor / 10000,
            land: params.opex_land * Math.pow(1 + params.opex_land_esc, year - 1) * inflationFactor / 10000,
            commercial: params.opex_commercial * params.power_mw * 
                       Math.pow(1 + params.opex_commercial_esc, year - 1) * inflationFactor / 10000,
            other: params.opex_other * params.power_mw * 
                  Math.pow(1 + params.opex_other_esc, year - 1) * inflationFactor / 10000,
            decommissioning: annualDecommissioning * inflationFactor  // 拆除准备金也受通胀影响
        };
        yearData.total = yearData.technical + yearData.insurance + yearData.grid + 
                         yearData.land + yearData.commercial + yearData.other + yearData.decommissioning;
        opexData.push(yearData);
    }
    
    return opexData;
}

/**
 * 更新OPEX年度表格
 * @param {Object[]} opexData OPEX数据
 */
function updateOpexTable(opexData) {
    const tbody = document.getElementById('opex_yearly_body');
    tbody.innerHTML = '';
    
    opexData.forEach(data => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>第${data.year}年</td>
            <td>${formatNumber(data.technical)}</td>
            <td>${formatNumber(data.insurance)}</td>
            <td>${formatNumber(data.grid)}</td>
            <td>${formatNumber(data.land)}</td>
            <td>${formatNumber(data.commercial)}</td>
            <td>${formatNumber(data.other)}</td>
            <td>${formatNumber(data.decommissioning || 0)}</td>
            <td><strong>${formatNumber(data.total)}</strong></td>
        `;
        tbody.appendChild(row);
    });
}

// ==================== 收入计算 ====================

/**
 * 计算年度收入
 * @param {Object} params 参数
 * @returns {Object[]} 年度收入数组
 */
function calculateRevenue(params) {
    const revenueData = [];
    const spotPrices = getSpotPrices(params.operation_years);
    
    for (let year = 1; year <= params.operation_years; year++) {
        // 计算当年可用容量
        const capacityFactor = params.initial_capacity_pct / 100 * 
                               Math.pow(1 - params.degradation_rate, year - 1);
        
        // Tolling收入
        let tollingRevenue = 0;
        if (year <= params.tolling_years) {
            const tollingPrice = params.tolling_price * Math.pow(1 + params.tolling_escalation, year - 1);
            tollingRevenue = tollingPrice * params.power_mw * 1000 * params.tolling_ratio / 10000;
        }
        
        // 现货收入
        const spotRatio = year <= params.tolling_years ? (1 - params.tolling_ratio) : 1;
        const spotRevenue = spotPrices[year - 1] * params.power_mw * spotRatio * capacityFactor / 10000;
        
        revenueData.push({
            year: year,
            capacityFactor: capacityFactor * 100,
            tollingRevenue: tollingRevenue,
            spotRevenue: spotRevenue,
            totalRevenue: tollingRevenue + spotRevenue
        });
    }
    
    return revenueData;
}

/**
 * 更新收入年度表格
 * @param {Object[]} revenueData 收入数据
 */
function updateRevenueTable(revenueData) {
    const tbody = document.getElementById('revenue_yearly_body');
    tbody.innerHTML = '';
    
    revenueData.forEach(data => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>第${data.year}年</td>
            <td>${formatNumber(data.capacityFactor, 1)}%</td>
            <td>${formatNumber(data.tollingRevenue)}</td>
            <td>${formatNumber(data.spotRevenue)}</td>
            <td><strong>${formatNumber(data.totalRevenue)}</strong></td>
        `;
        tbody.appendChild(row);
    });
}

// ==================== 折旧计算 ====================

/**
 * 计算年度折旧
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @returns {Object[]} 年度折旧数组
 */
function calculateDepreciation(params, capex) {
    const depreciationData = [];
    
    // 无形资产（开发费用+土地）
    const intangibleAssets = (capex.dev_cost || 0) + (capex.land || 0);
    
    // 固定资产原值 = 动态总投资 - 无形资产（建设期利息已资本化）
    const fixedAssetOriginal = capex.dynamic_total - intangibleAssets;
    const depreciableAmount = fixedAssetOriginal * (1 - params.salvage_rate);
    
    for (let year = 1; year <= params.operation_years; year++) {
        let depreciation = 0;
        
        if (year <= params.depreciation_years) {
            switch (params.depreciation_method) {
                case 'straight_line':
                    depreciation = depreciableAmount / params.depreciation_years;
                    break;
                case 'double_declining':
                    const rate = 2 / params.depreciation_years;
                    let bookValue = fixedAssetOriginal;
                    for (let i = 1; i < year; i++) {
                        bookValue -= bookValue * rate;
                    }
                    depreciation = Math.min(bookValue * rate, bookValue - fixedAssetOriginal * params.salvage_rate);
                    break;
                case 'sum_of_years':
                    const sumYears = params.depreciation_years * (params.depreciation_years + 1) / 2;
                    depreciation = depreciableAmount * (params.depreciation_years - year + 1) / sumYears;
                    break;
            }
        }
        
        // 无形资产摊销
        const amortization = year <= params.amortization_years ? 
                            intangibleAssets / params.amortization_years : 0;
        
        depreciationData.push({
            year: year,
            depreciation: depreciation,
            amortization: amortization,
            total: depreciation + amortization
        });
    }
    
    return depreciationData;
}

// ==================== 贷款计算 ====================

/**
 * 计算贷款还款计划
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @returns {Object[]} 贷款还款计划
 */
function calculateLoan(params, capex) {
    const loanData = [];
    const loanAmount = capex.dynamic_total * (1 - params.equity_ratio);
    let balance = loanAmount;
    const repaymentYears = Math.max(params.loan_years - params.grace_period, 0);
    
    for (let year = 1; year <= params.loan_years; year++) {
        const interest = balance * params.loan_rate;
        let principal = 0;
        
        if (year > params.grace_period && repaymentYears > 0) {
            if (params.repayment_method === 'equal_principal') {
                // 等额本金
                principal = loanAmount / repaymentYears;
            } else {
                // 等额本息
                const n = repaymentYears;
                const r = params.loan_rate;
                const payment = r === 0
                    ? loanAmount / n
                    : loanAmount * r * Math.pow(1 + r, n) / (Math.pow(1 + r, n) - 1);
                principal = payment - interest;
            }
        }
        
        const payment = interest + principal;
        const endBalance = balance - principal;
        
        loanData.push({
            year: year,
            beginBalance: balance,
            interest: interest,
            principal: principal,
            payment: payment,
            endBalance: Math.max(0, endBalance)
        });
        
        balance = Math.max(0, endBalance);
    }
    
    // 补充剩余年份(贷款已还清)
    for (let year = params.loan_years + 1; year <= params.operation_years; year++) {
        loanData.push({
            year: year,
            beginBalance: 0,
            interest: 0,
            principal: 0,
            payment: 0,
            endBalance: 0
        });
    }
    
    return loanData;
}

/**
 * 更新贷款还款表格
 * @param {Object[]} loanData 贷款数据
 */
function updateLoanTable(loanData) {
    const tbody = document.getElementById('loan_body');
    tbody.innerHTML = '';
    
    loanData.forEach(data => {
        if (data.beginBalance > 0 || data.year === 1) {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>第${data.year}年</td>
                <td>${formatNumber(data.beginBalance)}</td>
                <td>${formatNumber(data.interest)}</td>
                <td>${formatNumber(data.principal)}</td>
                <td>${formatNumber(data.payment)}</td>
                <td>${formatNumber(data.endBalance)}</td>
            `;
            tbody.appendChild(row);
        }
    });
}

// ==================== 利润表计算 ====================

/**
 * 计算利润表
 * @param {Object} params 参数
 * @param {Object[]} revenueData 收入数据
 * @param {Object[]} opexData OPEX数据
 * @param {Object[]} depreciationData 折旧数据
 * @param {Object[]} loanData 贷款数据
 * @returns {Object[]} 利润表数据
 */
function calculateIncomeStatement(params, revenueData, opexData, depreciationData, loanData) {
    const incomeData = [];
    
    // 计算综合税率
    const effectiveTaxRate = params.corporate_tax_rate * (1 + params.solidarity_tax_rate) + 
                            params.trade_tax_rate + params.other_tax_rate;
    
    for (let year = 0; year < params.operation_years; year++) {
        const revenue = revenueData[year].totalRevenue;
        const opex = opexData[year].total;
        const depreciation = depreciationData[year].total;
        const interest = loanData[year].interest;
        
        const grossProfit = revenue - opex;
        const ebitda = grossProfit;
        const ebit = ebitda - depreciation;
        const ebt = ebit - interest;
        const tax = Math.max(0, ebt * effectiveTaxRate);
        const netProfit = ebt - tax;
        
        incomeData.push({
            year: year + 1,
            revenue: revenue,
            opex: opex,
            grossProfit: grossProfit,
            ebitda: ebitda,
            depreciation: depreciation,
            ebit: ebit,
            interest: interest,
            ebt: ebt,
            tax: tax,
            netProfit: netProfit
        });
    }
    
    return incomeData;
}

/**
 * 更新利润表显示
 * @param {Object[]} incomeData 利润表数据
 */
function updateIncomeTable(incomeData) {
    const table = document.getElementById('income_table');
    const thead = table.querySelector('thead tr');
    const tbody = document.getElementById('income_body');
    
    // 生成表头
    thead.innerHTML = '<th>项目（万EUR）</th>';
    incomeData.forEach(data => {
        thead.innerHTML += `<th>第${data.year}年</th>`;
    });
    
    // 生成表体
    const rows = [
        { label: '营业收入', key: 'revenue', class: '' },
        { label: '减：运营成本', key: 'opex', class: '' },
        { label: '毛利润', key: 'grossProfit', class: 'subtotal' },
        { label: 'EBITDA', key: 'ebitda', class: '' },
        { label: '减：折旧摊销', key: 'depreciation', class: '' },
        { label: 'EBIT(息税前利润)', key: 'ebit', class: 'subtotal' },
        { label: '减：财务费用(利息)', key: 'interest', class: '' },
        { label: 'EBT(税前利润)', key: 'ebt', class: 'subtotal' },
        { label: '减：所得税', key: 'tax', class: '' },
        { label: '净利润', key: 'netProfit', class: 'total' }
    ];
    
    tbody.innerHTML = '';
    rows.forEach(row => {
        const tr = document.createElement('tr');
        tr.className = row.class;
        tr.innerHTML = `<td>${row.label}</td>`;
        incomeData.forEach(data => {
            tr.innerHTML += `<td>${formatNumber(data[row.key])}</td>`;
        });
        tbody.appendChild(tr);
    });
}

// ==================== 现金流量表计算 ====================

/**
 * 计算现金流量表
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @param {Object[]} incomeData 利润表数据
 * @param {Object[]} depreciationData 折旧数据
 * @param {Object[]} loanData 贷款数据
 * @returns {Object[]} 现金流量表数据
 */
function calculateCashFlow(params, capex, incomeData, depreciationData, loanData) {
    const cashFlowData = [];
    
    // 第0年 - 建设期
    const equity = capex.dynamic_total * params.equity_ratio;
    const loan = capex.dynamic_total * (1 - params.equity_ratio);
    
    cashFlowData.push({
        year: 0,
        // 经营活动
        netProfit: 0,
        depreciation: 0,
        workingCapital: 0,
        operatingCashFlow: 0,
        // 投资活动
        capex: -capex.dynamic_total,
        investingCashFlow: -capex.dynamic_total,
        // 筹资活动
        equityInflow: equity,
        loanInflow: loan,
        loanRepayment: 0,
        financingCashFlow: equity + loan,
        // 现金净增加
        netCashFlow: 0,
        // 全投资现金流
        projectCashFlow: -capex.dynamic_total,
        // 资本金现金流
        equityCashFlow: -equity
    });
    
    // 运营期
    for (let i = 0; i < params.operation_years; i++) {
        const income = incomeData[i];
        const depreciation = depreciationData[i].total;
        const loanInfo = loanData[i];
        
        const operatingCashFlow = income.netProfit + depreciation;
        const investingCashFlow = 0;
        const financingCashFlow = -loanInfo.principal;
        const netCashFlow = operatingCashFlow + investingCashFlow + financingCashFlow;
        
        // 全投资现金流 = 经营现金流 (不考虑融资)
        const projectCashFlow = income.ebitda - income.tax;
        
        // 资本金现金流 = 净现金流
        const equityCashFlow = netCashFlow;
        
        cashFlowData.push({
            year: i + 1,
            netProfit: income.netProfit,
            depreciation: depreciation,
            workingCapital: 0,
            operatingCashFlow: operatingCashFlow,
            capex: 0,
            investingCashFlow: investingCashFlow,
            equityInflow: 0,
            loanInflow: 0,
            loanRepayment: -loanInfo.principal,
            financingCashFlow: financingCashFlow,
            netCashFlow: netCashFlow,
            projectCashFlow: projectCashFlow,
            equityCashFlow: equityCashFlow
        });
    }
    
    // 添加残值回收（基于固定资产原值，即动态投资减去无形资产）
    const intangibleAssets = (capex.dev_cost || 0) + (capex.land || 0);
    const fixedAssetOriginal = capex.dynamic_total - intangibleAssets;
    const salvageValue = fixedAssetOriginal * params.salvage_rate;
    
    if (cashFlowData.length > params.operation_years) {
        const lastYearIndex = params.operation_years;
        cashFlowData[lastYearIndex].investingCashFlow += salvageValue;
        cashFlowData[lastYearIndex].netCashFlow += salvageValue;
        cashFlowData[lastYearIndex].projectCashFlow += salvageValue;
        cashFlowData[lastYearIndex].equityCashFlow += salvageValue;
    }
    
    return cashFlowData;
}

/**
 * 更新现金流量表显示
 * @param {Object[]} cashFlowData 现金流量表数据
 */
function updateCashFlowTable(cashFlowData) {
    const table = document.getElementById('cashflow_table');
    const thead = table.querySelector('thead tr');
    const tbody = document.getElementById('cashflow_body');
    
    // 生成表头
    thead.innerHTML = '<th>项目（万EUR）</th>';
    cashFlowData.forEach(data => {
        thead.innerHTML += `<th>${data.year === 0 ? '建设期' : '第' + data.year + '年'}</th>`;
    });
    
    // 生成表体
    const rows = [
        { label: '一、经营活动现金流', key: null, class: 'section-header' },
        { label: '  净利润', key: 'netProfit', class: '' },
        { label: '  加：折旧摊销', key: 'depreciation', class: '' },
        { label: '  经营活动现金流小计', key: 'operatingCashFlow', class: 'subtotal' },
        { label: '二、投资活动现金流', key: null, class: 'section-header' },
        { label: '  资本性支出/残值回收', key: 'investingCashFlow', class: '' },
        { label: '三、筹资活动现金流', key: null, class: 'section-header' },
        { label: '  股权资金流入', key: 'equityInflow', class: '' },
        { label: '  贷款流入', key: 'loanInflow', class: '' },
        { label: '  贷款偿还', key: 'loanRepayment', class: '' },
        { label: '  筹资活动现金流小计', key: 'financingCashFlow', class: 'subtotal' },
        { label: '现金净增加额', key: 'netCashFlow', class: 'total' },
        { label: '全投资现金流(税后)', key: 'projectCashFlow', class: 'total' },
        { label: '资本金现金流(税后)', key: 'equityCashFlow', class: 'total' }
    ];
    
    tbody.innerHTML = '';
    rows.forEach(row => {
        const tr = document.createElement('tr');
        tr.className = row.class;
        if (row.key === null) {
            tr.innerHTML = `<td colspan="${cashFlowData.length + 1}">${row.label}</td>`;
        } else {
            tr.innerHTML = `<td>${row.label}</td>`;
            cashFlowData.forEach(data => {
                tr.innerHTML += `<td>${formatNumber(data[row.key])}</td>`;
            });
        }
        tbody.appendChild(tr);
    });
}

// ==================== 资产负债表计算 ====================

/**
 * 计算资产负债表
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @param {Object[]} incomeData 利润表数据
 * @param {Object[]} depreciationData 折旧数据
 * @param {Object[]} loanData 贷款数据
 * @param {Object[]} cashFlowData 现金流数据
 * @returns {Object[]} 资产负债表数据
 */
function calculateBalanceSheet(params, capex, incomeData, depreciationData, loanData, cashFlowData) {
    const balanceData = [];
    
    let accumulatedDepreciation = 0;
    let retainedEarnings = 0;
    let cash = 0;
    const equity = capex.dynamic_total * params.equity_ratio;
    const loanAmount = capex.dynamic_total * (1 - params.equity_ratio);
    
    // 固定资产原值 = 总投资 - 无形资产（开发费用+土地）
    // 建设期利息已资本化计入固定资产
    const intangibleAssetsOriginal = capex.dev_cost + capex.land;
    const fixedAssetOriginal = capex.dynamic_total - intangibleAssetsOriginal;
    
    // 初始资产负债表 (建设完成时点)
    balanceData.push({
        year: 0,
        // 资产
        cash: 0,
        fixedAssetOriginal: fixedAssetOriginal,
        accumulatedDepreciation: 0,
        fixedAssetNet: fixedAssetOriginal,
        intangibleAssets: intangibleAssetsOriginal,
        totalAssets: capex.dynamic_total,
        // 负债
        longTermLoan: loanAmount,
        totalLiabilities: loanAmount,
        // 所有者权益
        paidInCapital: equity,
        retainedEarnings: 0,
        totalEquity: equity,
        // 验证
        totalLiabilitiesAndEquity: loanAmount + equity
    });
    
    for (let i = 0; i < params.operation_years; i++) {
        accumulatedDepreciation += depreciationData[i].depreciation;
        retainedEarnings += incomeData[i].netProfit;
        
        // 累计现金流（从运营期第1年开始）
        cash += cashFlowData[i + 1].netCashFlow;
        
        let fixedAssetNet = fixedAssetOriginal - accumulatedDepreciation;
        const accumulatedAmortization = depreciationData.slice(0, i + 1).reduce((a, b) => a + b.amortization, 0);
        let intangibleAssets = Math.max(0, intangibleAssetsOriginal - accumulatedAmortization);
        const longTermLoan = i < loanData.length ? loanData[i].endBalance : 0;
        
        // 最后一年：固定资产处置（残值已回收计入现金），资产清零
        if (i === params.operation_years - 1) {
            // 残值已通过现金流回收，固定资产净值设为0
            fixedAssetNet = 0;
            intangibleAssets = 0;
        }
        
        // 总资产 = 货币资金 + 固定资产净值 + 无形资产
        const totalAssets = cash + fixedAssetNet + intangibleAssets;
        const totalEquity = equity + retainedEarnings;
        const totalLiabilitiesAndEquity = longTermLoan + totalEquity;
        
        balanceData.push({
            year: i + 1,
            cash: cash,
            fixedAssetOriginal: fixedAssetOriginal,
            accumulatedDepreciation: accumulatedDepreciation,
            fixedAssetNet: fixedAssetNet,
            intangibleAssets: intangibleAssets,
            totalAssets: totalAssets,
            longTermLoan: longTermLoan,
            totalLiabilities: longTermLoan,
            paidInCapital: equity,
            retainedEarnings: retainedEarnings,
            totalEquity: totalEquity,
            totalLiabilitiesAndEquity: totalLiabilitiesAndEquity
        });
    }
    
    return balanceData;
}

/**
 * 更新资产负债表显示
 * @param {Object[]} balanceData 资产负债表数据
 */
function updateBalanceTable(balanceData) {
    const table = document.getElementById('balance_table');
    const thead = table.querySelector('thead tr');
    const tbody = document.getElementById('balance_body');
    
    // 生成表头
    thead.innerHTML = '<th>项目（万EUR）</th>';
    balanceData.forEach(data => {
        thead.innerHTML += `<th>${data.year === 0 ? '期初' : '第' + data.year + '年末'}</th>`;
    });
    
    // 生成表体
    const rows = [
        { label: '资产', key: null, class: 'section-header' },
        { label: '  货币资金', key: 'cash', class: '' },
        { label: '  固定资产原值', key: 'fixedAssetOriginal', class: '' },
        { label: '  减：累计折旧', key: 'accumulatedDepreciation', class: '' },
        { label: '  固定资产净值', key: 'fixedAssetNet', class: 'subtotal' },
        { label: '  无形资产', key: 'intangibleAssets', class: '' },
        { label: '资产合计', key: 'totalAssets', class: 'total' },
        { label: '负债', key: null, class: 'section-header' },
        { label: '  长期借款', key: 'longTermLoan', class: '' },
        { label: '负债合计', key: 'totalLiabilities', class: 'total' },
        { label: '所有者权益', key: null, class: 'section-header' },
        { label: '  实收资本', key: 'paidInCapital', class: '' },
        { label: '  未分配利润', key: 'retainedEarnings', class: '' },
        { label: '所有者权益合计', key: 'totalEquity', class: 'total' },
        { label: '负债和所有者权益合计', key: 'totalLiabilitiesAndEquity', class: 'total' }
    ];
    
    tbody.innerHTML = '';
    rows.forEach(row => {
        const tr = document.createElement('tr');
        tr.className = row.class;
        if (row.key === null) {
            tr.innerHTML = `<td colspan="${balanceData.length + 1}">${row.label}</td>`;
        } else {
            tr.innerHTML = `<td>${row.label}</td>`;
            balanceData.forEach(data => {
                tr.innerHTML += `<td>${formatNumber(data[row.key])}</td>`;
            });
        }
        tbody.appendChild(tr);
    });
}

// ==================== 财务指标计算 ====================

/**
 * 计算IRR (内部收益率)
 * @param {number[]} cashFlows 现金流数组
 * @returns {number} IRR百分比
 */
function calculateIRR(cashFlows) {
    const maxIterations = 1000;
    const tolerance = 0.00001;
    let rate = 0.1;
    
    for (let i = 0; i < maxIterations; i++) {
        let npv = 0;
        let dnpv = 0;
        
        for (let j = 0; j < cashFlows.length; j++) {
            npv += cashFlows[j] / Math.pow(1 + rate, j);
            dnpv -= j * cashFlows[j] / Math.pow(1 + rate, j + 1);
        }
        
        const newRate = rate - npv / dnpv;
        
        if (Math.abs(newRate - rate) < tolerance) {
            return newRate * 100;
        }
        
        rate = newRate;
        
        // 防止发散
        if (rate < -0.99 || rate > 10 || isNaN(rate)) {
            return NaN;
        }
    }
    
    return rate * 100;
}

/**
 * 计算NPV (净现值)
 * @param {number[]} cashFlows 现金流数组
 * @param {number} discountRate 折现率
 * @returns {number} NPV
 */
function calculateNPV(cashFlows, discountRate) {
    let npv = 0;
    for (let i = 0; i < cashFlows.length; i++) {
        npv += cashFlows[i] / Math.pow(1 + discountRate, i);
    }
    return npv;
}

/**
 * 计算静态回收期
 * @param {number[]} cashFlows 现金流数组
 * @returns {number} 回收期（年）
 */
function calculateStaticPayback(cashFlows) {
    let cumulative = 0;
    for (let i = 0; i < cashFlows.length; i++) {
        cumulative += cashFlows[i];
        if (cumulative >= 0) {
            // 插值计算精确年份
            const prevCumulative = cumulative - cashFlows[i];
            return i - 1 + Math.abs(prevCumulative) / cashFlows[i];
        }
    }
    return cashFlows.length; // 未回收
}

/**
 * 计算动态回收期
 * @param {number[]} cashFlows 现金流数组
 * @param {number} discountRate 折现率
 * @returns {number} 动态回收期（年）
 */
function calculateDynamicPayback(cashFlows, discountRate) {
    let cumulative = 0;
    for (let i = 0; i < cashFlows.length; i++) {
        const discountedCF = cashFlows[i] / Math.pow(1 + discountRate, i);
        cumulative += discountedCF;
        if (cumulative >= 0) {
            const prevCumulative = cumulative - discountedCF;
            return i - 1 + Math.abs(prevCumulative) / discountedCF;
        }
    }
    return cashFlows.length;
}

/**
 * 计算所有财务指标
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @param {Object[]} revenueData 收入数据
 * @param {Object[]} incomeData 利润表数据
 * @param {Object[]} cashFlowData 现金流数据
 * @param {Object[]} balanceData 资产负债表数据
 * @param {Object[]} loanData 贷款数据
 * @returns {Object} 财务指标
 */
function calculateIndicators(params, capex, revenueData, incomeData, cashFlowData, balanceData, loanData) {
    // 提取全投资现金流
    const projectCashFlows = cashFlowData.map(d => d.projectCashFlow);
    
    // 提取资本金现金流
    const equityCashFlows = cashFlowData.map(d => d.equityCashFlow);
    
    // 计算各项指标
    const totalRevenue = revenueData.reduce((a, b) => a + b.totalRevenue, 0);
    const avgRevenue = totalRevenue / params.operation_years;
    const first3Revenue = revenueData.slice(0, 3).reduce((a, b) => a + b.totalRevenue, 0);
    
    const totalProfit = incomeData.reduce((a, b) => a + b.ebt, 0);
    const avgProfit = totalProfit / params.operation_years;
    const first3Profit = incomeData.slice(0, 3).reduce((a, b) => a + b.ebt, 0);
    
    const totalNetProfit = incomeData.reduce((a, b) => a + b.netProfit, 0);
    const avgNetProfit = totalNetProfit / params.operation_years;
    const first3NetProfit = incomeData.slice(0, 3).reduce((a, b) => a + b.netProfit, 0);
    
    // IRR计算
    const projectIRR = calculateIRR(projectCashFlows);
    const equityIRR = calculateIRR(equityCashFlows);
    
    // 回收期计算
    const staticPayback = calculateStaticPayback(projectCashFlows);
    const equityStaticPayback = calculateStaticPayback(equityCashFlows);
    const dynamicPayback = calculateDynamicPayback(projectCashFlows, 0.08);
    const equityDynamicPayback = calculateDynamicPayback(equityCashFlows, 0.08);
    
    // ROE (第三年净资产收益率)
    const year3NetProfit = incomeData.length >= 3 ? incomeData[2].netProfit : 0;
    const year3Equity = balanceData.length >= 4 ? balanceData[3].totalEquity : capex.dynamic_total * params.equity_ratio;
    const roe3 = year3Equity > 0 ? (year3NetProfit / year3Equity) * 100 : 0;
    
    // ROI (总投资收益率)
    const roi = (avgNetProfit / capex.dynamic_total) * 100;
    
    // EBITDA回报率
    const avgEBITDA = incomeData.reduce((a, b) => a + b.ebitda, 0) / params.operation_years;
    const ebitdaReturn = (avgEBITDA / capex.dynamic_total) * 100;
    
    // DSCR (债务偿付覆盖率) - 使用运营期平均值
    let totalDebtService = 0;
    let totalEBITDA = 0;
    let debtServiceYears = 0;
    
    for (let i = 0; i < params.loan_years && i < incomeData.length; i++) {
        if (loanData[i].payment > 0) {
            totalDebtService += loanData[i].payment;
            totalEBITDA += incomeData[i].ebitda;
            debtServiceYears++;
        }
    }
    
    const dscr = debtServiceYears > 0 && totalDebtService > 0 ? 
                 totalEBITDA / totalDebtService : 0;

    // LCOE (平准化度电成本，EUR/MWh)
    /** @type {number} */
    const totalOpex = incomeData.reduce((a, b) => a + b.opex, 0);
    /** @type {number} */
    const totalCostEur = (capex.dynamic_total + totalOpex) * 10000;
    /** @type {number} */
    const totalEnergyMwh = incomeData.reduce((sum, _, index) => {
        const year = index + 1;
        const capacityFactor = params.initial_capacity_pct / 100 * Math.pow(1 - params.degradation_rate, year - 1);
        const annualEnergy = params.capacity_mwh * capacityFactor * params.annual_cycles *
            params.charge_efficiency * params.discharge_efficiency;
        return sum + annualEnergy;
    }, 0);
    const lcoe = totalEnergyMwh > 0 ? totalCostEur / totalEnergyMwh : 0;
    
    return {
        static_investment: capex.total,
        dynamic_investment: capex.dynamic_total,
        total_revenue: totalRevenue,
        avg_revenue: avgRevenue,
        first3_revenue: first3Revenue,
        total_profit: totalProfit,
        avg_profit: avgProfit,
        first3_profit: first3Profit,
        total_net_profit: totalNetProfit,
        avg_net_profit: avgNetProfit,
        first3_net_profit: first3NetProfit,
        project_irr: projectIRR,
        equity_irr: equityIRR,
        static_payback: staticPayback,
        equity_payback: equityStaticPayback,
        dynamic_payback: dynamicPayback,
        equity_dynamic_payback: equityDynamicPayback,
        roe_year3: roe3,
        roi: roi,
        ebitda_return: ebitdaReturn,
        dscr: dscr,
        lcoe: lcoe
    };
}

/**
 * 更新财务指标显示
 * @param {Object} indicators 财务指标
 */
function updateIndicatorsDisplay(indicators) {
    // 金额类指标（使用千分位格式）
    document.getElementById('ind_static_investment').textContent = formatNumber(indicators.static_investment);
    document.getElementById('ind_dynamic_investment').textContent = formatNumber(indicators.dynamic_investment);
    document.getElementById('ind_total_revenue').textContent = formatNumber(indicators.total_revenue);
    document.getElementById('ind_avg_revenue').textContent = formatNumber(indicators.avg_revenue);
    document.getElementById('ind_first3_revenue').textContent = formatNumber(indicators.first3_revenue);
    document.getElementById('ind_total_profit').textContent = formatNumber(indicators.total_profit);
    document.getElementById('ind_avg_profit').textContent = formatNumber(indicators.avg_profit);
    document.getElementById('ind_first3_profit').textContent = formatNumber(indicators.first3_profit);
    document.getElementById('ind_total_net_profit').textContent = formatNumber(indicators.total_net_profit);
    document.getElementById('ind_avg_net_profit').textContent = formatNumber(indicators.avg_net_profit);
    document.getElementById('ind_first3_net_profit').textContent = formatNumber(indicators.first3_net_profit);
    
    // 百分比类指标
    document.getElementById('ind_project_irr').textContent = formatNumber(indicators.project_irr);
    document.getElementById('ind_equity_irr').textContent = formatNumber(indicators.equity_irr);
    document.getElementById('ind_roe_year3').textContent = formatNumber(indicators.roe_year3);
    document.getElementById('ind_roi').textContent = formatNumber(indicators.roi);
    document.getElementById('ind_ebitda_return').textContent = formatNumber(indicators.ebitda_return);
    document.getElementById('ind_lcoe').textContent = formatNumber(indicators.lcoe, 2);
    
    // 回收期（年）
    document.getElementById('ind_static_payback').textContent = formatNumber(indicators.static_payback, 1);
    document.getElementById('ind_equity_payback').textContent = formatNumber(indicators.equity_payback, 1);
    document.getElementById('ind_dynamic_payback').textContent = formatNumber(indicators.dynamic_payback, 1);
    document.getElementById('ind_equity_dynamic_payback').textContent = formatNumber(indicators.equity_dynamic_payback, 1);
    
    // DSCR（倍数）
    document.getElementById('ind_dscr').textContent = formatNumber(indicators.dscr);
}

// ==================== 图表绑定 ====================

/**
 * 更新收入图表
 * @param {Object[]} revenueData 收入数据
 */
function updateRevenueChart(revenueData) {
    const ctx = document.getElementById('revenueChart').getContext('2d');
    
    if (charts.revenue) {
        charts.revenue.destroy();
    }
    
    charts.revenue = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: revenueData.map(d => `第${d.year}年`),
            datasets: [
                {
                    label: 'Tolling收入',
                    data: revenueData.map(d => d.tollingRevenue),
                    backgroundColor: 'rgba(0, 212, 170, 0.8)',
                    borderColor: 'rgba(0, 212, 170, 1)',
                    borderWidth: 1
                },
                {
                    label: '现货收入',
                    data: revenueData.map(d => d.spotRevenue),
                    backgroundColor: 'rgba(255, 107, 53, 0.8)',
                    borderColor: 'rgba(255, 107, 53, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: '年度收入构成 (万EUR)',
                    color: '#f1f5f9',
                    font: { size: 16 }
                },
                legend: {
                    labels: { color: '#94a3b8' }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    ticks: { color: '#94a3b8' },
                    grid: { color: 'rgba(255, 255, 255, 0.1)' }
                },
                y: {
                    stacked: true,
                    ticks: { color: '#94a3b8' },
                    grid: { color: 'rgba(255, 255, 255, 0.1)' }
                }
            }
        }
    });
}

// ==================== 敏感性分析 ====================

/**
 * 运行敏感性分析
 */
function runSensitivityAnalysis() {
    const var1 = document.getElementById('sens_var1').value;
    const var2 = document.getElementById('sens_var2').value;
    const target = document.getElementById('sens_target').value;
    const minChange = parseFloat(document.getElementById('sens_min').value) / 100;
    const maxChange = parseFloat(document.getElementById('sens_max').value) / 100;
    const step = parseFloat(document.getElementById('sens_step').value) / 100;
    
    const baseParams = getParameters();
    const changes = [];
    for (let c = minChange; c <= maxChange + 0.001; c += step) {
        changes.push(Math.round(c * 100) / 100);
    }
    
    if (!var2) {
        // 单变量敏感性分析
        const results = runSingleVariableSensitivity(baseParams, var1, target, changes);
        displaySingleVariableSensitivity(results, var1, target, changes);
        document.getElementById('sensitivity_matrix_card').style.display = 'none';
    } else {
        // 双变量敏感性分析
        const results = runDoubleVariableSensitivity(baseParams, var1, var2, target, changes);
        displayDoubleVariableSensitivity(results, var1, var2, target, changes);
    }
}

/**
 * 单变量敏感性分析
 */
function runSingleVariableSensitivity(baseParams, variable, target, changes) {
    const results = [];
    
    changes.forEach(change => {
        const params = JSON.parse(JSON.stringify(baseParams));
        applyVariableChange(params, variable, change);
        const value = calculateTargetIndicator(params, target);
        results.push({
            change: change * 100,
            value: value
        });
    });
    
    return results;
}

/**
 * 双变量敏感性分析
 */
function runDoubleVariableSensitivity(baseParams, var1, var2, target, changes) {
    const results = [];
    
    changes.forEach(change1 => {
        const row = [];
        changes.forEach(change2 => {
            const params = JSON.parse(JSON.stringify(baseParams));
            applyVariableChange(params, var1, change1);
            applyVariableChange(params, var2, change2);
            const value = calculateTargetIndicator(params, target);
            row.push(value);
        });
        results.push(row);
    });
    
    return results;
}

/**
 * 应用变量变化
 */
function applyVariableChange(params, variable, change) {
    switch (variable) {
        case 'capex':
            params.battery_unit_price *= (1 + change);
            params.pcs_unit_price *= (1 + change);
            params.mv_transformer_price *= (1 + change);
            params.hv_transformer_price *= (1 + change);
            break;
        case 'tolling_price':
            params.tolling_price *= (1 + change);
            break;
        case 'spot_price':
            // 这里需要特殊处理现货价格
            params.spot_price_change = change;
            break;
        case 'opex':
            params.opex_technical *= (1 + change);
            params.opex_insurance *= (1 + change);
            params.opex_grid *= (1 + change);
            params.opex_land *= (1 + change);
            params.opex_commercial *= (1 + change);
            params.opex_other *= (1 + change);
            break;
        case 'loan_rate':
            params.loan_rate *= (1 + change);
            break;
        case 'degradation':
            params.degradation_rate *= (1 + change);
            break;
    }
}

/**
 * 计算目标指标
 */
function calculateTargetIndicator(params, target) {
    // 重新计算所有数据
    const capex = calculateCapex(params);
    const opexData = calculateOpex(params, capex);
    
    // 处理现货价格变化
    const originalSpotPrices = getSpotPrices(params.operation_years);
    const spotPriceChange = params.spot_price_change || 0;
    const adjustedSpotPrices = originalSpotPrices.map(p => p * (1 + spotPriceChange));
    
    // 临时修改DOM值来计算收入
    const revenueData = calculateRevenueWithPrices(params, adjustedSpotPrices);
    
    const depreciationData = calculateDepreciation(params, capex);
    const loanData = calculateLoan(params, capex);
    const incomeData = calculateIncomeStatement(params, revenueData, opexData, depreciationData, loanData);
    const cashFlowData = calculateCashFlow(params, capex, incomeData, depreciationData, loanData);
    const balanceData = calculateBalanceSheet(params, capex, incomeData, depreciationData, loanData, cashFlowData);
    
    const indicators = calculateIndicators(params, capex, revenueData, incomeData, cashFlowData, balanceData, loanData);
    
    switch (target) {
        case 'project_irr': return indicators.project_irr;
        case 'equity_irr': return indicators.equity_irr;
        case 'npv': 
            const projectCashFlows = cashFlowData.map(d => d.projectCashFlow);
            return calculateNPV(projectCashFlows, 0.08);
        case 'payback': return indicators.static_payback;
        default: return 0;
    }
}

/**
 * 使用指定价格计算收入
 */
function calculateRevenueWithPrices(params, spotPrices) {
    const revenueData = [];
    
    for (let year = 1; year <= params.operation_years; year++) {
        const capacityFactor = params.initial_capacity_pct / 100 * 
                               Math.pow(1 - params.degradation_rate, year - 1);
        
        let tollingRevenue = 0;
        if (year <= params.tolling_years) {
            const tollingPrice = params.tolling_price * Math.pow(1 + params.tolling_escalation, year - 1);
            tollingRevenue = tollingPrice * params.power_mw * 1000 * params.tolling_ratio / 10000;
        }
        
        const spotRatio = year <= params.tolling_years ? (1 - params.tolling_ratio) : 1;
        const spotRevenue = spotPrices[year - 1] * params.power_mw * spotRatio * capacityFactor / 10000;
        
        revenueData.push({
            year: year,
            capacityFactor: capacityFactor * 100,
            tollingRevenue: tollingRevenue,
            spotRevenue: spotRevenue,
            totalRevenue: tollingRevenue + spotRevenue
        });
    }
    
    return revenueData;
}

/**
 * 显示单变量敏感性分析结果
 */
function displaySingleVariableSensitivity(results, variable, target, changes) {
    const varNames = {
        capex: '投资成本',
        tolling_price: 'Tolling价格',
        spot_price: '现货价格',
        opex: '运营成本',
        loan_rate: '贷款利率',
        degradation: '电池衰减率'
    };
    
    const targetNames = {
        project_irr: '全投资IRR',
        equity_irr: '资本金IRR',
        npv: '项目NPV',
        payback: '投资回收期'
    };
    
    // 创建结果表格
    const resultDiv = document.getElementById('sensitivity_result');
    let tableHTML = `
        <table class="sensitivity-matrix">
            <thead>
                <tr>
                    <th>${varNames[variable]}变化</th>
                    <th>${targetNames[target]}</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    const baseValue = results.find(r => r.change === 0)?.value || 0;
    
    results.forEach(r => {
        const changeClass = r.value > baseValue ? 'cell-positive' : (r.value < baseValue ? 'cell-negative' : 'cell-neutral');
        tableHTML += `
            <tr>
                <td>${r.change >= 0 ? '+' : ''}${formatNumber(r.change, 0)}%</td>
                <td class="${changeClass}">${isNaN(r.value) ? '-' : formatNumber(r.value)}${target.includes('irr') || target === 'payback' ? '' : ' 万EUR'}</td>
            </tr>
        `;
    });
    
    tableHTML += '</tbody></table>';
    resultDiv.innerHTML = tableHTML;
    
    // 更新图表
    updateSensitivityChart(results, varNames[variable], targetNames[target]);
}

/**
 * 显示双变量敏感性分析结果
 */
function displayDoubleVariableSensitivity(results, var1, var2, target, changes) {
    const varNames = {
        capex: '投资成本',
        tolling_price: 'Tolling价格',
        spot_price: '现货价格',
        opex: '运营成本',
        loan_rate: '贷款利率',
        degradation: '电池衰减率'
    };
    
    document.getElementById('sensitivity_matrix_card').style.display = 'block';
    
    const table = document.getElementById('sensitivity_matrix');
    let html = '<thead><tr><th>' + varNames[var1] + ' \\ ' + varNames[var2] + '</th>';
    
    changes.forEach(c => {
        html += `<th>${(c * 100) >= 0 ? '+' : ''}${formatNumber(c * 100, 0)}%</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // 找到基准值（0%变化）
    const baseIndex = changes.findIndex(c => Math.abs(c) < 0.001);
    const baseValue = baseIndex >= 0 ? results[baseIndex][baseIndex] : results[Math.floor(results.length/2)][Math.floor(results.length/2)];
    
    results.forEach((row, i) => {
        html += `<tr><th>${(changes[i] * 100) >= 0 ? '+' : ''}${formatNumber(changes[i] * 100, 0)}%</th>`;
        row.forEach(value => {
            const changeClass = value > baseValue ? 'cell-positive' : (value < baseValue ? 'cell-negative' : 'cell-neutral');
            html += `<td class="${changeClass}">${isNaN(value) ? '-' : formatNumber(value)}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody>';
    table.innerHTML = html;
    
    // 清空单变量结果区
    document.getElementById('sensitivity_result').innerHTML = '<p class="hint-text">双变量分析结果见下方矩阵</p>';
    
    // 清空图表
    if (charts.sensitivity) {
        charts.sensitivity.destroy();
        charts.sensitivity = null;
    }
}

/**
 * 更新敏感性分析图表
 */
function updateSensitivityChart(results, varName, targetName) {
    const ctx = document.getElementById('sensitivityChart').getContext('2d');
    
    if (charts.sensitivity) {
        charts.sensitivity.destroy();
    }
    
    charts.sensitivity = new Chart(ctx, {
        type: 'line',
        data: {
            labels: results.map(r => `${r.change >= 0 ? '+' : ''}${formatNumber(r.change, 0)}%`),
            datasets: [{
                label: targetName,
                data: results.map(r => r.value),
                borderColor: 'rgba(0, 212, 170, 1)',
                backgroundColor: 'rgba(0, 212, 170, 0.2)',
                fill: true,
                tension: 0.3,
                pointRadius: 5,
                pointHoverRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: `${varName}变化对${targetName}的影响`,
                    color: '#f1f5f9',
                    font: { size: 16 }
                },
                legend: {
                    labels: { color: '#94a3b8' }
                }
            },
            scales: {
                x: {
                    title: {
                        display: true,
                        text: varName + '变化幅度',
                        color: '#94a3b8'
                    },
                    ticks: { color: '#94a3b8' },
                    grid: { color: 'rgba(255, 255, 255, 0.1)' }
                },
                y: {
                    title: {
                        display: true,
                        text: targetName,
                        color: '#94a3b8'
                    },
                    ticks: { color: '#94a3b8' },
                    grid: { color: 'rgba(255, 255, 255, 0.1)' }
                }
            }
        }
    });
}

// ==================== 报表切换 ====================

/**
 * 切换财务报表显示
 * @param {string} type 报表类型
 */
function switchStatement(type) {
    // 更新按钮状态
    document.querySelectorAll('.statement-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    event.target.classList.add('active');
    
    // 切换内容
    document.querySelectorAll('.statement-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(type + '_statement').classList.add('active');
}

// ==================== 主计算函数 ====================

/**
 * 执行所有计算
 */
function calculateAll() {
    try {
        // 获取参数
        const params = getParameters();
        
        // 初始化现货价格表（如果年限变化）
        initSpotPriceTable();
        fillSpotPrices();
        
        // 计算CAPEX
        const capex = calculateCapex(params);
        updateCapexTable(capex, params);
        
        // 计算OPEX
        const opexData = calculateOpex(params, capex);
        updateOpexTable(opexData);
        
        // 计算收入
        const revenueData = calculateRevenue(params);
        updateRevenueTable(revenueData);
        updateRevenueChart(revenueData);
        
        // 计算折旧
        const depreciationData = calculateDepreciation(params, capex);
        
        // 计算贷款
        const loanData = calculateLoan(params, capex);
        updateLoanTable(loanData);
        
        // 计算利润表
        const incomeData = calculateIncomeStatement(params, revenueData, opexData, depreciationData, loanData);
        updateIncomeTable(incomeData);
        
        // 计算现金流量表
        const cashFlowData = calculateCashFlow(params, capex, incomeData, depreciationData, loanData);
        updateCashFlowTable(cashFlowData);
        
        // 计算资产负债表
        const balanceData = calculateBalanceSheet(params, capex, incomeData, depreciationData, loanData, cashFlowData);
        updateBalanceTable(balanceData);
        
        // 计算财务指标
        const indicators = calculateIndicators(params, capex, revenueData, incomeData, cashFlowData, balanceData, loanData);
        updateIndicatorsDisplay(indicators);
        
        // 保存结果
        calculationResults = {
            params,
            capex,
            opexData,
            revenueData,
            depreciationData,
            loanData,
            incomeData,
            cashFlowData,
            balanceData,
            indicators
        };
        
        // 更新融资报告
        updateBankReport(params, capex, revenueData, opexData, incomeData, cashFlowData, balanceData, loanData, indicators);
        
        console.log('计算完成', calculationResults);
        
    } catch (error) {
        console.error('计算错误:', error);
        alert('计算过程中发生错误，请检查输入参数');
    }
}

// ==================== 融资报告生成 ====================

/** @type {Object} 融资报告图表实例 */
let bankCharts = {
    revenueSource: null,
    revenueYearly: null,
    dscr: null,
    cashCoverage: null,
    profit: null,
    cumulativeCF: null,
    sensitivity: null,
    leverage: null,
    roe: null
};

/**
 * 更新融资报告界面
 * @param {Object} params 参数
 * @param {Object} capex CAPEX数据
 * @param {Object[]} revenueData 收入数据
 * @param {Object[]} opexData OPEX数据
 * @param {Object[]} incomeData 利润表数据
 * @param {Object[]} cashFlowData 现金流数据
 * @param {Object[]} balanceData 资产负债表数据
 * @param {Object[]} loanData 贷款数据
 * @param {Object} indicators 财务指标
 */
function updateBankReport(params, capex, revenueData, opexData, incomeData, cashFlowData, balanceData, loanData, indicators) {
    // 更新报告日期
    document.getElementById('report_date').textContent = new Date().toLocaleDateString('zh-CN');
    
    // 更新项目概要
    updateProjectSummary(params, capex, indicators);
    
    // 更新收入分析图表
    updateRevenueAnalysisCharts(params, revenueData);
    
    // 更新偿债能力分析
    updateDebtServiceAnalysis(params, incomeData, loanData);
    
    // 更新盈利能力分析
    updateProfitabilityAnalysis(params, incomeData, revenueData, capex);
    
    // 更新风险分析
    updateRiskAnalysis(params, capex, indicators);
    
    // 更新财务结构分析
    updateFinancialStructure(params, balanceData, incomeData);
    
    // 更新结论与建议
    updateConclusion(params, indicators, incomeData, loanData);
    
    // 更新附录表格
    updateAppendixTables(params, incomeData, cashFlowData, balanceData);
}

/**
 * 更新项目概要
 */
function updateProjectSummary(params, capex, indicators) {
    // 项目基本信息
    document.getElementById('rpt_capacity').textContent = `${formatInt(params.power_mw)} MW / ${formatInt(params.capacity_mwh)} MWh`;
    document.getElementById('rpt_duration').textContent = `${formatNumber(params.capacity_mwh / params.power_mw, 1)} 小时`;
    document.getElementById('rpt_operation_years').textContent = `${params.operation_years} 年`;
    
    // 投资与融资结构
    const equity = capex.dynamic_total * params.equity_ratio;
    const loan = capex.dynamic_total * (1 - params.equity_ratio);
    
    document.getElementById('rpt_total_investment').textContent = `${formatNumber(capex.dynamic_total)} 万EUR`;
    document.getElementById('rpt_equity').textContent = `${formatNumber(equity)} 万EUR`;
    document.getElementById('rpt_loan').textContent = `${formatNumber(loan)} 万EUR`;
    document.getElementById('rpt_equity_ratio').textContent = `${formatNumber(params.equity_ratio * 100, 0)}%`;
    document.getElementById('rpt_loan_term').textContent = `${params.loan_years} 年`;
    document.getElementById('rpt_loan_rate').textContent = `${formatNumber(params.loan_rate * 100)}%`;
    
    // 核心投资指标
    document.getElementById('rpt_project_irr').textContent = `${formatNumber(indicators.project_irr)}%`;
    document.getElementById('rpt_equity_irr').textContent = `${formatNumber(indicators.equity_irr)}%`;
    document.getElementById('rpt_payback').textContent = `${formatNumber(indicators.static_payback, 1)} 年`;
    document.getElementById('rpt_dynamic_payback').textContent = `${formatNumber(indicators.dynamic_payback, 1)} 年`;
    document.getElementById('rpt_dscr').textContent = `${formatNumber(indicators.dscr)}x`;
    document.getElementById('rpt_lcoe').textContent = `${formatNumber(indicators.lcoe, 2)} EUR/MWh`;
    
    // 计算最低DSCR
    const minDSCR = calculateMinDSCR(calculationResults.incomeData, calculationResults.loanData, params);
    document.getElementById('rpt_min_dscr').textContent = `${formatNumber(minDSCR)}x`;
}

/**
 * 计算最低DSCR
 */
function calculateMinDSCR(incomeData, loanData, params) {
    let minDSCR = Infinity;
    for (let i = 0; i < Math.min(params.loan_years, incomeData.length); i++) {
        if (loanData[i].payment > 0) {
            const dscr = incomeData[i].ebitda / loanData[i].payment;
            if (dscr < minDSCR) {
                minDSCR = dscr;
            }
        }
    }
    return minDSCR === Infinity ? 0 : minDSCR;
}

/**
 * 更新收入分析图表
 */
function updateRevenueAnalysisCharts(params, revenueData) {
    // 收入来源构成饼图
    const totalTolling = revenueData.reduce((a, b) => a + b.tollingRevenue, 0);
    const totalSpot = revenueData.reduce((a, b) => a + b.spotRevenue, 0);
    
    const ctx1 = document.getElementById('bankRevenueSourceChart').getContext('2d');
    if (bankCharts.revenueSource) bankCharts.revenueSource.destroy();
    
    bankCharts.revenueSource = new Chart(ctx1, {
        type: 'doughnut',
        data: {
            labels: ['Tolling收入', '现货交易收入'],
            datasets: [{
                data: [totalTolling, totalSpot],
                backgroundColor: ['rgba(0, 212, 170, 0.8)', 'rgba(255, 107, 53, 0.8)'],
                borderColor: ['rgba(0, 212, 170, 1)', 'rgba(255, 107, 53, 1)'],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8' }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = formatNumber(context.raw / total * 100, 1);
                            return `${context.label}: ${formatInt(context.raw)} 万EUR (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
    
    // 年度收入趋势图
    const ctx2 = document.getElementById('bankRevenueYearlyChart').getContext('2d');
    if (bankCharts.revenueYearly) bankCharts.revenueYearly.destroy();
    
    bankCharts.revenueYearly = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: revenueData.map(d => `Y${d.year}`),
            datasets: [
                {
                    label: 'Tolling',
                    data: revenueData.map(d => d.tollingRevenue),
                    backgroundColor: 'rgba(0, 212, 170, 0.8)'
                },
                {
                    label: '现货',
                    data: revenueData.map(d => d.spotRevenue),
                    backgroundColor: 'rgba(255, 107, 53, 0.8)'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8' }
                }
            },
            scales: {
                x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } }
            }
        }
    });
    
    // 收入稳定性评估
    const tollingCoverage = formatNumber(params.tolling_years / params.operation_years * 100, 0);
    const tollingRatio = formatNumber(totalTolling / (totalTolling + totalSpot) * 100, 0);
    
    document.getElementById('tolling_coverage_bar').style.width = tollingCoverage + '%';
    document.getElementById('rpt_tolling_coverage').textContent = `${tollingCoverage}% (${params.tolling_years}年)`;
    
    document.getElementById('tolling_ratio_bar').style.width = tollingRatio + '%';
    document.getElementById('rpt_tolling_ratio').textContent = `${tollingRatio}%`;
    
    // 收入波动系数
    const revenues = revenueData.map(d => d.totalRevenue);
    const avgRevenue = revenues.reduce((a, b) => a + b, 0) / revenues.length;
    const variance = revenues.reduce((a, b) => a + Math.pow(b - avgRevenue, 2), 0) / revenues.length;
    const volatility = formatNumber(Math.sqrt(variance) / avgRevenue * 100, 1);
    document.getElementById('rpt_revenue_volatility').textContent = `${volatility}%`;
}

/**
 * 更新偿债能力分析
 */
function updateDebtServiceAnalysis(params, incomeData, loanData) {
    // 年度DSCR趋势图
    const dscrData = [];
    const labels = [];
    
    for (let i = 0; i < Math.min(params.loan_years, incomeData.length); i++) {
        if (loanData[i].payment > 0) {
            labels.push(`Y${i + 1}`);
            dscrData.push(incomeData[i].ebitda / loanData[i].payment);
        }
    }
    
    const ctx1 = document.getElementById('bankDSCRChart').getContext('2d');
    if (bankCharts.dscr) bankCharts.dscr.destroy();
    
    bankCharts.dscr = new Chart(ctx1, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'DSCR',
                data: dscrData,
                borderColor: 'rgba(0, 212, 170, 1)',
                backgroundColor: 'rgba(0, 212, 170, 0.2)',
                fill: true,
                tension: 0.3,
                pointRadius: 4
            }, {
                label: '最低要求 (1.2x)',
                data: Array(labels.length).fill(1.2),
                borderColor: 'rgba(255, 193, 7, 1)',
                borderDash: [5, 5],
                pointRadius: 0,
                fill: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' }, min: 0 }
            }
        }
    });
    
    // 现金流覆盖分析图
    const ctx2 = document.getElementById('bankCashCoverageChart').getContext('2d');
    if (bankCharts.cashCoverage) bankCharts.cashCoverage.destroy();
    
    const ebitdaData = [];
    const debtServiceData = [];
    
    for (let i = 0; i < Math.min(params.loan_years, incomeData.length); i++) {
        ebitdaData.push(incomeData[i].ebitda);
        debtServiceData.push(loanData[i].payment);
    }
    
    bankCharts.cashCoverage = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'EBITDA',
                    data: ebitdaData,
                    backgroundColor: 'rgba(0, 212, 170, 0.8)'
                },
                {
                    label: '偿债支出',
                    data: debtServiceData,
                    backgroundColor: 'rgba(220, 53, 69, 0.8)'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } }
            }
        }
    });
    
    // 更新贷款偿还计划表
    const tbody = document.getElementById('bank_loan_body');
    tbody.innerHTML = '';
    
    for (let i = 0; i < Math.min(params.loan_years, incomeData.length); i++) {
        const dscr = loanData[i].payment > 0 ? incomeData[i].ebitda / loanData[i].payment : 0;
        let dscrClass = 'dscr-good';
        if (dscr < 1.2) dscrClass = 'dscr-warning';
        if (dscr < 1.0) dscrClass = 'dscr-danger';
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>第${i + 1}年</td>
            <td>${formatNumber(loanData[i].beginBalance)}</td>
            <td>${formatNumber(loanData[i].interest)}</td>
            <td>${formatNumber(loanData[i].principal)}</td>
            <td>${formatNumber(loanData[i].payment)}</td>
            <td>${formatNumber(loanData[i].endBalance)}</td>
            <td>${formatNumber(incomeData[i].ebitda)}</td>
            <td class="${dscrClass}">${formatNumber(dscr)}x</td>
        `;
        tbody.appendChild(row);
    }
}

/**
 * 更新盈利能力分析
 */
function updateProfitabilityAnalysis(params, incomeData, revenueData, capex) {
    // 年度利润趋势图
    const ctx1 = document.getElementById('bankProfitChart').getContext('2d');
    if (bankCharts.profit) bankCharts.profit.destroy();
    
    bankCharts.profit = new Chart(ctx1, {
        type: 'line',
        data: {
            labels: incomeData.map(d => `Y${d.year}`),
            datasets: [
                {
                    label: 'EBITDA',
                    data: incomeData.map(d => d.ebitda),
                    borderColor: 'rgba(0, 212, 170, 1)',
                    backgroundColor: 'transparent',
                    tension: 0.3
                },
                {
                    label: '净利润',
                    data: incomeData.map(d => d.netProfit),
                    borderColor: 'rgba(255, 107, 53, 1)',
                    backgroundColor: 'transparent',
                    tension: 0.3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } }
            }
        }
    });
    
    // 累计现金流图
    const ctx2 = document.getElementById('bankCumulativeCFChart').getContext('2d');
    if (bankCharts.cumulativeCF) bankCharts.cumulativeCF.destroy();
    
    const cumulativeCF = [];
    let cumulative = -capex.dynamic_total;
    cumulativeCF.push(cumulative);
    
    for (let i = 0; i < incomeData.length; i++) {
        cumulative += incomeData[i].netProfit + calculationResults.depreciationData[i].total - 
                      (calculationResults.loanData[i] ? calculationResults.loanData[i].principal : 0);
        cumulativeCF.push(cumulative);
    }
    
    bankCharts.cumulativeCF = new Chart(ctx2, {
        type: 'line',
        data: {
            labels: ['Y0', ...incomeData.map(d => `Y${d.year}`)],
            datasets: [{
                label: '累计现金流',
                data: cumulativeCF,
                borderColor: 'rgba(0, 212, 170, 1)',
                backgroundColor: 'rgba(0, 212, 170, 0.2)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } }
            }
        }
    });
    
    // 盈利指标
    const avgEbitda = incomeData.reduce((a, b) => a + b.ebitda, 0) / incomeData.length;
    const avgNetProfit = incomeData.reduce((a, b) => a + b.netProfit, 0) / incomeData.length;
    const totalRevenue = revenueData.reduce((a, b) => a + b.totalRevenue, 0);
    const totalEbitda = incomeData.reduce((a, b) => a + b.ebitda, 0);
    const totalNetProfit = incomeData.reduce((a, b) => a + b.netProfit, 0);
    
    document.getElementById('rpt_avg_ebitda').textContent = formatInt(avgEbitda);
    document.getElementById('rpt_avg_net_profit').textContent = formatInt(avgNetProfit);
    document.getElementById('rpt_ebitda_margin').textContent = formatNumber(totalEbitda / totalRevenue * 100, 1);
    document.getElementById('rpt_net_margin').textContent = formatNumber(totalNetProfit / totalRevenue * 100, 1);
}

/**
 * 更新风险分析
 */
function updateRiskAnalysis(params, capex, indicators) {
    // 敏感性分析图表
    const ctx = document.getElementById('bankSensitivityChart').getContext('2d');
    if (bankCharts.sensitivity) bankCharts.sensitivity.destroy();
    
    // 计算各变量敏感性
    const variables = ['CAPEX', 'Tolling价格', '现货价格', 'OPEX', '贷款利率'];
    const baseIRR = indicators.project_irr;
    const sensData = {
        minus20: [],
        minus10: [],
        base: [],
        plus10: [],
        plus20: []
    };
    
    const varKeys = ['capex', 'tolling_price', 'spot_price', 'opex', 'loan_rate'];
    
    varKeys.forEach(varKey => {
        [-0.2, -0.1, 0, 0.1, 0.2].forEach((change, idx) => {
            const baseParams = getParameters();
            applyVariableChange(baseParams, varKey, change);
            const irr = calculateTargetIndicator(baseParams, 'project_irr');
            const keys = ['minus20', 'minus10', 'base', 'plus10', 'plus20'];
            sensData[keys[idx]].push(irr);
        });
    });
    
    bankCharts.sensitivity = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: variables,
            datasets: [
                {
                    label: '-20%',
                    data: sensData.minus20,
                    backgroundColor: 'rgba(220, 53, 69, 0.7)'
                },
                {
                    label: '-10%',
                    data: sensData.minus10,
                    backgroundColor: 'rgba(255, 193, 7, 0.7)'
                },
                {
                    label: '基准',
                    data: sensData.base,
                    backgroundColor: 'rgba(0, 212, 170, 0.7)'
                },
                {
                    label: '+10%',
                    data: sensData.plus10,
                    backgroundColor: 'rgba(255, 193, 7, 0.7)'
                },
                {
                    label: '+20%',
                    data: sensData.plus20,
                    backgroundColor: 'rgba(220, 53, 69, 0.7)'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: '关键变量对项目IRR的影响 (%)',
                    color: '#f1f5f9'
                },
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } }
            }
        }
    });
    
    // 压力测试
    const stressScenarios = [
        { name: '基准情景', condition: '当前假设', capexChange: 0, priceChange: 0 },
        { name: '保守情景', condition: 'CAPEX+10%, 价格-10%', capexChange: 0.1, priceChange: -0.1 },
        { name: '悲观情景', condition: 'CAPEX+20%, 价格-20%', capexChange: 0.2, priceChange: -0.2 },
        { name: '极端情景', condition: 'CAPEX+30%, 价格-30%', capexChange: 0.3, priceChange: -0.3 }
    ];
    
    const tbody = document.getElementById('stress_test_body');
    tbody.innerHTML = '';
    
    stressScenarios.forEach(scenario => {
        const testParams = getParameters();
        applyVariableChange(testParams, 'capex', scenario.capexChange);
        applyVariableChange(testParams, 'tolling_price', scenario.priceChange);
        applyVariableChange(testParams, 'spot_price', scenario.priceChange);
        
        const irr = calculateTargetIndicator(testParams, 'project_irr');
        
        // 计算最低DSCR
        const testCapex = calculateCapex(testParams);
        const testOpex = calculateOpex(testParams, testCapex);
        const testRevenue = calculateRevenueWithPrices(testParams, getSpotPrices(testParams.operation_years).map(p => p * (1 + scenario.priceChange)));
        const testDepreciation = calculateDepreciation(testParams, testCapex);
        const testLoan = calculateLoan(testParams, testCapex);
        const testIncome = calculateIncomeStatement(testParams, testRevenue, testOpex, testDepreciation, testLoan);
        const minDSCR = calculateMinDSCR(testIncome, testLoan, testParams);
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${scenario.name}</td>
            <td>${scenario.condition}</td>
            <td>${isNaN(irr) ? '-' : formatNumber(irr) + '%'}</td>
            <td class="${minDSCR >= 1.2 ? 'dscr-good' : (minDSCR >= 1.0 ? 'dscr-warning' : 'dscr-danger')}">${formatNumber(minDSCR)}x</td>
        `;
        tbody.appendChild(row);
    });
}

/**
 * 更新财务结构分析
 */
function updateFinancialStructure(params, balanceData, incomeData) {
    // 资产负债率趋势
    const ctx1 = document.getElementById('bankLeverageChart').getContext('2d');
    if (bankCharts.leverage) bankCharts.leverage.destroy();
    
    const leverageData = balanceData.slice(1).map(d => 
        d.totalAssets > 0 ? (d.totalLiabilities / d.totalAssets * 100) : 0
    );
    
    bankCharts.leverage = new Chart(ctx1, {
        type: 'line',
        data: {
            labels: balanceData.slice(1).map(d => `Y${d.year}`),
            datasets: [{
                label: '资产负债率',
                data: leverageData,
                borderColor: 'rgba(0, 212, 170, 1)',
                backgroundColor: 'rgba(0, 212, 170, 0.2)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8', callback: v => v + '%' }, grid: { color: 'rgba(255,255,255,0.1)' }, min: 0 }
            }
        }
    });
    
    // ROE趋势
    const ctx2 = document.getElementById('bankROEChart').getContext('2d');
    if (bankCharts.roe) bankCharts.roe.destroy();
    
    const roeData = [];
    for (let i = 0; i < incomeData.length; i++) {
        const equity = balanceData[i + 1] ? balanceData[i + 1].totalEquity : 0;
        const roe = equity > 0 ? (incomeData[i].netProfit / equity * 100) : 0;
        roeData.push(roe);
    }
    
    bankCharts.roe = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: incomeData.map(d => `Y${d.year}`),
            datasets: [{
                label: 'ROE',
                data: roeData,
                backgroundColor: 'rgba(255, 107, 53, 0.8)'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.1)' } },
                y: { ticks: { color: '#94a3b8', callback: v => v + '%' }, grid: { color: 'rgba(255,255,255,0.1)' } }
            }
        }
    });
}

/**
 * 更新结论与建议
 */
function updateConclusion(params, indicators, incomeData, loanData) {
    // 投资评级
    let rating = 'A';
    let ratingText = '推荐投资';
    
    if (indicators.project_irr >= 12 && indicators.dscr >= 1.5) {
        rating = 'AAA';
        ratingText = '强烈推荐';
    } else if (indicators.project_irr >= 10 && indicators.dscr >= 1.3) {
        rating = 'AA';
        ratingText = '推荐';
    } else if (indicators.project_irr >= 8 && indicators.dscr >= 1.2) {
        rating = 'A';
        ratingText = '可接受';
    } else if (indicators.project_irr >= 6 && indicators.dscr >= 1.1) {
        rating = 'BBB';
        ratingText = '需谨慎';
    } else {
        rating = 'BB';
        ratingText = '风险较高';
    }
    
    document.getElementById('rpt_rating').textContent = `${rating} (${ratingText})`;
    
    // 项目优势
    const strengthsList = document.getElementById('project_strengths');
    const strengths = [];
    
    if (params.tolling_ratio >= 0.5) {
        strengths.push(`Tolling合同覆盖${formatNumber(params.tolling_ratio * 100, 0)}%容量，收入稳定性高`);
    }
    if (indicators.project_irr >= 8) {
        strengths.push(`项目IRR达到${formatNumber(indicators.project_irr)}%，投资回报良好`);
    }
    if (indicators.dscr >= 1.3) {
        strengths.push(`平均DSCR为${formatNumber(indicators.dscr)}x，偿债能力充足`);
    }
    if (params.equity_ratio >= 0.25) {
        strengths.push(`资本金比例${formatNumber(params.equity_ratio * 100, 0)}%，杠杆率合理`);
    }
    strengths.push('德国能源转型政策支持储能发展，市场前景良好');
    strengths.push('采用成熟的磷酸铁锂电池技术，技术风险可控');
    
    strengthsList.innerHTML = strengths.map(s => `<li>${s}</li>`).join('');
    
    // 关注事项
    const concernsList = document.getElementById('project_concerns');
    const concerns = [];
    
    if (params.degradation_rate > 0.02) {
        concerns.push('电池年衰减率较高，需关注长期容量损失');
    }
    if (indicators.dscr < 1.3) {
        concerns.push('DSCR相对较低，建议设置现金储备账户');
    }
    if (params.tolling_years < params.loan_years) {
        concerns.push(`Tolling合同期限(${params.tolling_years}年)短于贷款期限(${params.loan_years}年)`);
    }
    concerns.push('关注电力市场价格波动对现货收入的影响');
    concerns.push('建议建立设备运维储备金机制');
    
    concernsList.innerHTML = concerns.map(c => `<li>${c}</li>`).join('');
    
    // 融资建议
    const recommendation = `
        基于财务分析结果，本项目具有良好的投资价值和偿债能力。建议贷款机构以${formatNumber(params.loan_rate * 100)}%的年利率，
        ${params.loan_years}年期限提供项目融资，贷款金额${formatNumber(calculationResults.capex.dynamic_total * (1 - params.equity_ratio))}万欧元。
        建议设置DSCR不低于1.2x的财务约束条款，并要求项目公司建立偿债准备金账户。
        ${indicators.dscr >= 1.3 ? '项目偿债能力充足，融资风险可控。' : '建议要求提供额外增信措施。'}
    `;
    document.getElementById('financing_recommendation').textContent = recommendation;
}

/**
 * 更新附录表格
 */
function updateAppendixTables(params, incomeData, cashFlowData, balanceData) {
    // 现金流预测表
    const cfHead = document.getElementById('bank_cashflow_head');
    const cfBody = document.getElementById('bank_cashflow_body');
    
    cfHead.innerHTML = '<tr><th>项目 (万EUR)</th>' + 
        cashFlowData.map(d => `<th>${d.year === 0 ? '建设期' : 'Y' + d.year}</th>`).join('') + '</tr>';
    
    const cfRows = [
        { label: '经营活动现金流', key: 'operatingCashFlow' },
        { label: '投资活动现金流', key: 'investingCashFlow' },
        { label: '筹资活动现金流', key: 'financingCashFlow' },
        { label: '现金净增加额', key: 'netCashFlow' },
        { label: '全投资现金流', key: 'projectCashFlow' },
        { label: '资本金现金流', key: 'equityCashFlow' }
    ];
    
    cfBody.innerHTML = cfRows.map(row => 
        `<tr><td>${row.label}</td>${cashFlowData.map(d => `<td>${formatNumber(d[row.key])}</td>`).join('')}</tr>`
    ).join('');
    
    // 损益预测表
    const incHead = document.getElementById('bank_income_head');
    const incBody = document.getElementById('bank_income_body');
    
    incHead.innerHTML = '<tr><th>项目 (万EUR)</th>' + 
        incomeData.map(d => `<th>Y${d.year}</th>`).join('') + '</tr>';
    
    const incRows = [
        { label: '营业收入', key: 'revenue' },
        { label: '运营成本', key: 'opex' },
        { label: 'EBITDA', key: 'ebitda' },
        { label: '折旧摊销', key: 'depreciation' },
        { label: 'EBIT', key: 'ebit' },
        { label: '利息支出', key: 'interest' },
        { label: '税前利润', key: 'ebt' },
        { label: '所得税', key: 'tax' },
        { label: '净利润', key: 'netProfit' }
    ];
    
    incBody.innerHTML = incRows.map(row => 
        `<tr><td>${row.label}</td>${incomeData.map(d => `<td>${formatNumber(d[row.key])}</td>`).join('')}</tr>`
    ).join('');
    
    // 资产负债预测表
    const balHead = document.getElementById('bank_balance_head');
    const balBody = document.getElementById('bank_balance_body');
    
    balHead.innerHTML = '<tr><th>项目 (万EUR)</th>' + 
        balanceData.map(d => `<th>${d.year === 0 ? '期初' : 'Y' + d.year + '末'}</th>`).join('') + '</tr>';
    
    const balRows = [
        { label: '总资产', key: 'totalAssets' },
        { label: '固定资产净值', key: 'fixedAssetNet' },
        { label: '货币资金', key: 'cash' },
        { label: '总负债', key: 'totalLiabilities' },
        { label: '长期借款', key: 'longTermLoan' },
        { label: '所有者权益', key: 'totalEquity' }
    ];
    
    balBody.innerHTML = balRows.map(row => 
        `<tr><td>${row.label}</td>${balanceData.map(d => `<td>${formatNumber(d[row.key])}</td>`).join('')}</tr>`
    ).join('');
}

/**
 * 切换附录内容
 * @param {string} type 附录类型
 */
function switchAppendix(type) {
    document.querySelectorAll('.appendix-tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.appendix-content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById('appendix_' + type).classList.add('active');
}

// ==================== 模型导入/保存/导出功能 ====================

/**
 * 导入模型
 */
function importModel() {
    document.getElementById('modelFileInput').click();
}

/**
 * 处理模型文件导入
 * @param {Event} event 文件选择事件
 */
function handleModelImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const modelData = JSON.parse(e.target.result);
            
            // 验证模型数据
            if (!modelData.modelVersion || !modelData.parameters) {
                throw new Error('无效的模型文件格式');
            }
            
            // 恢复参数到界面
            restoreParameters(modelData.parameters);
            
            // 恢复现货价格
            if (modelData.spotPrices) {
                initSpotPriceTable();
                setTimeout(() => {
                    modelData.spotPrices.forEach((price, index) => {
                        const input = document.getElementById(`spot_price_${index + 1}`);
                        if (input) input.value = price;
                    });
                }, 100);
            }
            
            // 重新计算
            setTimeout(() => {
                calculateAll();
                alert('模型导入成功！');
            }, 200);
            
        } catch (error) {
            console.error('导入错误:', error);
            alert('模型导入失败: ' + error.message);
        }
    };
    reader.readAsText(file);
    
    // 重置文件输入
    event.target.value = '';
}

/**
 * 恢复参数到界面
 * @param {Object} params 参数对象
 */
function restoreParameters(params) {
    const paramMappings = {
        power_mw: 'power_mw',
        capacity_mwh: 'capacity_mwh',
        operation_years: 'operation_years',
        initial_capacity_pct: 'initial_capacity_pct',
        equity_ratio: { id: 'equity_ratio', transform: v => v * 100 },
        loan_years: 'loan_years',
        loan_rate: { id: 'loan_rate', transform: v => v * 100 },
        grace_period: 'grace_period',
        repayment_method: 'repayment_method',
        depreciation_years: 'depreciation_years',
        salvage_rate: { id: 'salvage_rate', transform: v => v * 100 },
        depreciation_method: 'depreciation_method',
        amortization_years: 'amortization_years',
        charge_efficiency: { id: 'charge_efficiency', transform: v => v * 100 },
        discharge_efficiency: { id: 'discharge_efficiency', transform: v => v * 100 },
        degradation_rate: { id: 'degradation_rate', transform: v => v * 100 },
        annual_cycles: 'annual_cycles',
        corporate_tax_rate: { id: 'corporate_tax_rate', transform: v => v * 100 },
        solidarity_tax_rate: { id: 'solidarity_tax_rate', transform: v => v * 100 },
        trade_tax_rate: { id: 'trade_tax_rate', transform: v => v * 100 },
        vat_rate: { id: 'vat_rate', transform: v => v * 100 },
        other_tax_rate: { id: 'other_tax_rate', transform: v => v * 100 },
        tolling_years: 'tolling_years',
        tolling_ratio: { id: 'tolling_ratio', transform: v => v * 100 },
        tolling_price: 'tolling_price',
        tolling_escalation: { id: 'tolling_escalation', transform: v => v * 100 },
        battery_cabinet_capacity: 'battery_cabinet_capacity',
        battery_cabinet_count: 'battery_cabinet_count',
        battery_unit_price: 'battery_unit_price',
        pcs_power: 'pcs_power',
        pcs_count: 'pcs_count',
        pcs_unit_price: 'pcs_unit_price',
        mv_transformer_capacity: 'mv_transformer_capacity',
        mv_transformer_count: 'mv_transformer_count',
        mv_transformer_price: 'mv_transformer_price',
        hv_transformer_capacity: 'hv_transformer_capacity',
        hv_transformer_count: 'hv_transformer_count',
        hv_transformer_price: 'hv_transformer_price',
        opex_technical: 'opex_technical',
        opex_technical_esc: { id: 'opex_technical_esc', transform: v => v * 100 },
        opex_insurance: { id: 'opex_insurance', transform: v => v * 100 },
        opex_insurance_esc: { id: 'opex_insurance_esc', transform: v => v * 100 },
        opex_grid: 'opex_grid',
        opex_grid_esc: { id: 'opex_grid_esc', transform: v => v * 100 },
        opex_land: 'opex_land',
        opex_land_esc: { id: 'opex_land_esc', transform: v => v * 100 },
        opex_commercial: 'opex_commercial',
        opex_commercial_esc: { id: 'opex_commercial_esc', transform: v => v * 100 },
        opex_other: 'opex_other',
        opex_other_esc: { id: 'opex_other_esc', transform: v => v * 100 }
    };
    
    Object.keys(paramMappings).forEach(key => {
        if (params[key] !== undefined) {
            const mapping = paramMappings[key];
            let elementId, value;
            
            if (typeof mapping === 'string') {
                elementId = mapping;
                value = params[key];
            } else {
                elementId = mapping.id;
                value = mapping.transform(params[key]);
            }
            
            const element = document.getElementById(elementId);
            if (element) {
                element.value = value;
            }
        }
    });
    
    // 更新衍生值
    updateDuration();
    updateRTE();
}

/**
 * 保存模型
 */
function saveModel() {
    try {
        const params = getParameters();
        const spotPrices = getSpotPrices(params.operation_years);
        
        const modelData = {
            modelVersion: '1.0',
            modelName: '德国独立储能电站财务模型',
            savedAt: new Date().toISOString(),
            parameters: params,
            spotPrices: spotPrices
        };
        
        const jsonStr = JSON.stringify(modelData, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = `储能电站模型_${params.power_mw}MW_${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        alert('模型保存成功！');
        
    } catch (error) {
        console.error('保存错误:', error);
        alert('模型保存失败: ' + error.message);
    }
}

/**
 * 导出报表
 */
function exportReport() {
    // 显示导出选项对话框
    const exportType = prompt(
        '请选择导出格式:\n' +
        '1 - 导出Excel格式 (CSV)\n' +
        '2 - 导出PDF格式 (打印)\n' +
        '3 - 导出完整HTML报告',
        '1'
    );
    
    if (!exportType) return;
    
    switch (exportType) {
        case '1':
            exportToCSV();
            break;
        case '2':
            exportToPDF();
            break;
        case '3':
            exportToHTML();
            break;
        default:
            alert('无效的选择');
    }
}

/**
 * 导出为CSV格式
 */
function exportToCSV() {
    try {
        const params = calculationResults.params;
        const capex = calculationResults.capex;
        const indicators = calculationResults.indicators;
        
        let csv = '\ufeff'; // BOM for Excel
        
        // 项目概要
        csv += '德国独立储能电站投资测算报告\n';
        csv += `导出时间,${new Date().toLocaleString()}\n\n`;
        
        csv += '一、项目概要\n';
        csv += `装机规模,${formatInt(params.power_mw)} MW / ${formatInt(params.capacity_mwh)} MWh\n`;
        csv += `运营期限,${params.operation_years} 年\n`;
        csv += `项目总投资,${formatNumber(capex.dynamic_total)} 万EUR\n`;
        csv += `资本金比例,${formatNumber(params.equity_ratio * 100, 0)}%\n`;
        csv += `贷款年限,${params.loan_years} 年\n`;
        csv += `贷款利率,${formatNumber(params.loan_rate * 100)}%\n\n`;
        
        // 核心指标
        csv += '二、核心财务指标\n';
        csv += `全投资IRR(税后),${formatNumber(indicators.project_irr)}%\n`;
        csv += `资本金IRR(税后),${formatNumber(indicators.equity_irr)}%\n`;
        csv += `静态投资回收期,${formatNumber(indicators.static_payback, 1)} 年\n`;
        csv += `动态投资回收期(8%),${formatNumber(indicators.dynamic_payback, 1)} 年\n`;
        csv += `平均DSCR,${formatNumber(indicators.dscr)}x\n`;
        csv += `LCOE(平准化度电成本),${formatNumber(indicators.lcoe, 2)} EUR/MWh\n`;
        csv += `累计净利润,${formatNumber(indicators.total_net_profit)} 万EUR\n\n`;
        
        // 年度明细
        csv += '三、年度财务明细\n';
        csv += '年份,收入(万EUR),OPEX(万EUR),EBITDA(万EUR),净利润(万EUR),DSCR\n';
        
        calculationResults.incomeData.forEach((income, i) => {
            const dscr = calculationResults.loanData[i].payment > 0 ? 
                        formatNumber(income.ebitda / calculationResults.loanData[i].payment) : '-';
            csv += `${income.year},${formatNumber(income.revenue)},${formatNumber(income.opex)},${formatNumber(income.ebitda)},${formatNumber(income.netProfit)},${dscr}\n`;
        });
        
        // 下载
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `储能电站财务报表_${params.power_mw}MW_${new Date().toISOString().slice(0, 10)}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        alert('CSV报表导出成功！');
        
    } catch (error) {
        console.error('导出错误:', error);
        alert('CSV导出失败: ' + error.message);
    }
}

/**
 * 导出为PDF格式（使用浏览器打印功能）
 */
function exportToPDF() {
    // 切换到融资报告页面
    document.querySelectorAll('.nav-link').forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('data-tab') === 'bankReport') {
            link.classList.add('active');
        }
    });
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.getElementById('bankReport').classList.add('active');
    
    // 延迟打印以确保页面渲染完成
    setTimeout(() => {
        window.print();
    }, 500);
}

/**
 * 导出为完整HTML报告
 */
function exportToHTML() {
    try {
        const params = calculationResults.params;
        const capex = calculationResults.capex;
        const indicators = calculationResults.indicators;
        
        let html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>德国独立储能电站融资可行性分析报告</title>
    <style>
        body { font-family: 'Microsoft YaHei', sans-serif; max-width: 1200px; margin: 0 auto; padding: 20px; background: #f5f5f5; }
        .header { text-align: center; padding: 30px; background: linear-gradient(135deg, #0a0f1c, #1a2332); color: white; border-radius: 10px; margin-bottom: 20px; }
        .header h1 { font-size: 24px; margin-bottom: 10px; }
        .header p { color: #94a3b8; }
        .section { background: white; padding: 20px; margin-bottom: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .section h2 { color: #00d4aa; border-bottom: 2px solid #00d4aa; padding-bottom: 10px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #eee; }
        th { background: #f8f9fa; color: #333; }
        td:last-child { text-align: right; font-family: 'Consolas', monospace; }
        .highlight { background: linear-gradient(135deg, rgba(0,212,170,0.1), transparent); }
        .key-value { font-weight: bold; color: #00d4aa; }
        .footer { text-align: center; color: #666; font-size: 12px; margin-top: 30px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🏦 项目融资可行性分析报告</h1>
        <p>Battery Energy Storage System (BESS) Project Financing Report</p>
        <p>生成日期: ${new Date().toLocaleDateString('zh-CN')}</p>
    </div>
    
    <div class="section">
        <h2>一、项目概要</h2>
        <table>
            <tr><td>项目名称</td><td>德国独立储能电站项目</td></tr>
            <tr><td>装机规模</td><td>${formatInt(params.power_mw)} MW / ${formatInt(params.capacity_mwh)} MWh</td></tr>
            <tr><td>储能时长</td><td>${formatNumber(params.capacity_mwh / params.power_mw, 1)} 小时</td></tr>
            <tr><td>运营期限</td><td>${params.operation_years} 年</td></tr>
            <tr><td>项目总投资</td><td class="key-value">${formatNumber(capex.dynamic_total)} 万EUR</td></tr>
            <tr><td>资本金</td><td>${formatNumber(capex.dynamic_total * params.equity_ratio)} 万EUR (${formatNumber(params.equity_ratio * 100, 0)}%)</td></tr>
            <tr><td>银行贷款</td><td>${formatNumber(capex.dynamic_total * (1 - params.equity_ratio))} 万EUR</td></tr>
            <tr><td>贷款期限</td><td>${params.loan_years} 年</td></tr>
            <tr><td>贷款利率</td><td>${formatNumber(params.loan_rate * 100)}%</td></tr>
        </table>
    </div>
    
    <div class="section highlight">
        <h2>二、核心财务指标</h2>
        <table>
            <tr><td>全投资内部收益率(税后)</td><td class="key-value">${formatNumber(indicators.project_irr)}%</td></tr>
            <tr><td>资本金内部收益率(税后)</td><td class="key-value">${formatNumber(indicators.equity_irr)}%</td></tr>
            <tr><td>静态投资回收期</td><td>${formatNumber(indicators.static_payback, 1)} 年</td></tr>
            <tr><td>动态投资回收期(8%)</td><td>${formatNumber(indicators.dynamic_payback, 1)} 年</td></tr>
            <tr><td>平均债务偿付覆盖率(DSCR)</td><td class="key-value">${formatNumber(indicators.dscr)}x</td></tr>
            <tr><td>LCOE(平准化度电成本)</td><td>${formatNumber(indicators.lcoe, 2)} EUR/MWh</td></tr>
            <tr><td>累计净利润</td><td>${formatNumber(indicators.total_net_profit)} 万EUR</td></tr>
            <tr><td>年均净利润</td><td>${formatNumber(indicators.avg_net_profit)} 万EUR</td></tr>
            <tr><td>EBITDA回报率</td><td>${formatNumber(indicators.ebitda_return)}%</td></tr>
        </table>
    </div>
    
    <div class="section">
        <h2>三、年度财务预测</h2>
        <table>
            <thead>
                <tr>
                    <th>年份</th>
                    <th>收入(万EUR)</th>
                    <th>OPEX(万EUR)</th>
                    <th>EBITDA(万EUR)</th>
                    <th>净利润(万EUR)</th>
                    <th>DSCR</th>
                </tr>
            </thead>
            <tbody>
                ${calculationResults.incomeData.map((income, i) => {
                    const dscr = calculationResults.loanData[i].payment > 0 ? 
                                formatNumber(income.ebitda / calculationResults.loanData[i].payment) + 'x' : '-';
                    return `<tr>
                        <td>第${income.year}年</td>
                        <td>${formatNumber(income.revenue)}</td>
                        <td>${formatNumber(income.opex)}</td>
                        <td>${formatNumber(income.ebitda)}</td>
                        <td>${formatNumber(income.netProfit)}</td>
                        <td>${dscr}</td>
                    </tr>`;
                }).join('')}
            </tbody>
        </table>
    </div>
    
    <div class="footer">
        <p>本报告由德国独立储能电站投资测算系统自动生成，仅供融资决策参考。</p>
        <p>This report is auto-generated for financing decision reference only.</p>
    </div>
</body>
</html>
        `;
        
        const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `储能电站融资报告_${params.power_mw}MW_${new Date().toISOString().slice(0, 10)}.html`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        alert('HTML报告导出成功！');
        
    } catch (error) {
        console.error('导出错误:', error);
        alert('HTML导出失败: ' + error.message);
    }
}

/**
 * 导出Excel版本
 * @description 提示用户运行Python脚本生成Excel文件
 */
function exportExcel() {
    const message = `Excel版本生成说明：

1. 确保已安装Python和openpyxl库：
   pip install openpyxl

2. 运行生成脚本：
   python generate_excel.py

3. Excel文件将保存在当前目录下，文件名包含时间戳。

4. 打开Excel文件后，修改"边界设定"和"设备配置"表中的参数，
   所有计算结果将自动更新。

注意：Excel版本与网页版功能完全同步，每次网页版更新后，
请重新运行 generate_excel.py 生成最新版本的Excel文件。`;
    
    alert(message);
}
