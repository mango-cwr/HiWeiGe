// 全局变量
let originalData = [];
let availableMonths = [];
let comparisonResult = [];
let selectedMonth1 = '';
let selectedMonth2 = '';

// DOM 元素
const fileUpload = document.getElementById('fileUpload');
const parseButton = document.getElementById('parseButton');
const monthSelectionCard = document.getElementById('monthSelectionCard');
const month1 = document.getElementById('month1');
const month2 = document.getElementById('month2');
const compareButton = document.getElementById('compareButton');
const resultCard = document.getElementById('resultCard');
const summary = document.getElementById('summary');
const resultTable = document.getElementById('resultTable');
const exportButton = document.getElementById('exportButton');
const resetButton = document.getElementById('resetButton');
const loading = document.getElementById('loading');

// 初始化事件监听
function initEventListeners() {
    // 文件上传监听
    fileUpload.addEventListener('change', function() {
        parseButton.disabled = !this.files.length;
    });
    
    // 解析文件按钮监听
    parseButton.addEventListener('click', parseExcelFile);
    
    // 月份选择监听
    month1.addEventListener('change', function() {
        selectedMonth1 = this.value;
        compareButton.disabled = !selectedMonth1 || !selectedMonth2;
    });
    
    month2.addEventListener('change', function() {
        selectedMonth2 = this.value;
        compareButton.disabled = !selectedMonth1 || !selectedMonth2;
    });
    
    // 对比分析按钮监听
    compareButton.addEventListener('click', compareMonths);
    
    // 导出按钮监听
    exportButton.addEventListener('click', exportResult);
    
    // 重置按钮监听
    resetButton.addEventListener('click', resetForm);
}

// 解析Excel文件
function parseExcelFile() {
    const file = fileUpload.files[0];
    if (!file) return;
    
    showLoading();
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            // 增强对xls文件的支持，添加适当的解析选项
            const workbook = XLSX.read(data, { 
                type: 'array',
                cellDates: true,
                cellNF: false,
                cellText: false
            });
            
            // 检查是否有工作表
            if (workbook.SheetNames.length === 0) {
                hideLoading();
                alert('Excel文件中没有找到工作表');
                return;
            }
            
            // 获取第一个工作表
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // 转换为JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            // 检查是否有数据
            if (jsonData.length === 0) {
                hideLoading();
                alert('Excel文件中没有找到数据');
                return;
            }
            
            // 验证数据格式
            if (!validateData(jsonData)) {
                hideLoading();
                alert('数据格式不正确，必须包含：设备号码、账务周期、账单费用列');
                return;
            }
            
            // 处理数据
            originalData = processData(jsonData);
            
            // 提取可用月份
            availableMonths = extractMonths(originalData);
            
            // 检查是否提取到月份
            if (availableMonths.length === 0) {
                hideLoading();
                alert('无法从账务周期中提取月份，请检查账务周期格式是否正确');
                return;
            }
            
            // 填充月份下拉框
            populateMonthDropdowns();
            
            // 显示月份选择卡片
            monthSelectionCard.style.display = 'block';
            
            hideLoading();
        } catch (error) {
            hideLoading();
            console.error('解析文件失败:', error);
            // 显示更具体的错误信息
            alert(`解析文件失败: ${error.message}\n请检查文件格式是否正确，或尝试将文件另存为.xlsx格式后再上传`);
        }
    };
    
    reader.onerror = function() {
        hideLoading();
        alert('读取文件失败，请检查文件是否损坏或被其他程序占用');
    };
    
    reader.readAsArrayBuffer(file);
}

// 验证数据格式
function validateData(data) {
    if (!data || data.length === 0) return false;
    
    const firstRow = data[0];
    return '设备号码' in firstRow && '账务周期' in firstRow && '账单费用' in firstRow;
}

// 处理数据
function processData(data) {
    return data.map(item => {
        // 提取月份
        const month = extractMonthFromCycle(item['账务周期']);
        
        return {
            设备号码: item['设备号码'],
            账务周期: item['账务周期'],
            账单费用: parseFloat(item['账单费用']) || 0,
            month: month
        };
    });
}

// 从账务周期中提取月份
function extractMonthFromCycle(cycle) {
    try {
        // 格式: [20240701]2024-07-01:2024-07-31
        const datePart = cycle.split(']')[0].replace(/^\[|\]$/g, '');
        return `${datePart.substring(0, 4)}-${datePart.substring(4, 6)}`;
    } catch (error) {
        return '';
    }
}

// 提取可用月份
function extractMonths(data) {
    const months = new Set();
    data.forEach(item => {
        if (item.month) {
            months.add(item.month);
        }
    });
    return Array.from(months).sort();
}

// 填充月份下拉框
function populateMonthDropdowns() {
    // 清空现有选项
    month1.innerHTML = '<option value="">请选择月份</option>';
    month2.innerHTML = '<option value="">请选择月份</option>';
    
    // 添加新选项
    availableMonths.forEach(month => {
        const option1 = document.createElement('option');
        option1.value = month;
        option1.textContent = month;
        month1.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = month;
        option2.textContent = month;
        month2.appendChild(option2);
    });
}

// 对比月份
function compareMonths() {
    showLoading();
    
    setTimeout(() => {
        try {
            // 过滤两个月份的数据
            const data1 = originalData.filter(item => item.month === selectedMonth1);
            const data2 = originalData.filter(item => item.month === selectedMonth2);
            
            // 按设备号码分组
            const data1ByDevice = groupByDevice(data1);
            const data2ByDevice = groupByDevice(data2);
            
            // 获取所有设备号码
            const allDevices = new Set([...Object.keys(data1ByDevice), ...Object.keys(data2ByDevice)]);
            
            // 计算差异
            comparisonResult = [];
            allDevices.forEach(device => {
                const amount1 = data1ByDevice[device] || 0;
                const amount2 = data2ByDevice[device] || 0;
                const diff = amount2 - amount1;
                const diffPercent = amount1 ? (diff / amount1 * 100).toFixed(2) : 'inf';
                
                // 只添加差异不为0的记录
                if (diff !== 0) {
                    comparisonResult.push({
                        设备号码: device,
                        [`账单费用_${selectedMonth1}`]: amount1,
                        [`账单费用_${selectedMonth2}`]: amount2,
                        差异金额: diff,
                        差异百分比: diffPercent
                    });
                }
            });
            
            // 显示结果
            displayResult();
            
            hideLoading();
        } catch (error) {
            hideLoading();
            console.error('对比分析失败:', error);
            alert('对比分析失败，请重试');
        }
    }, 500);
}

// 按设备号码分组
function groupByDevice(data) {
    const grouped = {};
    data.forEach(item => {
        if (!grouped[item.设备号码]) {
            grouped[item.设备号码] = 0;
        }
        grouped[item.设备号码] += item.账单费用;
    });
    return grouped;
}

// 显示结果
function displayResult() {
    // 计算汇总信息
    const totalDiff = comparisonResult.reduce((sum, item) => sum + item.差异金额, 0);
    const avgDiff = totalDiff / comparisonResult.length;
    const maxDiffItem = comparisonResult.reduce((max, item) => 
        Math.abs(item.差异金额) > Math.abs(max.差异金额) ? item : max, comparisonResult[0]
    );
    
    // 显示汇总信息
    summary.innerHTML = `
        <h3>汇总信息</h3>
        <p><strong>总差异金额:</strong> ${totalDiff.toFixed(2)}</p>
        <p><strong>平均差异金额:</strong> ${avgDiff.toFixed(2)}</p>
        <p><strong>差异最大的设备:</strong> ${maxDiffItem.设备号码} (差异: ${maxDiffItem.差异金额.toFixed(2)})</p>
    `;
    
    // 清空表格
    const tbody = resultTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    // 填充表格
    comparisonResult.forEach(item => {
        const row = tbody.insertRow();
        row.insertCell(0).textContent = item.设备号码;
        row.insertCell(1).textContent = item[`账单费用_${selectedMonth1}`].toFixed(2);
        row.insertCell(2).textContent = item[`账单费用_${selectedMonth2}`].toFixed(2);
        row.insertCell(3).textContent = item.差异金额.toFixed(2);
        row.insertCell(4).textContent = item.差异百分比;
    });
    
    // 显示结果卡片
    resultCard.style.display = 'block';
}

// 导出结果
function exportResult() {
    if (!comparisonResult.length) return;
    
    // 创建工作表
    const ws = XLSX.utils.json_to_sheet(comparisonResult);
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '对比结果');
    
    // 导出文件
    const fileName = `bill_comparison_${selectedMonth1}_vs_${selectedMonth2}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

// 重置表单
function resetForm() {
    // 清空文件上传
    fileUpload.value = '';
    parseButton.disabled = true;
    
    // 隐藏月份选择卡片
    monthSelectionCard.style.display = 'none';
    
    // 清空月份选择
    month1.innerHTML = '<option value="">请选择月份</option>';
    month2.innerHTML = '<option value="">请选择月份</option>';
    selectedMonth1 = '';
    selectedMonth2 = '';
    compareButton.disabled = true;
    
    // 隐藏结果卡片
    resultCard.style.display = 'none';
    
    // 清空数据
    originalData = [];
    availableMonths = [];
    comparisonResult = [];
}

// 显示加载动画
function showLoading() {
    const loading = document.getElementById('loading');
    if (loading) {
        loading.style.display = 'flex';
        loading.style.visibility = 'visible';
        loading.style.opacity = '1';
        loading.removeAttribute('hidden');
    }
}

// 隐藏加载动画
function hideLoading() {
    const loading = document.getElementById('loading');
    if (loading) {
        loading.style.display = 'none';
        loading.style.visibility = 'hidden';
        loading.style.opacity = '0';
        loading.setAttribute('hidden', 'hidden');
    }
}

// 页面加载时隐藏加载动画
window.addEventListener('load', function() {
    hideLoading();
});

// 扩展字符串方法
String.prototype.strip = function(char) {
    return this.replace(new RegExp(`^\\${char}+|\\${char}+$`, 'g'), '');
};

// 初始化
function init() {
    // 首先隐藏加载动画
    hideLoading();
    // 然后初始化事件监听
    initEventListeners();
}

// 页面加载完成后初始化
window.addEventListener('DOMContentLoaded', init);
