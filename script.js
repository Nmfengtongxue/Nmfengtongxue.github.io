// 全局变量
let headers = [];
let workbookData = [];
let numericColumns = [];
let imageCountIndex = -1; 
let companies = [];
let departments = [];
let dates = [];
let isAutoFilling = false;

// 定西公司人员名单（从图片中提取）
let dingxiPersonList = [
    "朱甦雅", "王斌斌", "杜永丽", "张巧花", "曹娟", "史正蓓", "陈炳森", "白金玉",
    "王莉莉", "董博", "王佳", "刘倩倩", "苏丹", "张改霞", "李雨婷", "赵改过",
    "张英玲", "胡立茹", "杨晓晖", "王烁楠", "朱丹", "朱英", "胡秀娥", "韩万忠",
    "程芸", "张文慧", "王倩倩", "孙悦", "郑晓燕", "魏霞", "姚建强", "潘丽萍",
    "朱艳丽", "董凯丽", "张鹤延", "孙绿萍", "刘丹", "张丽娟", "邵倩", "王晓霞",
    "党江娟", "刘沛", "周娜", "姜媛媛", "苏巧燕", "杨文娟"
];

// 人员名单管理功能
function addPerson(name) {
    if (!name || name.trim() === '') {
        alert('请输入人员姓名');
        return;
    }
    
    name = name.trim();
    if (!dingxiPersonList.includes(name)) {
        dingxiPersonList.push(name);
        alert('人员添加成功');
    } else {
        alert('该人员已存在');
    }
}

function removePerson(name) {
    if (!name || name.trim() === '') {
        alert('请输入人员姓名');
        return;
    }
    
    name = name.trim();
    const index = dingxiPersonList.indexOf(name);
    if (index !== -1) {
        dingxiPersonList.splice(index, 1);
        alert('人员删除成功');
    } else {
        alert('该人员不存在');
    }
}

function updatePerson(oldName, newName) {
    if (!oldName || oldName.trim() === '' || !newName || newName.trim() === '') {
        alert('请输入旧姓名和新姓名');
        return;
    }
    
    oldName = oldName.trim();
    newName = newName.trim();
    const index = dingxiPersonList.indexOf(oldName);
    if (index !== -1) {
        dingxiPersonList[index] = newName;
        alert('人员更新成功');
    } else {
        alert('旧姓名不存在');
    }
}

function getPersonList() {
    return dingxiPersonList;
}

// 部门人数数据
const departmentCounts = {
    "定西": {"销售部": 6, "医疗部": 30, "行政部": 7},
    "陇南": {"销售部": 9, "医疗部": 37, "行政部": 9},
    "庆阳": {"销售部": 19, "医疗部": 48, "行政部": 5},
    "临夏": {"销售部": 11, "医疗部": 35, "行政部": 8},
    "西宁": {"销售部": 33, "医疗部": 56, "行政部": 7},
    "兰州": {"销售部": 53, "医疗部": 216, "行政部": 18},
    "武威": {"销售部": 15, "医疗部": 39, "行政部": 7},
    "平凉": {"销售部": 28, "医疗部": 31, "行政部": 5},
    "酒泉": {"销售部": 3, "医疗部": 48, "行政部": 11},
    "天水": {"销售部": 28, "医疗部": 50, "行政部": 5},
    "白银": {"销售部": 19, "医疗部": 30, "行政部": 4}
};

// 固定部门顺序：销售部、医疗部、行政部
const fixedDepartments = ["销售部", "医疗部", "行政部"];

// 合计人数数据
const totalPersonCounts = {
    "定西": 43,
    "陇南": 55,
    "庆阳": 72,
    "临夏": 54,
    "西宁": 96,
    "兰州": 287,
    "武威": 61,
    "平凉": 64,
    "酒泉": 62,
    "天水": 83,
    "白银": 53
};

// 上传并解析Excel文件
document.getElementById('uploadBtn').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
        alert('请先选择Excel文件');
        return;
    }
    
    const file = fileInput.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', WTF: true });
        
        // 获取第一个工作表数据
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (jsonData.length === 0) {
            alert('Excel文件中没有数据');
            return;
        }
        
        // 处理表头和数据
        headers = jsonData[0];
        workbookData = jsonData.slice(1);
        
        // 初始化索引
        imageCountIndex = headers.indexOf("请输入您本次上传图片数量：");
        
        // 处理数字列和图片数量
        processNumericColumns();
        
        // 获取公司、部门和日期列表
        extractCompaniesDepartmentsAndDates();
        
        // 显示数据和筛选区域
        displayData(workbookData);
        populateFilters();
        createForwardTable(); // 创建转发统计表格
        
        document.getElementById('filterSection').style.display = 'block';
        document.getElementById('dataTable').style.display = 'table';
        document.getElementById('autoFillBtn').disabled = false;
        
        // 更新合计
        updateSum();
    };
    
    reader.readAsArrayBuffer(file);
});

// 一键遍历填入功能
document.getElementById('autoFillBtn').addEventListener('click', async function() {
    if (isAutoFilling) {
        isAutoFilling = false;
        document.getElementById('autoFillBtn').textContent = '一键遍历填入';
        document.getElementById('progressText').textContent = '已取消';
        return;
    }
    
    // 验证是否有数据
    if (companies.length === 0 || departments.length === 0 || dates.length === 0) {
        alert('请先上传并解析包含完整数据的Excel文件');
        return;
    }
    
    isAutoFilling = true;
    this.textContent = '取消遍历';
    document.getElementById('progressContainer').style.display = 'block';
    
    const totalIterations = companies.length * departments.length * dates.length;
    let currentIteration = 0;
    
    // 遍历所有组合
    for (const company of companies) {
        if (!isAutoFilling) break;
        
        for (const department of departments) {
            if (!isAutoFilling) break;
            
            for (const date of dates) {
                if (!isAutoFilling) break;
                
                // 更新进度
                currentIteration++;
                const progress = Math.round((currentIteration / totalIterations) * 100);
                document.getElementById('progressBar').style.width = `${progress}%`;
                document.getElementById('progressText').textContent = 
                    `正在处理: ${company} - ${department} - ${date} (${currentIteration}/${totalIterations})`;
                
                // 设置筛选条件
                document.getElementById('companyFilter').value = company;
                document.getElementById('departmentFilter').value = department;
                document.getElementById('dateFilter').value = date;
                
                // 触发筛选和高亮
                applyFilters();
                highlightSelectedCell();
                
                // 直接调用统计转发次数的核心逻辑
                countForwardTimes();
                
                // 等待一段时间，确保操作完成
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        }
    }
    
    // 完成后重置状态
    if (isAutoFilling) {
        document.getElementById('progressText').textContent = '遍历完成!';
    }
    isAutoFilling = false;
    this.textContent = '一键遍历填入';
});

// 处理数字列
function processNumericColumns() {
    numericColumns = [];
    
    workbookData.forEach(row => {
        headers.forEach((header, index) => {
            // 处理图片数量列（取个位数）
            if (index === imageCountIndex && index < row.length && !isNaN(row[index]) && row[index] !== "" && row[index] !== null) {
                row[index] = parseInt(row[index]) % 10;
            }
            
            // 识别数字列
            if (index < row.length && !isNaN(row[index]) && row[index] !== "" && row[index] !== null) {
                if (!numericColumns.includes(index)) {
                    numericColumns.push(index);
                }
            }
        });
    });
}

// 提取公司、部门和日期列表
function extractCompaniesDepartmentsAndDates() {
    const companyIndex = headers.indexOf("您所在的公司：");
    const departmentIndex = headers.indexOf("您所在的部门：");
    const dateIndex = headers.indexOf("请选择打卡日期：（默认当前日期不可跨越选择）");
    
    // 获取去重的公司、部门和日期列表
    companies = [...new Set(workbookData.map(row => 
        companyIndex < row.length ? row[companyIndex] : '').filter(Boolean))];
    
    departments = [...new Set(workbookData.map(row => 
        departmentIndex < row.length ? row[departmentIndex] : '').filter(Boolean))];
        
    dates = [...new Set(workbookData.map(row => 
        dateIndex < row.length ? row[dateIndex] : '').filter(Boolean))];
}

// 增强的公司名称匹配算法 - 获取部门人数
function getDepartmentCount(company, dept) {
    // 定义默认数据的键名
    const defaultCompanies = Object.keys(departmentCounts);
    
    // 1. 精确匹配
    if (departmentCounts[company] && departmentCounts[company][dept]) {
        return departmentCounts[company][dept].toString();
    }
    
    // 2. 标准化匹配（去除空格、统一大小写）
    const normalizedCompany = company.trim().replace(/\s+/g, '').toLowerCase();
    for (const defaultCompany of defaultCompanies) {
        const normalizedDefault = defaultCompany.trim().replace(/\s+/g, '').toLowerCase();
        if (normalizedCompany === normalizedDefault) {
            return departmentCounts[defaultCompany][dept]?.toString() || '0';
        }
    }
    
    // 3. 子字符串匹配（处理公司名称前缀或后缀差异）
    for (const defaultCompany of defaultCompanies) {
        const normalizedDefault = defaultCompany.toLowerCase();
        if (normalizedCompany.includes(normalizedDefault) || normalizedDefault.includes(normalizedCompany)) {
            return departmentCounts[defaultCompany][dept]?.toString() || '0';
        }
    }
    
    // 4. 拼音首字母匹配（作为最后手段）
    const getFirstLetter = (str) => {
        // 简单的拼音首字母获取，只处理部分常见汉字
        const pinyinMap = {
            '定': 'd', '西': 'x', '陇': 'l', '南': 'n', '庆': 'q', '阳': 'y',
            '临': 'l', '夏': 'x', '西': 'x', '宁': 'n', '兰': 'l', '州': 'z',
            '武': 'w', '威': 'w', '平': 'p', '凉': 'l', '酒': 'j', '泉': 'q',
            '天': 't', '水': 's', '白': 'b', '银': 'y'
        };
        
        let result = '';
        for (const char of str) {
            result += pinyinMap[char] || char.toLowerCase();
        }
        return result;
    };
    
    const companyFirstLetter = getFirstLetter(company);
    for (const defaultCompany of defaultCompanies) {
        const defaultFirstLetter = getFirstLetter(defaultCompany);
        if (companyFirstLetter === defaultFirstLetter) {
            return departmentCounts[defaultCompany][dept]?.toString() || '0';
        }
    }
    
    // 如果所有匹配都失败，返回默认值0
    return '0';
}

// 增强的公司名称匹配算法 - 获取合计人数
function getTotalPersonCount(company) {
    // 定义默认数据的键名
    const defaultCompanies = Object.keys(totalPersonCounts);
    
    // 1. 精确匹配
    if (totalPersonCounts[company]) {
        return totalPersonCounts[company].toString();
    }
    
    // 2. 标准化匹配（去除空格、统一大小写）
    const normalizedCompany = company.trim().replace(/\s+/g, '').toLowerCase();
    for (const defaultCompany of defaultCompanies) {
        const normalizedDefault = defaultCompany.trim().replace(/\s+/g, '').toLowerCase();
        if (normalizedCompany === normalizedDefault) {
            return totalPersonCounts[defaultCompany]?.toString() || '0';
        }
    }
    
    // 3. 子字符串匹配（处理公司名称前缀或后缀差异）
    for (const defaultCompany of defaultCompanies) {
        const normalizedDefault = defaultCompany.toLowerCase();
        if (normalizedCompany.includes(normalizedDefault) || normalizedDefault.includes(normalizedCompany)) {
            return totalPersonCounts[defaultCompany]?.toString() || '0';
        }
    }
    
    // 4. 拼音首字母匹配（作为最后手段）
    const getFirstLetter = (str) => {
        // 简单的拼音首字母获取，只处理部分常见汉字
        const pinyinMap = {
            '定': 'd', '西': 'x', '陇': 'l', '南': 'n', '庆': 'q', '阳': 'y',
            '临': 'l', '夏': 'x', '西': 'x', '宁': 'n', '兰': 'l', '州': 'z',
            '武': 'w', '威': 'w', '平': 'p', '凉': 'l', '酒': 'j', '泉': 'q',
            '天': 't', '水': 's', '白': 'b', '银': 'y'
        };
        
        let result = '';
        for (const char of str) {
            result += pinyinMap[char] || char.toLowerCase();
        }
        return result;
    };
    
    const companyFirstLetter = getFirstLetter(company);
    for (const defaultCompany of defaultCompanies) {
        const defaultFirstLetter = getFirstLetter(defaultCompany);
        if (companyFirstLetter === defaultFirstLetter) {
            return totalPersonCounts[defaultCompany]?.toString() || '0';
        }
    }
    
    // 如果所有匹配都失败，返回默认值0
    return '0';
}

// 创建转发统计表格
function createForwardTable() {
    const forwardTable = document.getElementById('forwardTable');
    forwardTable.innerHTML = '';
    
    // 创建表头
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    // 添加排名列
    const rankTh = document.createElement('th');
    rankTh.textContent = '排名';
    headerRow.appendChild(rankTh);
    
    // 添加公司列
    const companyTh = document.createElement('th');
    companyTh.textContent = '公司';
    headerRow.appendChild(companyTh);
    
    // 添加部门人数和转发列标题（固定顺序：销售部、医疗部、行政部）
    fixedDepartments.forEach(dept => {
        // 添加部门人数列标题
        const countTh = document.createElement('th');
        countTh.textContent = dept;
        headerRow.appendChild(countTh);
        
        // 添加部门转发列标题
        const forwardTh = document.createElement('th');
        forwardTh.textContent = '转发(条)';
        headerRow.appendChild(forwardTh);
    });
    
    // 添加合计人数列
    const totalPersonTh = document.createElement('th');
    totalPersonTh.textContent = '合计(人)';
    headerRow.appendChild(totalPersonTh);
    
    // 添加合计转发数列
    const totalForwardTh = document.createElement('th');
    totalForwardTh.textContent = '宣发总量(条)'; // 修改表头文本
    headerRow.appendChild(totalForwardTh);
    
    // 移除宣发总量列
    // const totalPostTh = document.createElement('th');
    // totalPostTh.textContent = '宣发总量(条)';
    // headerRow.appendChild(totalPostTh);
    
    thead.appendChild(headerRow);
    forwardTable.appendChild(thead);
    
    // 创建表体
    const tbody = document.createElement('tbody');
    tbody.id = 'forward-table-body';
    
    // 使用Excel数据中的公司列表（如果有），否则使用默认数据
    let companyList = companies.length > 0 ? companies : Object.keys(departmentCounts);
    
    // 创建默认数据的键名映射表，用于更可靠的模糊匹配
    const defaultCompanies = Object.keys(departmentCounts);
    
    // 添加各公司/分院行
    companyList.forEach((company, index) => {
        const row = document.createElement('tr');
        row.classList.add('draggable-row'); // 添加可拖拽行的类名
        
        // 排名单元格
        const rankCell = document.createElement('td');
        rankCell.textContent = index + 1;
        row.appendChild(rankCell);
        
        // 公司名称单元格
        const nameCell = document.createElement('td');
        nameCell.textContent = company;
        row.appendChild(nameCell);
        
        // 各部门人数和转发次数单元格（固定顺序：销售部、医疗部、行政部）
        fixedDepartments.forEach(dept => {
            // 添加部门人数单元格
            const countCell = document.createElement('td');
            countCell.className = 'department-count';
            countCell.dataset.branch = company;
            countCell.dataset.dept = dept;
            countCell.contentEditable = true; // 设为可编辑
            
            // 填入默认人数数据，使用增强的匹配算法
            let departmentCount = getDepartmentCount(company, dept);
            countCell.textContent = departmentCount;
            
            // 添加编辑完成事件监听
            countCell.addEventListener('blur', function() {
                // 获取编辑后的值
                const newValue = this.textContent.trim();
                // 验证输入是否为数字
                if (newValue && !isNaN(parseInt(newValue))) {
                    // 更新单元格内容
                    this.textContent = parseInt(newValue);
                    // 更新总计行中的对应部门人数总计
                    updateDepartmentTotal(dept);
                    // 更新合计人数
                    updateTotalPersonCount(company);
                } else {
                    // 如果输入无效，恢复原来的值
                    this.textContent = departmentCount;
                }
            });
            
            // 添加键盘事件监听（按Enter键完成编辑）
            countCell.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    // 阻止默认的换行行为
                    e.preventDefault();
                    // 失去焦点以触发blur事件
                    this.blur();
                }
            });
            
            row.appendChild(countCell);
            
            // 添加部门转发次数单元格
            const forwardCell = document.createElement('td');
            forwardCell.className = 'forward-count';
            forwardCell.dataset.branch = company;
            forwardCell.dataset.dept = dept;
            forwardCell.textContent = '0'; // 默认值为0
            row.appendChild(forwardCell);
        });
        
        // 合计人数单元格
        const totalPersonCell = document.createElement('td');
        totalPersonCell.className = 'total-person-count';
        
        // 获取合计人数，使用增强的匹配算法
        let totalCount = getTotalPersonCount(company);
        totalPersonCell.textContent = totalCount;
        row.appendChild(totalPersonCell);
        
        // 合计转发数单元格（现在显示为宣发总量）
        const totalForwardCell = document.createElement('td');
        totalForwardCell.className = 'total-forward-count';
        totalForwardCell.textContent = '0';
        row.appendChild(totalForwardCell);
        
        // 移除宣发总量单元格
        // const totalPostCell = document.createElement('td');
        // totalPostCell.className = 'total-post-count';
        // totalPostCell.textContent = '';
        // row.appendChild(totalPostCell);
        
        tbody.appendChild(row);
    });
    
    // 添加部门总计行
    const deptTotalRow = document.createElement('tr');
    deptTotalRow.className = 'sum-row';
    
    // 排名总计标签（空）
    const rankTotalCell = document.createElement('td');
    rankTotalCell.textContent = '';
    deptTotalRow.appendChild(rankTotalCell);
    
    // 部门总计标签
    const totalLabelCell = document.createElement('td');
    totalLabelCell.textContent = '合 计';
    deptTotalRow.appendChild(totalLabelCell);
    
    // 各部门人数总计和转发总计
    let totalPersons = 0;
    
    fixedDepartments.forEach(dept => {
        // 部门人数总计单元格
        const countTotalCell = document.createElement('td');
        countTotalCell.className = 'dept-count-total';
        countTotalCell.dataset.dept = dept;
        
        // 计算并填入部门人数总计
        let totalCount = 0;
        companies.forEach(company => {
            // 使用增强的匹配算法获取部门人数
            const count = parseInt(getDepartmentCount(company, dept)) || 0;
            totalCount += count;
        });
        
        countTotalCell.textContent = totalCount;
        deptTotalRow.appendChild(countTotalCell);
        
        // 累计总人数
        totalPersons += totalCount;
        
        // 部门转发总计单元格
        const forwardTotalCell = document.createElement('td');
        forwardTotalCell.className = 'dept-total';
        forwardTotalCell.dataset.dept = dept;
        forwardTotalCell.textContent = '0';
        deptTotalRow.appendChild(forwardTotalCell);
    });
    
    // 合计人数总计
    const totalPersonTotalCell = document.createElement('td');
    totalPersonTotalCell.className = 'total-person-total';
    totalPersonTotalCell.textContent = totalPersons;
    deptTotalRow.appendChild(totalPersonTotalCell);
    
    // 合计转发数总计（现在显示为宣发总量）
    const totalForwardTotalCell = document.createElement('td');
    totalForwardTotalCell.className = 'total-forward-total';
    totalForwardTotalCell.textContent = '0';
    deptTotalRow.appendChild(totalForwardTotalCell);
    
    // 移除宣发总量总计
    // const totalPostTotalCell = document.createElement('td');
    // totalPostTotalCell.id = 'grand-total';
    // totalPostTotalCell.textContent = '0';
    // deptTotalRow.appendChild(totalPostTotalCell);
    
    tbody.appendChild(deptTotalRow);
    forwardTable.appendChild(tbody);
    
    // 初始化拖拽排序
    initDragSort();
}

// 初始化拖拽排序功能
function initDragSort() {
    const tbody = document.getElementById('forward-table-body');
    if (!tbody) return;
    
    // 初始化Sortable，仅对可拖拽行进行排序
    new Sortable(tbody, {
        animation: 150,
        ghostClass: 'sortable-ghost',
        handle: 'td:first-child', // 只能通过第一列拖拽
        filter: '.sum-row', // 排除总计行
        onEnd: function(evt) {
            // 拖拽结束后更新companies数组的顺序
            updateCompaniesOrder();
        }
    });
}

// 更新companies数组的顺序以匹配拖拽后的表格顺序
function updateCompaniesOrder() {
    const tbody = document.getElementById('forward-table-body');
    if (!tbody) return;
    
    // 获取拖拽后的行顺序（排除总计行）
    const rows = tbody.querySelectorAll('.draggable-row');
    const newOrder = Array.from(rows).map(row => {
        return row.querySelector('td:first-child').textContent;
    });
    
    // 更新companies数组
    companies = newOrder;
}

// 显示数据表格
function displayData(data) {
    const tableHead = document.querySelector('#dataTable thead');
    const tableBody = document.querySelector('#dataTable tbody');
    
    // 清空表格
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    
    // 创建表头
    const headRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headRow.appendChild(th);
    });
    tableHead.appendChild(headRow);
    
    // 创建表体
    data.forEach((row) => {
        const tr = document.createElement('tr');
        headers.forEach((header, colIndex) => {
            const td = document.createElement('td');
            
            if (colIndex < row.length && row[colIndex] !== undefined) {
                // 处理图片数量列
                if (colIndex === imageCountIndex && !isNaN(row[colIndex]) && row[colIndex] > 10) {
                    td.textContent = row[colIndex] % 10;
                } else {
                    td.textContent = row[colIndex];
                }
            } else {
                td.textContent = '';
            }
            
            // 数字列右对齐
            if (numericColumns.includes(colIndex)) {
                td.classList.add('number-cell');
            }
            
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}

// 填充筛选下拉框
function populateFilters() {
    const companyIndex = headers.indexOf("您所在的公司：");
    const departmentIndex = headers.indexOf("您所在的部门：");
    const dateIndex = headers.indexOf("请选择打卡日期：（默认当前日期不可跨越选择）");
    
    // 获取去重的值
    const uniqueCompanies = [...new Set(workbookData.map(row => 
        companyIndex < row.length ? row[companyIndex] : '').filter(Boolean))];
    
    const uniqueDepartments = [...new Set(workbookData.map(row => 
        departmentIndex < row.length ? row[departmentIndex] : '').filter(Boolean))];
    
    const uniqueDates = [...new Set(workbookData.map(row => 
        dateIndex < row.length ? row[dateIndex] : '').filter(Boolean))];
    
    // 填充公司下拉菜单
    const companyFilter = document.getElementById('companyFilter');
    companyFilter.innerHTML = '<option value="">全部</option>';
    uniqueCompanies.forEach(company => {
        const option = document.createElement('option');
        option.value = company;
        option.textContent = company;
        companyFilter.appendChild(option);
    });
    
    // 填充部门下拉菜单
    const departmentFilter = document.getElementById('departmentFilter');
    departmentFilter.innerHTML = '<option value="">全部</option>';
    uniqueDepartments.forEach(department => {
        const option = document.createElement('option');
        option.value = department;
        option.textContent = department;
        departmentFilter.appendChild(option);
    });
    
    // 填充日期下拉菜单
    const dateFilter = document.getElementById('dateFilter');
    dateFilter.innerHTML = '<option value="">全部</option>';
    uniqueDates.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        dateFilter.appendChild(option);
    });
    
    // 添加筛选事件监听器 - 筛选时高亮对应单元格
    companyFilter.addEventListener('change', function() {
        applyFilters();
        highlightSelectedCell();
    });
    departmentFilter.addEventListener('change', function() {
        applyFilters();
        highlightSelectedCell();
    });
    dateFilter.addEventListener('change', applyFilters);
}

// 应用筛选条件
function applyFilters() {
    const companyFilter = document.getElementById('companyFilter').value;
    const departmentFilter = document.getElementById('departmentFilter').value;
    const dateFilter = document.getElementById('dateFilter').value;
    
    const companyIndex = headers.indexOf("您所在的公司：");
    const departmentIndex = headers.indexOf("您所在的部门：");
    const dateIndex = headers.indexOf("请选择打卡日期：（默认当前日期不可跨越选择）");
    
    const filteredData = workbookData.filter(row => {
        const companyMatch = !companyFilter || 
            (companyIndex < row.length && row[companyIndex] === companyFilter);
        
        const departmentMatch = !departmentFilter || 
            (departmentIndex < row.length && row[departmentIndex] === departmentFilter);
        
        const dateMatch = !dateFilter || 
            (dateIndex < row.length && row[dateIndex] === dateFilter);
        
        return companyMatch && departmentMatch && dateMatch;
    });
    
    displayData(filteredData);
    updateSum(filteredData);
}

// 高亮显示选中的分院和部门对应的单元格
function highlightSelectedCell() {
    // 移除所有高亮
    document.querySelectorAll('.forward-count.highlight').forEach(cell => {
        cell.classList.remove('highlight');
    });
    
    const selectedBranch = document.getElementById('companyFilter').value;
    const selectedDept = document.getElementById('departmentFilter').value;
    
    // 只有当两个筛选条件都选中时才高亮
    if (selectedBranch && selectedDept) {
        const targetCell = document.querySelector(
            `.forward-count[data-branch="${selectedBranch}"][data-dept="${selectedDept}"]`
        );
        if (targetCell) {
            targetCell.classList.add('highlight');
        }
    }
}

// 更新合计
function updateSum(filteredData) {
    const dataToSum = filteredData || workbookData;
    
    if (numericColumns.length === 0) {
        document.getElementById('sumSection').style.display = 'none';
        return;
    }
    
    document.getElementById('sumSection').style.display = 'block';
    
    const sums = {};
    numericColumns.forEach(colIndex => {
        sums[headers[colIndex]] = 0;
    });
    
    dataToSum.forEach(row => {
        numericColumns.forEach(colIndex => {
            if (colIndex < row.length && row[colIndex] !== undefined && !isNaN(row[colIndex])) {
                sums[headers[colIndex]] += parseFloat(row[colIndex]);
            }
        });
    });
    
    let sumText = "";
    for (const [header, sum] of Object.entries(sums)) {
        sumText += `${header}: ${sum.toLocaleString()} | `;
    }
    
    document.getElementById('sumDisplay').textContent = sumText.slice(0, -3);
}

// 统计转发次数的核心逻辑
function countForwardTimes() {
    const companyFilter = document.getElementById('companyFilter').value;
    const departmentFilter = document.getElementById('departmentFilter').value;
    
    // 验证是否选择了公司和部门
    if (!companyFilter || !departmentFilter) {
        return; // 一键遍历中不弹出警告
    }
    
    // 获取筛选后的数据
    const filteredData = getFilteredData();

    // 计算图片数量合计（取个位数）
    let totalCount = 0;
    const imageQuantityIndex = headers.indexOf("请输入您本次上传图片数量：");
    filteredData.forEach(row => {
        if (imageQuantityIndex < row.length && !isNaN(row[imageQuantityIndex])) {
            totalCount += parseInt(row[imageQuantityIndex]) % 10;
        }
    });

    // 更新对应的转发次数单元格
    const targetCell = document.querySelector(
        `.forward-count[data-branch="${companyFilter}"][data-dept="${departmentFilter}"]`
    );
    
    if (targetCell) {
        // 在原有值基础上累加
        const currentValue = parseInt(targetCell.textContent) || 0;
        targetCell.textContent = currentValue + totalCount;
        
        // 更新总计
        updateBranchTotals(companyFilter);
        updateDeptTotals(departmentFilter);
        updateGrandTotal();
    }
}

// 获取筛选后的数据
function getFilteredData() {
    const companyFilter = document.getElementById('companyFilter').value;
    const departmentFilter = document.getElementById('departmentFilter').value;
    const dateFilter = document.getElementById('dateFilter').value;
    
    const companyIndex = headers.indexOf("您所在的公司：");
    const departmentIndex = headers.indexOf("您所在的部门：");
    const dateIndex = headers.indexOf("请选择打卡日期：（默认当前日期不可跨越选择）");
    
    return workbookData.filter(row => {
        const companyMatch = !companyFilter || 
            (companyIndex < row.length && row[companyIndex] === companyFilter);
        
        const departmentMatch = !departmentFilter || 
            (departmentIndex < row.length && row[departmentIndex] === departmentFilter);
        
        const dateMatch = !dateFilter || 
            (dateIndex < row.length && row[dateIndex] === dateFilter);
        
        return companyMatch && departmentMatch && dateMatch;
    });
}

// 更新分院总计
function updateBranchTotals(branch) {
    const cells = document.querySelectorAll(`.forward-count[data-branch="${branch}"]`);
    let total = 0;
    
    cells.forEach(cell => {
        total += parseInt(cell.textContent) || 0;
    });
    
    // 找到对应的总计单元格
    const rows = document.querySelectorAll('#forwardTable tbody tr');
    for (const row of rows) {
        const nameCell = row.querySelector('td:nth-child(2)'); // 公司名称单元格在第2列
        if (nameCell && nameCell.textContent === branch) {
            // 更新合计转发数单元格
            const totalForwardCell = row.querySelector('.total-forward-count');
            if (totalForwardCell) {
                totalForwardCell.textContent = total;
            }
            break;
        }
    }
}

// 更新部门总计（转发次数）
function updateDeptTotals(dept) {
    const cells = document.querySelectorAll(`.forward-count[data-dept="${dept}"]`);
    let total = 0;
    
    cells.forEach(cell => {
        total += parseInt(cell.textContent) || 0;
    });
    
    const totalCell = document.querySelector(`.dept-total[data-dept="${dept}"]`);
    if (totalCell) {
        totalCell.textContent = total;
    }
}

// 更新部门人数总计
function updateDepartmentTotal(dept) {
    const cells = document.querySelectorAll(`.department-count[data-dept="${dept}"]`);
    let total = 0;
    
    cells.forEach(cell => {
        total += parseInt(cell.textContent) || 0;
    });
    
    const totalCell = document.querySelector(`.dept-count-total[data-dept="${dept}"]`);
    if (totalCell) {
        totalCell.textContent = total;
    }
}

// 更新公司合计人数
function updateTotalPersonCount(company) {
    const cells = document.querySelectorAll(`.department-count[data-branch="${company}"]`);
    let total = 0;
    
    cells.forEach(cell => {
        total += parseInt(cell.textContent) || 0;
    });
    
    // 找到对应的公司行
    const rows = document.querySelectorAll('#forwardTable tbody tr');
    for (const row of rows) {
        const nameCell = row.querySelector('td:nth-child(2)'); // 公司名称单元格在第2列
        if (nameCell && nameCell.textContent === company) {
            // 更新合计人数单元格
            const totalPersonCell = row.querySelector('.total-person-count');
            if (totalPersonCell) {
                totalPersonCell.textContent = total;
            }
            break;
        }
    }
    
    // 更新总合计人数
    updateTotalPersonsTotal();
}

// 更新总合计人数
function updateTotalPersonsTotal() {
    const cells = document.querySelectorAll('.total-person-count');
    let total = 0;
    
    cells.forEach(cell => {
        total += parseInt(cell.textContent) || 0;
    });
    
    const totalCell = document.querySelector('.total-person-total');
    if (totalCell) {
        totalCell.textContent = total;
    }
}

// 更新总合计
function updateGrandTotal() {
    // 更新合计转发数总计（宣发总量总计）
    const totalForwardCells = document.querySelectorAll('.total-forward-count');
    let totalForwardGrandTotal = 0;
    
    totalForwardCells.forEach(cell => {
        totalForwardGrandTotal += parseInt(cell.textContent) || 0;
    });
    
    const totalForwardTotalCell = document.querySelector('.total-forward-total');
    if (totalForwardTotalCell) {
        totalForwardTotalCell.textContent = totalForwardGrandTotal;
    }
}

// 页面加载完成后自动创建转发统计表格
// 跟踪排序方向（true为降序，false为升序）
let isDescending = true;

window.addEventListener('DOMContentLoaded', function() {
    createForwardTable();
    
    // 添加排序按钮事件监听器
    const sortBtn = document.getElementById('sortBtn');
    if (sortBtn) {
        sortBtn.addEventListener('click', () => {
            sortByTotalForwardCount();
        });
    }
    
    // 添加Excel导出按钮事件监听器
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    if (exportExcelBtn) {
        exportExcelBtn.addEventListener('click', () => {
            exportToExcel();
        });
    }
    
    // 添加SVG导出按钮事件监听器
    const exportSvgBtn = document.getElementById('exportSvgBtn');
    if (exportSvgBtn) {
        exportSvgBtn.addEventListener('click', () => {
            exportToSVG();
        });
    }
    
    // 添加人员对比按钮事件监听器
    const compareBtn = document.getElementById('compareBtn');
    if (compareBtn) {
        compareBtn.addEventListener('click', comparePersonLists);
    }
    
    // 添加管理人员名单按钮事件监听器
    const managePersonsBtn = document.getElementById('managePersonsBtn');
    if (managePersonsBtn) {
        managePersonsBtn.addEventListener('click', function() {
            const personManagementSection = document.getElementById('personManagementSection');
            personManagementSection.style.display = 'block';
            
            // 初始化查看名单
            renderPersonList();
        });
    }
    
    // 初始化标签页切换
    initTabs();
    
    // 添加人员按钮事件监听器
    const addPersonBtn = document.getElementById('addPersonBtn');
    if (addPersonBtn) {
        addPersonBtn.addEventListener('click', function() {
            const addName = document.getElementById('addName').value;
            addPerson(addName);
            document.getElementById('addName').value = '';
            renderPersonList();
        });
    }
    
    // 删除人员按钮事件监听器
    const removePersonBtn = document.getElementById('removePersonBtn');
    if (removePersonBtn) {
        removePersonBtn.addEventListener('click', function() {
            const removeName = document.getElementById('removeName').value;
            removePerson(removeName);
            document.getElementById('removeName').value = '';
            renderPersonList();
        });
    }
    
    // 更新人员按钮事件监听器
    const updatePersonBtn = document.getElementById('updatePersonBtn');
    if (updatePersonBtn) {
        updatePersonBtn.addEventListener('click', function() {
            const oldName = document.getElementById('oldName').value;
            const newName = document.getElementById('newName').value;
            updatePerson(oldName, newName);
            document.getElementById('oldName').value = '';
            document.getElementById('newName').value = '';
            renderPersonList();
        });
    }
});

// 初始化标签页
function initTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            const tab = this.getAttribute('data-tab');
            
            // 切换标签按钮状态
            tabBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            
            // 切换标签内容
            tabPanes.forEach(pane => {
                pane.classList.remove('active');
                if (pane.id === tab + '-tab') {
                    pane.classList.add('active');
                }
            });
            
            // 如果切换到查看名单标签，重新渲染名单
            if (tab === 'view') {
                renderPersonList();
            }
        });
    });
}

// 渲染人员名单
function renderPersonList() {
    const personListView = document.getElementById('personListView');
    const personCount = document.getElementById('personCount');
    
    if (personListView && personCount) {
        // 清空现有内容
        personListView.innerHTML = '';
        
        // 更新人数
        personCount.textContent = dingxiPersonList.length;
        
        // 创建人员列表
        const ul = document.createElement('ul');
        ul.style.listStyleType = 'none';
        ul.style.padding = '0';
        ul.style.display = 'grid';
        ul.style.gridTemplateColumns = 'repeat(auto-fill, minmax(120px, 1fr))';
        ul.style.gap = '8px';
        
        dingxiPersonList.forEach((person, index) => {
            const li = document.createElement('li');
            li.style.padding = '5px 10px';
            li.style.backgroundColor = '#e0f2fe';
            li.style.borderRadius = '4px';
            li.style.fontSize = '13px';
            li.style.textAlign = 'center';
            li.textContent = `${index + 1}. ${person}`;
            ul.appendChild(li);
        });
        
        personListView.appendChild(ul);
    }
}

// 判断是否为定西公司的辅助函数
function isDingxiCompany(company) {
    // 1. 精确匹配
    if (company === '定西') {
        return true;
    }
    
    // 2. 标准化匹配（去除空格、统一大小写）
    const normalizedCompany = company.trim().replace(/\s+/g, '').toLowerCase();
    const normalizedDingxi = '定西'.trim().replace(/\s+/g, '').toLowerCase();
    if (normalizedCompany === normalizedDingxi) {
        return true;
    }
    
    // 3. 子字符串匹配（处理公司名称前缀或后缀差异）
    if (normalizedCompany.includes(normalizedDingxi)) {
        return true;
    }
    
    // 4. 拼音首字母匹配（作为最后手段）
    const getFirstLetter = (str) => {
        // 简单的拼音首字母获取，只处理部分常见汉字
        const pinyinMap = {
            '定': 'd', '西': 'x', '陇': 'l', '南': 'n', '庆': 'q', '阳': 'y',
            '临': 'l', '夏': 'x', '西': 'x', '宁': 'n', '兰': 'l', '州': 'z',
            '武': 'w', '威': 'w', '平': 'p', '凉': 'l', '酒': 'j', '泉': 'q',
            '天': 't', '水': 's', '白': 'b', '银': 'y'
        };
        
        let result = '';
        for (const char of str) {
            result += pinyinMap[char] || char.toLowerCase();
        }
        return result;
    };
    
    const companyFirstLetter = getFirstLetter(company);
    const dingxiFirstLetter = getFirstLetter('定西');
    if (companyFirstLetter === dingxiFirstLetter) {
        return true;
    }
    
    return false;
}

// 对比人员名单功能
function comparePersonLists() {
    const companyFilter = document.getElementById('companyFilter').value;
    const dateFilter = document.getElementById('dateFilter').value;
    
    // 验证是否选择了公司和日期
    if (!companyFilter || !dateFilter) {
        alert('请选择公司和日期进行对比');
        return;
    }
    
    // 验证是否选择了定西公司
    if (!isDingxiCompany(companyFilter)) {
        alert('当前功能仅支持定西公司的人员对比');
        return;
    }
    
    // 获取筛选后的数据
    const filteredData = getFilteredData();
    
    // 提取Excel中的人员名单
    const excelPersonList = [];
    const nameIndex = headers.findIndex(header => 
        header.includes('姓名') || header.includes('名字') || header.includes('人员')
    );
    
    if (nameIndex === -1) {
        alert('无法在Excel文件中找到姓名列');
        return;
    }
    
    filteredData.forEach(row => {
        if (nameIndex < row.length && row[nameIndex] !== '' && row[nameIndex] !== null) {
            excelPersonList.push(row[nameIndex].trim());
        }
    });
    
    // 找出未打卡人员
    const missingPersons = dingxiPersonList.filter(person => 
        !excelPersonList.includes(person)
    );
    
    // 显示对比结果
    displayComparisonResult(missingPersons, excelPersonList.length);
}

// 显示对比结果
function displayComparisonResult(missingPersons, presentCount) {
    const comparisonSection = document.getElementById('comparisonSection');
    const missingPersonsList = document.getElementById('missingPersonsList');
    const totalPersonsEl = document.getElementById('totalPersons');
    const presentPersonsEl = document.getElementById('presentPersons');
    const missingPersonsEl = document.getElementById('missingPersons');
    
    // 显示对比结果区域
    comparisonSection.style.display = 'block';
    
    // 更新统计信息
    totalPersonsEl.textContent = dingxiPersonList.length;
    presentPersonsEl.textContent = presentCount;
    missingPersonsEl.textContent = missingPersons.length;
    
    // 显示未打卡人员名单
    if (missingPersons.length > 0) {
        let listHTML = '<ul>';
        missingPersons.forEach(person => {
            listHTML += `<li>${person}</li>`;
        });
        listHTML += '</ul>';
        missingPersonsList.innerHTML = listHTML;
    } else {
        missingPersonsList.innerHTML = '<p class="text-success">所有人员均已打卡</p>';
    }
}

// 导出SVG
function exportToSVG() {
    const table = document.getElementById('forwardTable');
    if (!table) return;
    
    // 获取表格尺寸
    const tableRect = table.getBoundingClientRect();
    const width = tableRect.width + 40; // 增加边距
    const height = tableRect.height + 40;
    
    // 创建SVG内容
    let svgContent = `
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
    <rect width="100%" height="100%" fill="white"/>
    <g transform="translate(20, 20)">
`;
    
    // 获取表头
    const headers = table.querySelectorAll('thead th');
    const headerRow = table.querySelector('thead tr');
    if (headerRow) {
        const headerRect = headerRow.getBoundingClientRect();
        const headerHeight = headerRect.height;
        
        let x = 0;
        headers.forEach((header, index) => {
            const colRect = header.getBoundingClientRect();
            const colWidth = colRect.width;
            
            // 添加单元格背景
            svgContent += `
        <rect x="${x}" y="0" width="${colWidth}" height="${headerHeight}" fill="#f0f0f0" stroke="black" stroke-width="1"/>
            `;
            
            // 添加文本
            svgContent += `
        <text x="${x + colWidth / 2}" y="${headerHeight / 2 + 5}" text-anchor="middle" fill="black" font-size="12">${header.textContent.trim()}</text>
            `;
            
            x += colWidth;
        });
    }
    
    // 获取数据行
    const rows = table.querySelectorAll('tbody tr');
    let y = 30; // 表头高度
    rows.forEach((row, rowIndex) => {
        const rowHeight = 30; // 默认行高
        const cells = row.querySelectorAll('td');
        
        let x = 0;
        cells.forEach((cell, cellIndex) => {
            const headers = table.querySelectorAll('thead th');
            const header = headers[cellIndex];
            const colWidth = header ? header.getBoundingClientRect().width : 80;
            
            // 交替行背景色
            const fillColor = rowIndex % 2 === 0 ? 'white' : '#f9f9f9';
            
            // 添加单元格背景
            svgContent += `
        <rect x="${x}" y="${y}" width="${colWidth}" height="${rowHeight}" fill="${fillColor}" stroke="black" stroke-width="1"/>
            `;
            
            // 添加文本
            svgContent += `
        <text x="${x + colWidth / 2}" y="${y + rowHeight / 2 + 5}" text-anchor="middle" fill="black" font-size="12">${cell.textContent.trim()}</text>
            `;
            
            x += colWidth;
        });
        
        y += rowHeight;
    });
    
    // 关闭SVG标签
    svgContent += `
    </g>
</svg>
`;
    
    // 创建Blob对象
    const blob = new Blob([svgContent], { type: 'image/svg+xml' });
    
    // 创建下载链接
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', '转发次数统计.svg');
    link.style.visibility = 'hidden';
    
    // 添加到DOM并触发下载
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 导出Excel（XLSX格式）
function exportToExcel() {
    const table = document.getElementById('forwardTable');
    if (!table) return;
    
    // 准备导出的数据数组
    const exportData = [];
    
    // 获取表头
    const headers = table.querySelectorAll('thead th');
    const headerRow = Array.from(headers)
        .map(header => header.textContent.trim());
    exportData.push(headerRow);
    
    // 获取数据行
    const rows = table.querySelectorAll('tbody tr');
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        const dataRow = Array.from(cells)
            .map(cell => cell.textContent.trim());
        exportData.push(dataRow);
    });
    
    // 使用SheetJS创建工作簿
    const ws = XLSX.utils.aoa_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '转发次数统计');
    
    // 生成Excel文件并下载
    XLSX.writeFile(wb, '转发次数统计.xlsx');
}

// 按宣发总量排序
function sortByTotalForwardCount() {
    const tbody = document.getElementById('forward-table-body');
    const sortBtn = document.getElementById('sortBtn');
    if (!tbody || !sortBtn) return;
    
    // 保存合计行引用（如果存在）
    const totalRow = tbody.querySelector('.sum-row');
    
    // 如果存在合计行，先将其从表格中移除
    if (totalRow) {
        totalRow.remove();
    }
    
    // 获取所有数据行（排除总计行）
    const rows = Array.from(tbody.querySelectorAll('.draggable-row'));
    
    // 排序行
    rows.sort((a, b) => {
        // 获取宣发总量值
        const aTotalForward = parseInt(a.querySelector('.total-forward-count').textContent) || 0;
        const bTotalForward = parseInt(b.querySelector('.total-forward-count').textContent) || 0;
        
        // 根据当前排序方向排序
        if (isDescending) {
            // 降序排序（从大到小）
            return bTotalForward - aTotalForward;
        } else {
            // 升序排序（从小到大）
            return aTotalForward - bTotalForward;
        }
    });
    
    // 更新表格顺序
    rows.forEach(row => tbody.appendChild(row));
    
    // 如果有合计行，将其重新添加到表格底部
    if (totalRow) {
        tbody.appendChild(totalRow);
    }
    
    // 更新排名
    updateRanks();
    
    // 更新companies数组顺序
    updateCompaniesOrder();
    
    // 切换排序方向
    isDescending = !isDescending;
    
    // 更新按钮文本
    sortBtn.textContent = `按宣发总量${isDescending ? '降序' : '升序'}排序`;
}

// 更新排名
function updateRanks() {
    const tbody = document.getElementById('forward-table-body');
    if (!tbody) return;
    
    const rows = tbody.querySelectorAll('.draggable-row');
    rows.forEach((row, index) => {
        const rankCell = row.querySelector('td:first-child');
        if (rankCell) {
            rankCell.textContent = index + 1;
        }
    });
}