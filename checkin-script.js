// 全局变量
let headers = [];
let workbookData = [];
let companies = [];
let departments = [];
let dates = [];

// 人员名单按部门分类
const departmentPersonLists = {
    "行政部": ["朱甦雅", "王斌斌", "杜永丽", "张巧花", "曹娟"],
    "销售部": ["史正蓓", "陈炳森", "白金玉", "王莉莉", "董博", "王佳", "刘倩倩"],
    "医疗部": ["苏丹", "张改霞", "李雨婷", "赵改过", "张英玲", "胡立茹", "杨晓晖", "王烁楠", "朱丹", "朱英", "胡秀娥", "韩万忠", "程芸", "张文慧", "王倩倩", "孙悦", "郑晓燕", "魏霞", "姚建强", "潘丽萍", "朱艳丽", "董凯丽", "张鹤延", "孙绿萍", "刘丹", "张丽娟", "邵倩", "王晓霞", "党江娟", "刘沛", "周娜", "姜媛媛", "苏巧燕", "杨文娟"]
};

// 获取所有人员名单
function getAllPersons() {
    let allPersons = [];
    Object.values(departmentPersonLists).forEach(persons => {
        allPersons = [...allPersons, ...persons];
    });
    return allPersons;
}

// 获取人员所属部门
function getPersonDepartment(personName) {
    for (const [department, persons] of Object.entries(departmentPersonLists)) {
        if (persons.includes(personName)) {
            return department;
        }
    }
    return "医疗部"; // 不在行政部和销售部的人员默认属于医疗部
}

// 人员名单管理功能
function addPerson(name) {
    if (!name || name.trim() === '') {
        alert('请输入人员姓名,格式为"张三+销售部",例如"张三+销售部",果没有按严格的格式添加');
        return;
    }
    
    const trimmedName = name.trim();
    
    // 检查是否是姓名+部门的格式，比如"张三+销售部"
    const match = trimmedName.match(/^(.*?)\s*\+\s*(行政部|销售部|医疗部)$/);
    
    let personName, department;
    if (match) {
        // 提取姓名和部门
        personName = match[1].trim();
        department = match[2];
    } else {
        // 默认添加到医疗部
        personName = trimmedName;
        department = '医疗部';
    }
    
    const allPersons = getAllPersons();
    if (!allPersons.includes(personName)) {
        // 添加到对应的部门
        if (!departmentPersonLists[department]) {
            departmentPersonLists[department] = [];
        }
        departmentPersonLists[department].push(personName);
        alert(`人员添加成功，已添加到${department}`);
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
    let found = false;
    
    for (const [department, persons] of Object.entries(departmentPersonLists)) {
        const index = persons.indexOf(name);
        if (index !== -1) {
            persons.splice(index, 1);
            found = true;
            break;
        }
    }
    
    if (found) {
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
    let found = false;
    
    for (const [department, persons] of Object.entries(departmentPersonLists)) {
        const index = persons.indexOf(oldName);
        if (index !== -1) {
            persons[index] = newName;
            found = true;
            break;
        }
    }
    
    if (found) {
        alert('人员更新成功');
    } else {
        alert('旧姓名不存在');
    }
}

function getPersonList() {
    return getAllPersons();
}

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
        
        // 获取公司、部门和日期列表
        extractCompaniesDepartmentsAndDates();
        
        // 显示筛选区域
        populateFilters();
        document.getElementById('filterSection').style.display = 'block';
    };
    
    reader.readAsArrayBuffer(file);
});

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
    let dingxiSelected = false;
    uniqueCompanies.forEach(company => {
        const option = document.createElement('option');
        option.value = company;
        option.textContent = company;
        // 默认选择包含"定西"的公司
        if (company.includes('定西')) {
            option.selected = true;
            dingxiSelected = true;
        }
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
    let todaySelected = false;
    uniqueDates.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        // 默认选择当天日期
        if (isToday(date)) {
            option.selected = true;
            todaySelected = true;
        }
        dateFilter.appendChild(option);
    });
    
    // 如果没有找到当天日期，选择第一个日期
    if (!todaySelected && uniqueDates.length > 0) {
        dateFilter.options[1].selected = true;
    }
}

// 判断日期是否为当天
function isToday(dateString) {
    const today = new Date();
    const date = new Date(dateString);
    return date.toDateString() === today.toDateString();
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



// 对比人员名单功能
function comparePersonLists() {
    const companyFilter = document.getElementById('companyFilter').value;
    const departmentFilter = document.getElementById('departmentFilter').value;
    const dateFilter = document.getElementById('dateFilter').value;
    
    // 验证是否选择了公司和日期
    if (!companyFilter || !dateFilter) {
        alert('请选择公司和日期进行对比');
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
    
    // 根据选择的部门获取人员名单
    let targetPersonList = [];
    if (departmentFilter) {
        // 选择了特定部门
        if (departmentPersonLists[departmentFilter]) {
            targetPersonList = departmentPersonLists[departmentFilter];
        } else {
            alert('所选部门暂无人员名单');
            return;
        }
    } else {
        // 选择了全部部门
        targetPersonList = getAllPersons();
    }
    
    // 找出未打卡人员
    const missingPersons = targetPersonList.filter(person => 
        !excelPersonList.includes(person)
    );
    
    // 显示对比结果
    displayComparisonResult(missingPersons);
}

// 显示对比结果
function displayComparisonResult(missingPersons) {
    const comparisonSection = document.getElementById('comparisonSection');
    const missingPersonsList = document.getElementById('missingPersonsList');
    const missingCountEl = document.getElementById('missingCount');
    const resultTitleEl = document.getElementById('resultTitle');
    
    // 获取选择的日期
    const selectedDate = document.getElementById('dateFilter').value;
    
    // 更新结果标题，显示日期
    if (selectedDate) {
        resultTitleEl.textContent = `${selectedDate} 未打卡人员名单`;
    } else {
        resultTitleEl.textContent = '未打卡人员名单';
    }
    
    // 显示对比结果区域
    comparisonSection.style.display = 'block';
    
    // 更新未打卡人数
    missingCountEl.textContent = missingPersons.length;
    
    // 按部门分组未打卡人员
    const groupedPersons = {
        '行政部': [],
        '销售部': [],
        '医疗部': []
    };
    
    missingPersons.forEach(person => {
        const department = getPersonDepartment(person);
        groupedPersons[department].push(person);
    });
    
    // 显示未打卡人员名单
    if (missingPersons.length > 0) {
        let listHTML = '';
        
        // 按部门显示
        Object.entries(groupedPersons).forEach(([department, persons]) => {
            if (persons.length > 0) {
                // 获取部门对应的颜色类
                let departmentClass = '';
                switch (department) {
                    case '行政部':
                        departmentClass = 'admin';
                        break;
                    case '销售部':
                        departmentClass = 'sales';
                        break;
                    case '医疗部':
                        departmentClass = 'medical';
                        break;
                }
                
                listHTML += `<div class="department-group">
                    <h7>${department} (${persons.length}人)</h7>
                    <div class="department-persons">`;
                
                persons.forEach(person => {
                    listHTML += `<span class="person-item ${departmentClass}">${person}</span>`;
                });
                
                listHTML += `</div>
                </div>`;
            }
        });
        
        missingPersonsList.innerHTML = listHTML;
    } else {
        missingPersonsList.innerHTML = '<div class="no-persons">所有人员均已打卡</div>';
    }
}

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
        
        // 获取所有人员
        const allPersons = getAllPersons();
        // 更新人数
        personCount.textContent = allPersons.length;
        
        // 按部门分组显示
        for (const [department, persons] of Object.entries(departmentPersonLists)) {
            if (persons.length > 0) {
                const deptDiv = document.createElement('div');
                deptDiv.style.marginBottom = '15px';
                
                const deptTitle = document.createElement('h6');
                deptTitle.style.marginBottom = '8px';
                deptTitle.style.color = '#3b82f6';
                deptTitle.textContent = `${department} (${persons.length}人)`;
                deptDiv.appendChild(deptTitle);
                
                const ul = document.createElement('ul');
                ul.style.listStyleType = 'none';
                ul.style.padding = '0';
                ul.style.display = 'grid';
                ul.style.gridTemplateColumns = 'repeat(auto-fill, minmax(120px, 1fr))';
                ul.style.gap = '8px';
                
                persons.forEach((person, index) => {
                    const li = document.createElement('li');
                    li.style.padding = '5px 10px';
                    li.style.backgroundColor = '#e0f2fe';
                    li.style.borderRadius = '4px';
                    li.style.fontSize = '13px';
                    li.style.textAlign = 'center';
                    li.textContent = `${index + 1}. ${person}`;
                    ul.appendChild(li);
                });
                
                deptDiv.appendChild(ul);
                personListView.appendChild(deptDiv);
            }
        }
    }
}

// 页面加载完成后初始化
window.addEventListener('DOMContentLoaded', function() {
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
    
    // 导出为图片按钮事件监听器
    const exportToPngBtn = document.getElementById('exportToPngBtn');
    if (exportToPngBtn) {
        exportToPngBtn.addEventListener('click', function() {
            const comparisonSection = document.getElementById('comparisonSection');
            if (comparisonSection.style.display === 'block') {
                // 克隆结果区域，以便修改后导出
                const cloneSection = comparisonSection.cloneNode(true);
                
                // 隐藏克隆区域中的导出按钮
                const cloneExportButtons = cloneSection.querySelector('.export-buttons');
                if (cloneExportButtons) {
                    cloneExportButtons.style.display = 'none';
                }
                
                // 将克隆区域添加到页面中，以便html2canvas可以捕获它
                cloneSection.style.position = 'absolute';
                cloneSection.style.left = '-9999px';
                cloneSection.style.top = '-9999px';
                document.body.appendChild(cloneSection);
                
                // 使用html2canvas将结果区域转换为图片
                html2canvas(cloneSection, {
                    scale: 2, // 提高图片清晰度
                    useCORS: true, // 允许加载跨域图片
                    logging: false
                }).then(canvas => {
                    // 移除克隆区域
                    document.body.removeChild(cloneSection);
                    
                    // 创建下载链接
                    const link = document.createElement('a');
                    link.download = `未打卡人员名单_${new Date().toISOString().slice(0,10)}.png`;
                    link.href = canvas.toDataURL('image/png');
                    link.click();
                }).catch(error => {
                    // 移除克隆区域
                    if (cloneSection && cloneSection.parentNode) {
                        document.body.removeChild(cloneSection);
                    }
                    console.error('导出图片失败:', error);
                    alert('导出图片失败，请重试');
                });
            } else {
                alert('请先进行人员对比，生成结果后再导出');
            }
        });
    }
    
    // 导出为Excel按钮事件监听器
    const exportToExcelBtn = document.getElementById('exportToExcelBtn');
    if (exportToExcelBtn) {
        exportToExcelBtn.addEventListener('click', function() {
            const comparisonSection = document.getElementById('comparisonSection');
            if (comparisonSection.style.display === 'block') {
                const selectedDate = document.getElementById('dateFilter').value;
                const missingPersonsList = document.getElementById('missingPersonsList');
                
                // 提取未打卡人员名单
                const missingPersons = [];
                const listItems = missingPersonsList.querySelectorAll('li');
                listItems.forEach(item => {
                    missingPersons.push(item.textContent.trim());
                });
                
                // 创建Excel数据
                const data = [];
                data.push(['日期', selectedDate || '']);
                data.push(['未打卡人数', missingPersons.length]);
                data.push([]); // 空行
                data.push(['未打卡人员名单']);
                
                // 添加未打卡人员
                missingPersons.forEach(person => {
                    data.push([person]);
                });
                
                // 创建工作簿和工作表
                const ws = XLSX.utils.aoa_to_sheet(data);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, '未打卡人员');
                
                // 导出Excel文件
                XLSX.writeFile(wb, `未打卡人员名单_${selectedDate || new Date().toISOString().slice(0,10)}.xlsx`);
            } else {
                alert('请先进行人员对比，生成结果后再导出');
            }
        });
    }
});