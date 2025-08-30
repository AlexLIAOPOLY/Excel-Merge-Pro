// 全局变量
let currentData = [];
let filteredData = [];
let currentSchema = [];
let tableGroups = [];
let currentGroupId = null;


// 控制台日志函数 - 修复：处理控制台元素不存在的情况
function addConsoleLog(message, type = 'system') {
    const consoleElement = document.getElementById('console');
    
    // 如果控制台元素不存在（已被删除），则静默跳过
    if (!consoleElement) {
        // 可选：在浏览器控制台输出日志，方便调试
        window.console && window.console.log(`[${type.toUpperCase()}] ${message}`);
        return;
    }
    
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry log-${type}`;
    
    const timestamp = new Date().toLocaleTimeString();
    logEntry.textContent = `[${timestamp}] ${message}`;
    
    consoleElement.appendChild(logEntry);
    consoleElement.scrollTop = consoleElement.scrollHeight;
    
    // 限制日志条数
    const logs = consoleElement.children;
    if (logs.length > 50) {
        consoleElement.removeChild(logs[0]);
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    addConsoleLog('DataMerge Pro 系统初始化完成', 'system');
    
    initializeEventListeners();
    initializeNavbar();
    loadTableList();
});

// 初始化导航栏功能
function initializeNavbar() {
    const navbar = document.getElementById('navbar');
    const navToggle = document.getElementById('navToggle');
    const navMenuMobile = document.getElementById('navMenuMobile');
    
    // 滚动效果
    let lastScrollY = window.scrollY;
    
    function handleScroll() {
        const scrollY = window.scrollY;
        
        if (scrollY > 50) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
        
        // 向下滚动时隐藏移动端菜单
        if (scrollY > lastScrollY && scrollY > 100) {
            if (navMenuMobile.classList.contains('active')) {
                toggleMobileMenu();
            }
        }
        
        lastScrollY = scrollY;
    }
    
    // 移动端菜单切换
    function toggleMobileMenu() {
        navToggle.classList.toggle('active');
        navMenuMobile.classList.toggle('active');
        
        // 阻止背景滚动
        if (navMenuMobile.classList.contains('active')) {
            document.body.style.overflow = 'hidden';
            addConsoleLog('移动端菜单已打开', 'system');
        } else {
            document.body.style.overflow = '';
            addConsoleLog('移动端菜单已关闭', 'system');
        }
    }
    
    // 绑定事件
    window.addEventListener('scroll', handleScroll, { passive: true });
    
    if (navToggle) {
        navToggle.addEventListener('click', function(e) {
            e.stopPropagation();
            toggleMobileMenu();
        });
    }
    
    // 点击菜单项后关闭移动端菜单
    if (navMenuMobile) {
        navMenuMobile.addEventListener('click', function(e) {
            if (e.target.classList.contains('nav-link')) {
                toggleMobileMenu();
            }
        });
    }
    
    // 点击页面其他地方关闭移动端菜单
    document.addEventListener('click', function(e) {
        if (navMenuMobile.classList.contains('active')) {
            if (!navMenuMobile.contains(e.target) && !navToggle.contains(e.target)) {
                toggleMobileMenu();
            }
        }
    });
    
    // ESC键关闭移动端菜单
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && navMenuMobile.classList.contains('active')) {
            toggleMobileMenu();
        }
        

    });
    

    
    addConsoleLog('导航栏功能初始化完成', 'system');
}

// 初始化表格选择器事件
function initializeTableSelector() {
    const tableSelector = document.getElementById('tableSelector');
    if (tableSelector) {
        // 移除之前的监听器（如果存在）
        tableSelector.removeEventListener('change', handleTableSelectorChange);
        // 添加新的监听器
        tableSelector.addEventListener('change', handleTableSelectorChange);
    }
}

// 表格选择器变化处理函数
function handleTableSelectorChange() {
    const selectedGroupId = this.value;
    if (selectedGroupId && selectedGroupId !== 'back-to-groups' && selectedGroupId !== 'single-file-view') {
        currentGroupId = parseInt(selectedGroupId);
        loadGroupData(currentGroupId);
    }
}

// 初始化事件监听器
function initializeEventListeners() {
    const fileInput = document.getElementById('fileInput');
    const uploadArea = document.getElementById('uploadArea');
    const fileList = document.getElementById('fileList');
    const searchInput = document.getElementById('searchInput');
    
    // 文件选择事件 - 选择后自动上传
    fileInput.addEventListener('change', function() {
        const files = Array.from(this.files);
        updateFileList(files);
        
        if (files.length > 0) {
            addConsoleLog(`已选择 ${files.length} 个文件，开始自动上传`, 'system');
            // 自动上传
            uploadFiles();
        }
    });
    
    // 搜索框回车事件和实时搜索
    let searchTimeout;
    searchInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            clearTimeout(searchTimeout);
            searchData();
        }
    });
    
    // 实时搜索 - 输入时延迟搜索
    searchInput.addEventListener('input', function(e) {
        clearTimeout(searchTimeout);
        searchTimeout = setTimeout(() => {
            const currentValue = this.value.trim();
            if (currentValue.length >= 1 || currentValue.length === 0) {
                // 当输入至少1个字符或清空时触发搜索
                searchData();
            }
        }, 300); // 300ms延迟，避免过于频繁的搜索
    });
    
    // 拖拽上传
    uploadArea.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', function(e) {
        e.preventDefault();
        this.classList.remove('dragover');
        
        const files = Array.from(e.dataTransfer.files);
        const excelFiles = files.filter(file => 
            file.name.toLowerCase().endsWith('.xlsx') || 
            file.name.toLowerCase().endsWith('.xls')
        );
        
        if (excelFiles.length > 0) {
            const dt = new DataTransfer();
            excelFiles.forEach(file => dt.items.add(file));
            fileInput.files = dt.files;
            
            updateFileList(excelFiles);
            addConsoleLog(`通过拖拽选择了 ${excelFiles.length} 个Excel文件，开始自动上传`, 'system');
            // 自动上传
            uploadFiles();
        } else {
            addConsoleLog('请选择Excel文件 (.xlsx 或 .xls)', 'warning');
        }
    });
}

// 更新文件列表显示
function updateFileList(files) {
    const fileList = document.getElementById('fileList');
    
    if (files.length === 0) {
        fileList.style.display = 'none';
        return;
    }
    
    fileList.style.display = 'block';
    fileList.innerHTML = '';
    
    files.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <span class="file-name">${file.name}</span>
            <span class="file-size">${formatFileSize(file.size)}</span>
        `;
        fileList.appendChild(fileItem);
    });
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 上传文件
async function uploadFiles() {
    const fileInput = document.getElementById('fileInput');
    const files = fileInput.files;
    
    if (files.length === 0) {
        addConsoleLog('请先选择文件', 'warning');
        return;
    }
    
    addConsoleLog(`开始上传 ${files.length} 个文件...`, 'system');
    
    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
        formData.append('files[]', files[i]);
    }
    
    // 显示上传状态
    addConsoleLog('文件上传中...', 'system');
    
    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        displayUploadResults(result.results);
        
        // 刷新表格列表和数据
        await Promise.all([
            loadTableList(),
            loadTablesList()  // 刷新文件管理列表
        ]);
        
        // 清空文件选择
        fileInput.value = '';
        document.getElementById('fileList').style.display = 'none';
        
    } catch (error) {
        addConsoleLog(`上传失败: ${error.message}`, 'error');
    } finally {
        addConsoleLog('文件上传完成', 'system');
    }
}

// 显示上传结果 - 改为右侧弹窗形式
function displayUploadResults(results) {
    let totalSuccess = 0;
    let totalRecords = 0;
    
    results.forEach((result, index) => {
        setTimeout(() => {
            if (result.success) {
                showNotification(
                    '上传成功',
                    `${result.filename}: 成功导入 ${result.count} 条记录`,
                    'success'
                );
                totalSuccess++;
                totalRecords += result.count;
                addConsoleLog(`${result.filename} 处理成功，导入 ${result.count} 条记录`, 'system');
                
                // 记录最新的分组ID
                if (result.group_id) {
                    currentGroupId = result.group_id;
                }
            } else {
                showNotification(
                    '上传失败',
                    `${result.filename}: ${result.message}`,
                    'error'
                );
                addConsoleLog(`${result.filename} 处理失败: ${result.message}`, 'error');
            }
        }, index * 200); // 错开显示时间
    });
    
    // 最后显示汇总信息
    if (results.length > 1) {
        setTimeout(() => {
            const successCount = results.filter(r => r.success).length;
            const totalCount = results.reduce((sum, r) => sum + (r.success ? (r.count || 0) : 0), 0);
            if (successCount > 0) {
                showNotification(
                    '批量上传完成',
                    `成功处理 ${successCount} 个文件，共导入 ${totalCount} 条记录`,
                    'info'
                );
                addConsoleLog(`上传完成: ${successCount}/${results.length} 个文件成功，共导入 ${totalCount} 条记录`, 'system');
                
                // 批量上传成功后自动折叠上传区域
                autoCollapseUploadArea();
            }
        }, results.length * 200 + 500);
    } else {
        // 单文件上传成功后也自动折叠
        const hasSuccess = results.some(r => r.success);
        if (hasSuccess) {
            autoCollapseUploadArea();
        }
    }
}

// 显示弹窗通知
function showNotification(title, message, type = 'info') {
    const container = document.getElementById('notificationContainer');
    
    // 创建通知元素
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    
    notification.innerHTML = `
        <div class="notification-header">
            <div class="notification-title">${title}</div>
        </div>
        <div class="notification-message">${message}</div>
        <div class="notification-progress"></div>
    `;
    
    // 添加到容器
    container.appendChild(notification);
    
    // 触发显示动画
    setTimeout(() => {
        notification.classList.add('show');
    }, 10);
    
    // 自动隐藏（1秒后开始隐藏动画）
    setTimeout(() => {
        notification.classList.add('hide');
        setTimeout(() => {
            if (container.contains(notification)) {
                container.removeChild(notification);
            }
        }, 400); // 等待动画完成
    }, 1000);
}

// 加载数据
async function loadData() {
    addConsoleLog('正在加载数据...', 'system');
    
    const loadingState = document.getElementById('loadingState');
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    
    loadingState.style.display = 'flex';
    dataTable.style.display = 'none';
    emptyState.style.display = 'none';
    
    try {
        const response = await fetch('/data');
        const result = await response.json();
        
        if (result.success) {
            currentData = result.data;
            filteredData = [...currentData];
            currentSchema = result.schema || [];
            
            updateStats(result.stats);
            renderTable();
            
            addConsoleLog(`数据加载完成，共 ${currentData.length} 条记录，${currentSchema.length} 个列`, 'system');
        } else {
            addConsoleLog(`加载数据失败: ${result.message}`, 'error');
            showEmptyState();
        }
    } catch (error) {
        addConsoleLog(`加载数据时发生错误: ${error.message}`, 'error');
        showEmptyState();
    } finally {
        loadingState.style.display = 'none';
    }
}

// 显示空状态
function showEmptyState() {
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    
    dataTable.style.display = 'none';
    emptyState.style.display = 'block';
    
    // 重置为默认的空状态内容
    emptyState.innerHTML = `
        <h3>暂无数据</h3>
        <p>请上传Excel文件开始使用</p>
        <button class="btn btn-primary" onclick="document.getElementById('fileInput').click()">上传第一个文件</button>
    `;
    
    // 隐藏工具栏按钮
    hideToolbarButtons();
}

// 显示搜索空状态
function showSearchEmptyState(searchTerm) {
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    
    dataTable.style.display = 'none';
    emptyState.style.display = 'block';
    
    // 显示搜索无结果的内容
    emptyState.innerHTML = `
        <h3>搜索无结果</h3>
        <p>未找到包含 <strong>"${searchTerm}"</strong> 的记录</p>
        <p style="color: #666; font-size: 14px;">
            搜索范围包括：所有列名和数据内容<br>
            支持模糊匹配和拼音首字母搜索
        </p>
        <div style="margin-top: 20px;">
            <button class="btn btn-secondary" onclick="resetSearch()">重置搜索</button>
            <button class="btn btn-primary" onclick="document.getElementById('searchInput').focus()">修改搜索</button>
        </div>
    `;
    
    // 保持工具栏按钮显示
    showToolbarButtons();
}

// 控制工具栏按钮的显示/隐藏
function hideToolbarButtons() {
    const toolbarRight = document.querySelector('.toolbar-right');
    if (toolbarRight) {
        toolbarRight.style.display = 'none';
    }
}

function showToolbarButtons() {
    const toolbarRight = document.querySelector('.toolbar-right');
    if (toolbarRight) {
        toolbarRight.style.display = 'flex';
    }
}

// 更新统计信息 - 修复：处理统计元素不存在的情况
function updateStats(stats) {
    // 检查统计元素是否存在（因为我们可能删除了数据概览部分）
    const totalRecordsEl = document.getElementById('totalRecords');
    const sourceFilesEl = document.getElementById('sourceFiles');
    const totalColumnsEl = document.getElementById('totalColumns');
    const lastUpdateEl = document.getElementById('lastUpdate');
    
    // 如果统计元素不存在，静默跳过
    if (!totalRecordsEl || !sourceFilesEl || !totalColumnsEl) {
        // 在浏览器控制台输出统计信息，方便调试
        if (stats && window.console) {
            window.console.log('统计信息:', {
                总记录: stats.total_records || 0,
                来源文件: stats.source_files || 0,
                数据列: stats.total_columns || 0
            });
        }
        return;
    }
    
    // 如果元素存在，正常更新
    if (stats) {
        totalRecordsEl.textContent = stats.total_records || 0;
        sourceFilesEl.textContent = stats.source_files || 0;
        totalColumnsEl.textContent = stats.total_columns || 0;
    }
    
    // 更新时间显示
    if (lastUpdateEl) {
        const now = new Date();
        const timeString = now.toLocaleTimeString('zh-CN', { 
            hour12: false,
            hour: '2-digit', 
            minute: '2-digit', 
            second: '2-digit' 
        });
        lastUpdateEl.textContent = timeString;
    }
}

// 渲染表格
function renderTable() {
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    const tableHeader = document.getElementById('tableHeader');
    const tableBody = document.getElementById('dataBody');
    
    // 检查是否有搜索条件
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
    const hasSearchTerm = searchTerm.length > 0;
    
    if (filteredData.length === 0 || currentSchema.length === 0) {
        if (hasSearchTerm && currentData.length > 0) {
            // 有搜索条件但没找到结果
            showSearchEmptyState(searchTerm);
        } else {
            // 没有数据或没有搜索条件
            showEmptyState();
        }
        return;
    }
    
    // 显示表格
    dataTable.style.display = 'table';
    emptyState.style.display = 'none';
    
    // 显示工具栏按钮
    showToolbarButtons();
    
    // 构建表头
    const headerRow = document.createElement('tr');
    const baseColumnWidth = 150; // 基础列宽
    const totalColumns = currentSchema.length + 1; // 包括操作列
    const actionColumnWidth = 100;
    
    currentSchema.forEach((col, index) => {
        const th = document.createElement('th');
        th.className = 'resizable-col';
        th.innerHTML = `
            <div class="table-header-actions">
                <span class="column-title" ondblclick="showRenameColumnModal('${col}')" title="双击重命名列">${col}</span>
                <div class="column-actions">
                    <button class="column-btn" onclick="sortColumn('${col}', 'asc')" title="升序排序">↑</button>
                    <button class="column-btn" onclick="sortColumn('${col}', 'desc')" title="降序排序">↓</button>
                    <button class="column-btn column-delete-btn" onclick="deleteColumnConfirm('${col}')" title="删除此列">×</button>
                </div>
            </div>
            <div class="col-resizer" data-column="${index}"></div>
            <div class="col-inserter" onclick="insertColumnAfter(${index})" title="在此列后插入新列"></div>
        `;
        headerRow.appendChild(th);
    });
    
    // 添加操作列
    const actionTh = document.createElement('th');
    actionTh.textContent = '操作';
    actionTh.style.width = actionColumnWidth + 'px';
    headerRow.appendChild(actionTh);
    
    // 重置表格宽度设置
    const table = document.getElementById('dataTable');
    table.style.width = '';
    table.style.minWidth = '';
    
    tableHeader.innerHTML = '';
    tableHeader.appendChild(headerRow);
    
    // 构建表体
    tableBody.innerHTML = '';
    
    filteredData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        // 第一列添加行插入器
        let firstTd = true;
        currentSchema.forEach((col, colIndex) => {
            const td = document.createElement('td');
            td.className = 'editable';
            td.setAttribute('data-field', col);
            td.setAttribute('data-id', row.id);
            td.textContent = row[col] || '';
            
            // 在第一列添加行插入器和行调整器
            if (firstTd) {
                td.innerHTML += `
                    <div class="row-resizer" data-row="${rowIndex}"></div>
                    <div class="row-inserter" onclick="insertRowAfter(${rowIndex})" title="在此行后插入新行"></div>
                `;
                firstTd = false;
            }
            
            tr.appendChild(td);
        });
        
        // 操作列
        const actionTd = document.createElement('td');
        actionTd.style.width = actionColumnWidth + 'px';
        actionTd.innerHTML = `
            <button class="btn btn-danger" onclick="deleteRecord(${row.id})">删除</button>
        `;
        tr.appendChild(actionTd);
        
        tableBody.appendChild(tr);
    });
    
    // 添加编辑功能
    addEditFunctionality();
    
    // 添加行高调整功能
    addRowResizeFunctionality();
    
    // 添加列宽调整功能
    addColumnResizeFunctionality();
    
    // 确保表头固定功能正常工作
    ensureHeaderSticky();
    
    // 添加表格独立滚动功能
    addTableScrollControl();
}

// 添加编辑功能
function addEditFunctionality() {
    const editableCells = document.querySelectorAll('.editable');
    
    editableCells.forEach(cell => {
        cell.addEventListener('click', function() {
            if (this.classList.contains('editing')) return;
            
            const currentValue = this.textContent;
            const input = document.createElement('input');
            input.type = 'text';
            input.value = currentValue;
            input.className = 'editing';
            input.style.width = '100%';
            input.style.padding = '8px';
            input.style.border = '2px solid #667eea';
            input.style.borderRadius = '6px';
            input.style.fontSize = '14px';
            input.style.background = 'white';
            
            this.innerHTML = '';
            this.appendChild(input);
            input.focus();
            input.select();
            
            const saveEdit = async () => {
                let newValue = input.value;
                const field = this.dataset.field;
                const id = this.dataset.id;
                
                // 数据验证和格式化
                newValue = validateAndFormatCellValue(newValue, field);
                
                if (newValue !== currentValue) {
                    addConsoleLog(`更新记录 ${id} 的 ${field} 字段`, 'system');
                    
                    try {
                        const response = await fetch('/update', {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                            },
                            body: JSON.stringify({
                                id: parseInt(id),
                                [field]: newValue
                            })
                        });
                        
                        const result = await response.json();
                        
                        if (result.success) {
                            this.textContent = newValue;
                            addConsoleLog(`字段 ${field} 更新成功`, 'system');
                            
                            // 更新本地数据
                            const record = currentData.find(r => r.id == id);
                            if (record) {
                                record[field] = newValue;
                            }
                            const filteredRecord = filteredData.find(r => r.id == id);
                            if (filteredRecord) {
                                filteredRecord[field] = newValue;
                            }
                        } else {
                            this.textContent = currentValue;
                            addConsoleLog(`更新失败: ${result.message}`, 'error');
                        }
                    } catch (error) {
                        this.textContent = currentValue;
                        addConsoleLog(`更新时发生错误: ${error.message}`, 'error');
                    }
                } else {
                    this.textContent = currentValue;
                }
                
                this.classList.remove('editing');
            };
            
            input.addEventListener('blur', saveEdit);
            input.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    saveEdit();
                }
                if (e.key === 'Escape') {
                    cell.textContent = currentValue;
                    cell.classList.remove('editing');
                }
            });
        });
    });
}

// 添加行高调整功能
function addRowResizeFunctionality() {
    const rowResizers = document.querySelectorAll('.row-resizer');
    
    rowResizers.forEach(resizer => {
        let isResizing = false;
        let startY = 0;
        let startHeight = 0;
        let targetRow = null;
        
        resizer.addEventListener('mousedown', function(e) {
            e.preventDefault();
            isResizing = true;
            startY = e.clientY;
            
            // 找到目标行
            targetRow = this.closest('tr');
            if (targetRow) {
                startHeight = targetRow.offsetHeight;
                document.body.style.cursor = 'row-resize';
                document.body.style.userSelect = 'none';
            }
        });
        
        document.addEventListener('mousemove', function(e) {
            if (!isResizing || !targetRow) return;
            
            e.preventDefault();
            const deltaY = e.clientY - startY;
            const newHeight = Math.max(36, startHeight + deltaY); // 最小高度36px
            
            // 设置行高
            targetRow.style.height = newHeight + 'px';
            targetRow.querySelectorAll('td').forEach(td => {
                td.style.height = newHeight + 'px';
            });
        });
        
        document.addEventListener('mouseup', function() {
            if (isResizing) {
                isResizing = false;
                targetRow = null;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                
                addConsoleLog('行高调整完成', 'system');
            }
        });
    });
}

// 添加列宽调整功能
function addColumnResizeFunctionality() {
    const colResizers = document.querySelectorAll('.col-resizer');
    
    colResizers.forEach(resizer => {
        let isResizing = false;
        let startX = 0;
        let startWidth = 0;
        let targetColumn = null;
        let columnIndex = 0;
        
        resizer.addEventListener('mousedown', function(e) {
            e.preventDefault();
            e.stopPropagation();
            isResizing = true;
            startX = e.clientX;
            
            // 找到目标列
            targetColumn = this.closest('th');
            
            if (targetColumn) {
                startWidth = targetColumn.offsetWidth;
                columnIndex = parseInt(this.dataset.column);
                document.body.style.cursor = 'col-resize';
                document.body.style.userSelect = 'none';
                
                addConsoleLog(`开始调整第${columnIndex + 1}列宽度`, 'system');
            }
        });
        
        document.addEventListener('mousemove', function(e) {
            if (!isResizing || !targetColumn) return;
            
            e.preventDefault();
            const deltaX = e.clientX - startX;
            const newWidth = Math.max(80, startWidth + deltaX); // 最小宽度80px
            
            // 只调整被拖动的列宽，设置最小和最大宽度
            targetColumn.style.width = newWidth + 'px';
            targetColumn.style.minWidth = newWidth + 'px';
            targetColumn.style.maxWidth = newWidth + 'px';
            
            // 同时设置对应的数据列宽度
            const table = targetColumn.closest('table');
            if (table) {
                const rows = table.querySelectorAll('tbody tr');
                rows.forEach(row => {
                    const cell = row.children[columnIndex];
                    if (cell) {
                        cell.style.width = newWidth + 'px';
                        cell.style.minWidth = newWidth + 'px';
                        cell.style.maxWidth = newWidth + 'px';
                    }
                });
            }
        });
        
        document.addEventListener('mouseup', function() {
            if (isResizing) {
                isResizing = false;
                targetColumn = null;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                
                addConsoleLog('列宽调整完成', 'system');
            }
        });
    });
}

// 确保表头固定功能正常工作
function ensureHeaderSticky() {
    const tableWrapper = document.querySelector('.table-wrapper');
    const tableHeaders = document.querySelectorAll('thead th');
    
    if (!tableWrapper || tableHeaders.length === 0) return;
    
    // 添加滚动监听，确保表头始终可见
    tableWrapper.addEventListener('scroll', function() {
        const scrollTop = this.scrollTop;
        
        // 动态调整表头的top位置，确保始终在视口顶部
        tableHeaders.forEach(header => {
            // 如果sticky失效，使用JavaScript手动固定
            if (scrollTop > 0) {
                header.style.transform = `translateY(${scrollTop}px)`;
                header.style.position = 'relative';
                header.style.zIndex = '101';
            } else {
                header.style.transform = '';
                header.style.position = 'sticky';
                header.style.zIndex = '100';
            }
        });
    });
    
    // 确保表头在resize时保持正确位置
    window.addEventListener('resize', function() {
        tableHeaders.forEach(header => {
            header.style.top = '0';
        });
    });
    
    addConsoleLog('表头固定功能已启用', 'system');
}

// 搜索数据
function searchData() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
    
    if (!searchTerm) {
        filteredData = [...currentData];
        addConsoleLog('已清除搜索条件', 'system');
    } else {
        filteredData = currentData.filter(row => {
            // 搜索所有字段，包括列名和数据值
            return searchInRow(row, searchTerm);
        });
        
        addConsoleLog(`搜索 "${searchTerm}" 找到 ${filteredData.length} 条记录`, 'system');
    }
    
    renderTable();
    highlightSearchResults(searchTerm);
}

// 智能搜索函数
function searchInRow(row, searchTerm) {
    // 1. 搜索列名（表头）
    for (let columnName of currentSchema) {
        if (columnName.toLowerCase().includes(searchTerm)) {
            // 如果列名包含搜索词，且该列有值，则匹配
            const value = row[columnName];
            if (value != null && String(value).trim() !== '') {
                return true;
            }
        }
    }
    
    // 2. 搜索数据值
    for (let [key, value] of Object.entries(row)) {
        if (value == null || value === undefined) continue;
        
        const stringValue = String(value).toLowerCase();
        
        // 精确匹配
        if (stringValue === searchTerm) {
            return true;
        }
        
        // 包含匹配
        if (stringValue.includes(searchTerm)) {
            return true;
        }
        
        // 模糊匹配（去除特殊字符后匹配）
        const normalizedValue = stringValue.replace(/[\s\-_]/g, '');
        const normalizedSearch = searchTerm.replace(/[\s\-_]/g, '');
        if (normalizedValue.includes(normalizedSearch)) {
            return true;
        }
        
        // 拼音首字母匹配（简单实现）
        if (matchPinyinInitials(stringValue, searchTerm)) {
            return true;
        }
    }
    
    return false;
}

// 简单的拼音首字母匹配
function matchPinyinInitials(text, search) {
    // 这是一个简化版本，实际应用中可以集成更完整的拼音库
    const pinyinMap = {
        '申': 's', '请': 'q', '单': 'd', '位': 'w',
        '项': 'x', '目': 'm', '活': 'h', '动': 'd',
        '概': 'g', '算': 's', '金': 'j', '额': 'e',
        '计': 'j', '划': 'h', '开': 'k', '始': 's',
        '时': 's', '间': 'j', '完': 'w', '成': 'c',
        '管': 'g', '理': 'l', '部': 'b', '门': 'm'
    };
    
    let initials = '';
    for (let char of text) {
        if (pinyinMap[char]) {
            initials += pinyinMap[char];
        }
    }
    
    return initials.includes(search.toLowerCase());
}

// 高亮搜索结果
function highlightSearchResults(searchTerm) {
    if (!searchTerm) {
        // 清除所有高亮
        clearHighlights();
        return;
    }
    
    // 高亮表头
    highlightTableHeaders(searchTerm);
    
    // 高亮表格内容
    const tableBody = document.getElementById('dataBody');
    if (!tableBody) return;
    
    const cells = tableBody.querySelectorAll('td.editable');
    
    cells.forEach(cell => {
        const originalText = cell.textContent;
        if (originalText) {
            // 创建高亮版本的文本
            const highlightedText = highlightText(originalText, searchTerm);
            if (highlightedText !== originalText) {
                cell.innerHTML = highlightedText;
            }
        }
    });
}

// 高亮表头
function highlightTableHeaders(searchTerm) {
    const headers = document.querySelectorAll('th .column-title');
    
    headers.forEach(header => {
        const originalText = header.textContent;
        if (originalText) {
            // 创建高亮版本的文本
            const highlightedText = highlightText(originalText, searchTerm);
            if (highlightedText !== originalText) {
                header.innerHTML = highlightedText;
                // 给表头添加特殊的高亮样式
                header.closest('th').style.backgroundColor = 'rgba(255, 235, 59, 0.2)';
            }
        }
    });
}

// 文本高亮函数
function highlightText(text, searchTerm) {
    if (!searchTerm || !text) return text;
    
    const regex = new RegExp(`(${escapeRegExp(searchTerm)})`, 'gi');
    return text.replace(regex, '<mark style="background-color: #ffeb3b; padding: 0 2px; border-radius: 2px;">$1</mark>');
}

// 转义正则表达式特殊字符
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// 清除搜索高亮
function clearHighlights() {
    // 清除表格内容高亮
    const tableBody = document.getElementById('dataBody');
    if (tableBody) {
        const highlightedCells = tableBody.querySelectorAll('td.editable mark');
        highlightedCells.forEach(mark => {
            const parent = mark.parentNode;
            parent.replaceChild(document.createTextNode(mark.textContent), mark);
            parent.normalize(); // 合并相邻的文本节点
        });
    }
    
    // 清除表头高亮
    const headers = document.querySelectorAll('th .column-title');
    headers.forEach(header => {
        const marks = header.querySelectorAll('mark');
        marks.forEach(mark => {
            const parent = mark.parentNode;
            parent.replaceChild(document.createTextNode(mark.textContent), mark);
            parent.normalize();
        });
        
        // 重置表头背景色
        const th = header.closest('th');
        if (th) {
            th.style.backgroundColor = '';
        }
    });
}

// 清除搜索
function clearSearch() {
    document.getElementById('searchInput').value = '';
    filteredData = [...currentData];
    clearHighlights(); // 清除高亮
    renderTable();
    addConsoleLog('已清除搜索条件', 'system');
}

// 重置搜索（新的重置按钮功能）
function resetSearch() {
    const searchInput = document.getElementById('searchInput');
    searchInput.value = '';
    searchInput.focus(); // 聚焦到搜索框
    filteredData = [...currentData];
    clearHighlights(); // 清除高亮
    renderTable();
    addConsoleLog('已重置搜索条件', 'system');
}

// 刷新数据
function refreshData() {
    addConsoleLog('手动刷新数据...', 'system');
    loadData();
}

// 导出数据
async function exportData() {
    if (currentData.length === 0) {
        addConsoleLog('没有数据可导出', 'warning');
        return;
    }
    
    // 检查是否有有效数据（非空行）
    const validDataCount = currentData.filter(row => {
        return currentSchema.some(col => {
            const value = row[col];
            return value && value.toString().trim() !== '';
        });
    }).length;
    
    if (validDataCount === 0) {
        addConsoleLog('没有有效数据可导出，请先添加一些内容', 'warning');
        return;
    }
    
    addConsoleLog(`开始导出数据，共 ${validDataCount} 条有效记录...`, 'system');
    
    try {
        const response = await fetch('/export');
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            // 从响应头获取文件名，如果没有则使用默认名称
            const contentDisposition = response.headers.get('Content-Disposition');
            // 生成更有规律的文件名
            const now = new Date();
            const dateStr = now.toISOString().slice(0,10).replace(/-/g, '');
            const timeStr = now.toTimeString().slice(0,8).replace(/:/g, '');
            const groupName = currentGroupId ? 
                document.querySelector('#groupSelect option:checked')?.textContent?.trim() || `表格组${currentGroupId}` : 
                '当前表格';
            let filename = `${groupName}_${dateStr}_${timeStr}.xlsx`;
            
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                if (filenameMatch) {
                    filename = filenameMatch[1].replace(/['"]/g, '');
                }
            }
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            addConsoleLog(`数据导出成功：${filename}`, 'system');
            showNotification('导出成功', `当前表格已导出: ${filename}`, 'success');
        } else {
            const result = await response.json();
            addConsoleLog(`导出失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`导出时发生错误: ${error.message}`, 'error');
    }
}



// 导出所有表格分组
async function exportAllGroups() {
    addConsoleLog('开始导出所有表格分组...', 'system');
    
    try {
        const response = await fetch('/export-all-groups');
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            // 从响应头获取文件名，如果没有则使用默认名称
            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = `所有表格_${new Date().toISOString().slice(0,19).replace(/:/g, '-')}.xlsx`;
            
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;=\n]*=(['"]?)([^'"\n]*?)\1/);
                if (filenameMatch) {
                    filename = decodeURIComponent(filenameMatch[2]);
                }
            }
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            addConsoleLog(`✅ 所有表格导出成功: ${filename}`, 'system');
            showNotification('导出成功', `所有表格已导出: ${filename}`, 'success');
        } else {
            const errorData = await response.json();
            addConsoleLog(`导出失败: ${errorData.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`导出失败: ${error.message}`, 'error');
    }
}

// 删除记录
async function deleteRecord(id) {
    if (!confirm('确定要删除这条记录吗？')) return;
    
    addConsoleLog(`删除记录 ${id}...`, 'system');
    
    try {
        const response = await fetch(`/delete/${id}`, {
            method: 'DELETE'
        });
        
        const result = await response.json();
        
        if (result.success) {
            // 从本地数据中移除
            currentData = currentData.filter(r => r.id !== id);
            filteredData = filteredData.filter(r => r.id !== id);
            
            renderTable();
            updateStats({
                total_records: currentData.length,
                source_files: new Set(currentData.map(r => r.source_file).filter(f => f)).size,
                total_columns: currentSchema.length,
                last_update: new Date().toLocaleString()
            });
            
            addConsoleLog('记录删除成功', 'system');
        } else {
            addConsoleLog(`删除失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`删除时发生错误: ${error.message}`, 'error');
    }
}

// 清空所有数据
async function clearAllData() {
    if (!confirm('确定要清空所有数据吗？此操作不可撤销！')) return;
    
    addConsoleLog('清空所有数据...', 'system');
    
    try {
        const response = await fetch('/clear', {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (result.success) {
            currentData = [];
            filteredData = [];
            currentSchema = [];
            
            renderTable();
            updateStats({
                total_records: 0,
                source_files: 0,
                total_columns: 0,
                last_update: new Date().toLocaleString()
            });
            
            addConsoleLog('所有数据已清空', 'system');
        } else {
            addConsoleLog(`清空失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`清空数据时发生错误: ${error.message}`, 'error');
    }
}



// 显示重命名列模态框
function showRenameColumnModal(columnName) {
    const modal = document.getElementById('renameColumnModal');
    const currentNameInput = document.getElementById('currentColumnName');
    const newNameInput = document.getElementById('newColumnNameRename');
    
    currentNameInput.value = columnName;
    newNameInput.value = '';
    modal.classList.add('show');
    newNameInput.focus();
}

// 隐藏重命名列模态框
function hideRenameColumnModal() {
    const modal = document.getElementById('renameColumnModal');
    modal.classList.remove('show');
}





// 插入列（在指定列后插入）
function insertColumnAfter(columnIndex) {
    // 自动生成唯一的列名
    let columnName;
    let counter = 1;
    do {
        columnName = `新列${counter}`;
        counter++;
    } while (currentSchema.includes(columnName));
    
    addConsoleLog(`在第${columnIndex + 1}列后插入新列: ${columnName}`, 'system');
    addNewColumnWithPosition(columnName, columnIndex + 1);
}

// 插入行（在指定行后插入）
async function insertRowAfter(rowIndex) {
    addConsoleLog(`在第${rowIndex + 1}行后插入新行`, 'system');
    
    try {
        // 获取当前行的实际ID，用于确定插入位置
        const currentRow = filteredData[rowIndex];
        const insertAfterId = currentRow ? currentRow.id : null;
        
        const response = await fetch('/add_row', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                insert_after_id: insertAfterId,  // 传递实际的行ID而不是索引
                insert_position: rowIndex + 1    // 传递期望的显示位置
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            addConsoleLog('新行插入成功', 'system');
            
            // 直接在前端插入新行，避免页面跳转
            const newRow = {
                id: result.id,
                source_file: '手动添加',
                created_at: new Date().toLocaleString(),
                updated_at: new Date().toLocaleString()
            };
            
            // 为新行添加所有列的空值
            currentSchema.forEach(col => {
                newRow[col] = '';
            });
            
            // 在指定位置插入新行
            currentData.splice(rowIndex + 1, 0, newRow);
            filteredData.splice(rowIndex + 1, 0, newRow);
            
            // 只重新渲染表格，不重新加载数据
            renderTable();
            
            // 更新统计信息
            updateStats({
                total_records: currentData.length,
                source_files: new Set(currentData.map(r => r.source_file).filter(f => f)).size,
                total_columns: currentSchema.length,
                last_update: new Date().toLocaleString()
            });
        } else {
            addConsoleLog(`插入行失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`插入行时发生错误: ${error.message}`, 'error');
    }
}

// 带位置的添加列功能，支持指定位置插入
async function addNewColumnWithPosition(columnName, position) {
    try {
        const response = await fetch('/column/add', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                column_name: columnName,
                insert_position: position  // 传递位置参数
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            addConsoleLog(`列 "${columnName}" 在位置 ${position} 添加成功`, 'system');
            
            // 直接在前端更新表格结构，避免页面跳转
            // 在指定位置插入新列名
            currentSchema.splice(position, 0, columnName);
            
            // 为所有现有数据行添加新列的空值
            currentData.forEach(row => {
                row[columnName] = '';
            });
            filteredData.forEach(row => {
                row[columnName] = '';
            });
            
            // 只重新渲染表格，不重新加载数据
            renderTable();
            
            // 更新统计信息
            updateStats({
                total_records: currentData.length,
                source_files: new Set(currentData.map(r => r.source_file).filter(f => f)).size,
                total_columns: currentSchema.length,
                last_update: new Date().toLocaleString()
            });
        } else {
            addConsoleLog(`添加列失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`添加列时发生错误: ${error.message}`, 'error');
    }
}

// 确认重命名列
async function renameColumnConfirm() {
    const oldName = document.getElementById('currentColumnName').value.trim();
    const newName = document.getElementById('newColumnNameRename').value.trim();
    
    if (!newName) {
        addConsoleLog('新列名不能为空', 'warning');
        return;
    }
    
    if (oldName === newName) {
        addConsoleLog('新列名与原列名相同', 'warning');
        return;
    }
    
    addConsoleLog(`重命名列: ${oldName} -> ${newName}`, 'system');
    
    try {
        const response = await fetch('/column/rename', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                old_name: oldName,
                new_name: newName
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            addConsoleLog(`列 "${oldName}" 重命名为 "${newName}" 成功`, 'system');
            hideRenameColumnModal();
            await loadData(); // 重新加载数据以更新表格结构
        } else {
            addConsoleLog(`重命名列失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`重命名列时发生错误: ${error.message}`, 'error');
    }
}

// 点击模态框外部关闭
document.addEventListener('click', function(e) {
    const renameColumnModal = document.getElementById('renameColumnModal');
    const renameTableModal = document.getElementById('renameTableModal');
    
    if (e.target === renameColumnModal) {
        hideRenameColumnModal();
    }
    if (e.target === renameTableModal) {
        hideRenameTableModal();
    }
});

// 回车键操作
document.addEventListener('keypress', function(e) {
    const renameColumnModal = document.getElementById('renameColumnModal');
    const renameTableModal = document.getElementById('renameTableModal');
    
    if (renameColumnModal.classList.contains('show') && e.key === 'Enter') {
        renameColumnConfirm();
    }
    if (renameTableModal.classList.contains('show') && e.key === 'Enter') {
        renameTableConfirm();
    }
});

// 删除列确认
async function deleteColumnConfirm(columnName) {
    if (!confirm(`确定要删除列 "${columnName}" 吗？此操作将删除该列的所有数据！`)) return;
    
    addConsoleLog(`删除列: ${columnName}`, 'system');
    
    try {
        const response = await fetch('/column/delete', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                column_name: columnName
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            addConsoleLog(`列 "${columnName}" 删除成功`, 'system');
            
            // 直接在前端删除列，避免页面跳转
            const columnIndex = currentSchema.indexOf(columnName);
            if (columnIndex > -1) {
                // 从schema中移除列
                currentSchema.splice(columnIndex, 1);
                
                // 从所有数据行中删除该列
                currentData.forEach(row => {
                    delete row[columnName];
                });
                filteredData.forEach(row => {
                    delete row[columnName];
                });
                
                // 重新渲染表格
                renderTable();
                
                // 更新统计信息
                updateStats({
                    total_records: currentData.length,
                    source_files: new Set(currentData.map(r => r.source_file).filter(f => f)).size,
                    total_columns: currentSchema.length,
                    last_update: new Date().toLocaleString()
                });
            }
        } else {
            addConsoleLog(`删除列失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`删除列时发生错误: ${error.message}`, 'error');
    }
}



// ESC键关闭模态框
document.addEventListener('keydown', function(e) {
    const renameColumnModal = document.getElementById('renameColumnModal');
    const renameTableModal = document.getElementById('renameTableModal');
    
    if (e.key === 'Escape') {
        if (renameColumnModal.classList.contains('show')) {
            hideRenameColumnModal();
        }
        if (renameTableModal.classList.contains('show')) {
            hideRenameTableModal();
        }
    }
});

// 加载表格列表
async function loadTableList() {
    addConsoleLog('正在加载表格列表...', 'system');
    
    try {
        const response = await fetch('/table-groups');
        const result = await response.json();
        
        if (result.success) {
            tableGroups = result.groups;
            updateTableSelector();
            
            if (tableGroups.length > 0) {
                // 如果有指定的分组ID，选择它；否则选择第一个表格
                if (currentGroupId && tableGroups.find(g => g.id === currentGroupId)) {
                    document.getElementById('tableSelector').value = currentGroupId;
                    loadGroupData(currentGroupId);
                } else {
                    currentGroupId = tableGroups[0].id;
                    document.getElementById('tableSelector').value = currentGroupId;
                    loadGroupData(currentGroupId);
                }
            } else {
                // 没有表格时显示空状态
                showEmptyState();
                updateStats({
                    total_records: 0,
                    source_files: 0,
                    total_columns: 0,
                    last_update: '--'
                });
                // 确保隐藏工具栏按钮
                hideToolbarButtons();
            }
            
            addConsoleLog(`发现 ${tableGroups.length} 张表格`, 'system');
        } else {
            addConsoleLog(`加载表格列表失败: ${result.message}`, 'error');
            showEmptyState();
            hideToolbarButtons();
        }
    } catch (error) {
        addConsoleLog(`加载表格列表时发生错误: ${error.message}`, 'error');
        showEmptyState();
        hideToolbarButtons();
    }
}

// 更新表格选择器
function updateTableSelector() {
    const selector = document.getElementById('tableSelector');
    const editBtn = document.getElementById('editTableNameBtn');
    
    selector.innerHTML = '';
    
    // 清除单文件查看模式的标记和事件监听器
    selector.removeAttribute('data-single-file');
    selector.removeAttribute('data-filename');
    selector.onchange = null; // 清除之前设置的onchange
    
    if (tableGroups.length === 0) {
        selector.innerHTML = '<option value="">暂无表格</option>';
        // 隐藏编辑按钮
        if (editBtn) {
            editBtn.style.display = 'none';
        }
        return;
    }
    
    // 显示编辑按钮
    if (editBtn) {
        editBtn.style.display = 'inline-flex';
    }
    
    tableGroups.forEach((group, index) => {
        const option = document.createElement('option');
        option.value = group.id;
        // 简化显示名称，去掉不必要的前缀
        let displayName = group.group_name;
        if (displayName.startsWith('表格组_')) {
            displayName = displayName.replace('表格组_', '表格');
        }
        // 去掉时间戳等后缀，保留主要名称
        displayName = displayName.replace(/_\d{8}_\d{6}$/, '');
        displayName += ` (${group.record_count}条)`;
        option.textContent = displayName;
        selector.appendChild(option);
    });
    
    // 重新初始化选择器事件
    initializeTableSelector();
}

// 更新表格选择器显示单个文件信息
function updateTableSelectorForSingleFile(filename, recordCount) {
    const selector = document.getElementById('tableSelector');
    const editBtn = document.getElementById('editTableNameBtn');
    
    if (!selector) return;
    
    // 清空现有选项
    selector.innerHTML = '';
    
    // 添加返回总表的选项
    const backOption = document.createElement('option');
    backOption.value = 'back-to-groups';
    backOption.textContent = '← 返回总表';
    selector.appendChild(backOption);
    
    // 创建显示当前文件的选项
    const currentOption = document.createElement('option');
    currentOption.value = 'single-file-view';
    currentOption.textContent = `${filename} (${recordCount}条)`;
    currentOption.selected = true;
    selector.appendChild(currentOption);
    
    // 隐藏编辑按钮，因为查看单个文件时不需要重命名功能
    if (editBtn) {
        editBtn.style.display = 'none';
    }
    
    // 添加数据属性标记当前是单文件查看模式
    selector.setAttribute('data-single-file', 'true');
    selector.setAttribute('data-filename', filename);
    
    // 添加选择器变化监听
    selector.onchange = function() {
        if (this.value === 'back-to-groups') {
            backToTableGroups();
        }
    };
    
    addConsoleLog(`📋 表格标题已更新为: ${filename}`, 'system');
}

// 返回表格组模式
async function backToTableGroups() {
    try {
        addConsoleLog('🔄 正在返回总表...', 'system');
        showNotification('正在加载总表...', 'info');
        
        // 重新加载表格组数据
        await loadTableList();
        
        // 如果有表格组，选择第一个并加载其数据
        if (tableGroups.length > 0) {
            currentGroupId = tableGroups[0].id;
            const selector = document.getElementById('tableSelector');
            if (selector) {
                selector.value = currentGroupId;
            }
            await loadGroupData(currentGroupId);
        }
        
        addConsoleLog(`✅ 已返回总表，当前显示: ${tableGroups.length > 0 ? tableGroups[0].group_name : '无表格'}`, 'system');
        showNotification('已返回总表', 'success');
        
    } catch (error) {
        console.error('返回总表失败:', error);
        addConsoleLog('❌ 返回总表失败: ' + error.message, 'error');
        showNotification('返回总表失败: ' + error.message, 'error');
    }
}

// 加载指定表格的数据
async function loadGroupData(groupId) {
    addConsoleLog(`正在加载表格数据...`, 'system');
    
    const loadingState = document.getElementById('loadingState');
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    
    loadingState.style.display = 'flex';
    dataTable.style.display = 'none';
    emptyState.style.display = 'none';
    
    try {
        const response = await fetch(`/table-groups/${groupId}/data`);
        const result = await response.json();
        
        if (result.success) {
            currentData = result.data;
            filteredData = [...currentData];
            currentSchema = result.schema || [];
            
            updateStats(result.stats);
            renderTable();
            
            addConsoleLog(`表格加载完成，共 ${currentData.length} 条记录，${currentSchema.length} 个列`, 'system');
        } else {
            addConsoleLog(`加载表格数据失败: ${result.message}`, 'error');
            showEmptyState();
        }
    } catch (error) {
        addConsoleLog(`加载表格数据时发生错误: ${error.message}`, 'error');
        showEmptyState();
    } finally {
        loadingState.style.display = 'none';
    }
}

// 清空所有数据
async function clearAllData() {
    if (!confirm('⚠️ 确定要清空所有表格数据吗？\n\n此操作将删除所有上传的表格和记录，且不可撤销！')) {
        return;
    }
    
    addConsoleLog('正在清空所有数据...', 'system');
    
    try {
        const response = await fetch('/clear-all', {
            method: 'POST'
        });
        
        const result = await response.json();
        
        if (result.success) {
            // 重置所有状态
            currentData = [];
            filteredData = [];
            currentSchema = [];
            tableGroups = [];
            currentGroupId = null;
            
            // 更新界面
            showEmptyState();
            updateStats({
                total_records: 0,
                source_files: 0,
                total_columns: 0,
                last_update: '--'
            });
            
            // 更新选择器
            const selector = document.getElementById('tableSelector');
            selector.innerHTML = '<option value="">暂无表格</option>';
            
            addConsoleLog('✅ 所有数据已清空', 'system');
            showNotification('清空成功', '所有表格数据已清空', 'info');
        } else {
            addConsoleLog(`清空失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`清空数据时发生错误: ${error.message}`, 'error');
    }
}

// 显示重命名表格模态框
function showRenameTableModal() {
    if (!currentGroupId) {
        addConsoleLog('请先选择一个表格', 'warning');
        return;
    }
    
    const currentGroup = tableGroups.find(g => g.id === currentGroupId);
    if (!currentGroup) {
        addConsoleLog('找不到当前表格信息', 'error');
        return;
    }
    
    const modal = document.getElementById('renameTableModal');
    const currentNameInput = document.getElementById('currentTableName');
    const newNameInput = document.getElementById('newTableName');
    
    // 提取并优化表格名称显示
    let currentName = currentGroup.group_name;
    if (currentName.startsWith('表格组_')) {
        currentName = currentName.replace('表格组_', '表格');
    }
    // 去掉时间戳等后缀，保留主要名称
    currentName = currentName.replace(/_\d{8}_\d{6}$/, '');
    // 去除记录数信息
    currentName = currentName.replace(/\s*\(\d+条\)$/, '');
    
    currentNameInput.value = currentName;
    newNameInput.value = '';
    modal.classList.add('show');
    newNameInput.focus();
}

// 隐藏重命名表格模态框
function hideRenameTableModal() {
    const modal = document.getElementById('renameTableModal');
    modal.classList.remove('show');
}

// 确认重命名表格
async function renameTableConfirm() {
    const newName = document.getElementById('newTableName').value.trim();
    const currentName = document.getElementById('currentTableName').value.trim();
    
    if (!newName) {
        addConsoleLog('新表格名不能为空', 'warning');
        return;
    }
    
    if (newName === currentName) {
        addConsoleLog('新表格名与原表格名相同', 'warning');
        return;
    }
    
    addConsoleLog(`重命名表格: ${currentName} -> ${newName}`, 'system');
    
    try {
        const response = await fetch('/table-groups/rename', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                group_id: currentGroupId,
                new_name: newName
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            addConsoleLog(`表格 "${currentName}" 重命名为 "${newName}" 成功`, 'system');
            hideRenameTableModal();
            
            // 刷新表格列表
            await loadTableList();
            
            // 保持当前选中的表格
            document.getElementById('tableSelector').value = currentGroupId;
            
            showNotification('重命名成功', `表格已重命名为："${newName}"`, 'success');
        } else {
            addConsoleLog(`重命名表格失败: ${result.message}`, 'error');
        }
    } catch (error) {
        addConsoleLog(`重命名表格时发生错误: ${error.message}`, 'error');
    }
}

// 兼容原有的loadData函数 - 重新加载表格列表
async function loadData() {
    if (currentGroupId) {
        await loadGroupData(currentGroupId);
    } else {
        await loadTableList();
    }
}

// 刷新数据 - 重新加载表格列表
function refreshData() {
    addConsoleLog('手动刷新数据...', 'system');
    loadTableList();
}

// 列排序功能
function sortColumn(columnName, direction) {
    addConsoleLog(`对列"${columnName}"进行${direction === 'asc' ? '升序' : '降序'}排序`, 'system');
    
    filteredData.sort((a, b) => {
        let valueA = a[columnName];
        let valueB = b[columnName];
        
        // 处理空值
        if (valueA == null && valueB == null) return 0;
        if (valueA == null) return direction === 'asc' ? -1 : 1;
        if (valueB == null) return direction === 'asc' ? 1 : -1;
        
        // 尝试数字比较
        const numA = parseFloat(valueA);
        const numB = parseFloat(valueB);
        
        if (!isNaN(numA) && !isNaN(numB)) {
            return direction === 'asc' ? numA - numB : numB - numA;
        }
        
        // 字符串比较
        const strA = String(valueA).toLowerCase();
        const strB = String(valueB).toLowerCase();
        
        if (direction === 'asc') {
            return strA.localeCompare(strB);
        } else {
            return strB.localeCompare(strA);
        }
    });
    
    renderTable();
}



// ====== 数据验证和格式化 ======

// 验证和格式化单元格值
function validateAndFormatCellValue(value, fieldName) {
    if (!value || value.trim() === '') {
        return value;
    }
    
    const trimmedValue = value.trim();
    const fieldLower = fieldName.toLowerCase();
    
    // 数字格式化 (金额、薪资、投资等)
    if (fieldLower.includes('金额') || fieldLower.includes('薪资') || 
        fieldLower.includes('投资') || fieldLower.includes('元') || 
        fieldLower.includes('万元') || fieldLower.includes('资金')) {
        
        const numMatch = trimmedValue.match(/[\d,]+\.?\d*/);
        if (numMatch) {
            let num = parseFloat(numMatch[0].replace(/,/g, ''));
            if (!isNaN(num)) {
                // 保留原有的单位
                if (trimmedValue.includes('万')) {
                    return num.toLocaleString() + '万元';
                } else if (trimmedValue.includes('元')) {
                    return num.toLocaleString() + '元';
                } else {
                    return num.toLocaleString();
                }
            }
        }
    }
    
    // 日期格式化
    if (fieldLower.includes('时间') || fieldLower.includes('日期') || 
        fieldLower.includes('开始') || fieldLower.includes('完成')) {
        
        // 检测日期格式：YYYY-MM-DD
        const dateMatch = trimmedValue.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (dateMatch) {
            const year = dateMatch[1];
            const month = dateMatch[2].padStart(2, '0');
            const day = dateMatch[3].padStart(2, '0');
            return `${year}-${month}-${day}`;
        }
    }
    
    // 部门格式化
    if (fieldLower.includes('部门') || fieldLower.includes('单位')) {
        if (trimmedValue && !trimmedValue.endsWith('部') && !trimmedValue.endsWith('门') && 
            !trimmedValue.endsWith('司') && !trimmedValue.endsWith('院')) {
            // 自动添加"部门"后缀的逻辑可以根据需要调整
        }
    }
    
    // 项目序号格式化
    if (fieldLower.includes('序号') || fieldLower.includes('编号') || fieldLower === '项目序号') {
        const upperValue = trimmedValue.toUpperCase();
        // 保持项目编号的格式如 CW001, FW001 等
        if (/^[A-Z]{2}\d{3}$/.test(upperValue)) {
            return upperValue;
        }
    }
    
    // 状态标准化
    if (fieldLower.includes('状态') || fieldLower.includes('情况')) {
        const statusMap = {
            '进行中': '同意执行',
            '执行中': '同意执行',
            '待审批': '待审批',
            '审批中': '待审批',
            '已完成': '同意执行',
            '完成': '同意执行',
            '暂停': '待审批'
        };
        
        return statusMap[trimmedValue] || trimmedValue;
    }
    
    return trimmedValue;
}

// 添加单元格验证状态指示
function addValidationIndicator(cell, isValid, message = '') {
    // 移除旧的验证指示器
    const oldIndicator = cell.querySelector('.validation-indicator');
    if (oldIndicator) {
        oldIndicator.remove();
    }
    
    if (!isValid) {
        const indicator = document.createElement('span');
        indicator.className = 'validation-indicator';
        indicator.style.cssText = `
            position: absolute;
            top: 2px;
            right: 2px;
            width: 6px;
            height: 6px;
            background: #ef4444;
            border-radius: 50%;
            cursor: help;
        `;
        indicator.title = message || '数据格式不正确';
        cell.style.position = 'relative';
        cell.appendChild(indicator);
    }
}

// 添加表格独立滚动控制功能
function addTableScrollControl() {
    const tableWrapper = document.querySelector('.table-wrapper');
    if (!tableWrapper) return;
    
    // 增强的滚动控制（优化Mac触摸板支持）
    tableWrapper.addEventListener('wheel', function(e) {
        e.stopPropagation(); // 阻止事件冒泡到页面
        
        const scrollTop = this.scrollTop;
        const scrollLeft = this.scrollLeft;
        const scrollHeight = this.scrollHeight;
        const scrollWidth = this.scrollWidth;
        const clientHeight = this.clientHeight;
        const clientWidth = this.clientWidth;
        let deltaY = e.deltaY;
        let deltaX = e.deltaX;
        
        // Mac触摸板优化：增强水平滚动检测
        const isMacTouchpad = /Mac|iOS/.test(navigator.platform) && Math.abs(e.deltaY) > 0 && Math.abs(e.deltaX) > 0.1;
        
        // 处理Shift+滚轮的水平滚动（鼠标滚轮用户）
        if (e.shiftKey && Math.abs(deltaY) > 0) {
            deltaX = deltaY; // 将垂直滚动转换为水平滚动
            deltaY = 0;      // 取消垂直滚动
        }
        
        // Mac触摸板特殊处理：降低水平滚动阈值，增加敏感度
        if (isMacTouchpad && Math.abs(deltaX) > 0.1) {
            deltaX = deltaX * 1.5; // 增加触摸板水平滚动的敏感度
        }
        
        // 检查垂直滚动边界
        const atTop = scrollTop <= 1;
        const atBottom = Math.abs(scrollTop + clientHeight - scrollHeight) <= 1;
        
        // 检查水平滚动边界
        const atLeft = scrollLeft <= 1;
        const atRight = Math.abs(scrollLeft + clientWidth - scrollWidth) <= 1;
        
        let shouldPreventDefault = false;
        let scrolled = false;
        
        // 处理垂直滚动
        if (Math.abs(deltaY) > 0.1) {
            if (!(atTop && deltaY < 0) && !(atBottom && deltaY > 0)) {
                // 在边界内，允许表格滚动
                const newScrollTop = Math.max(0, Math.min(this.scrollTop + deltaY, scrollHeight - clientHeight));
                this.scrollTop = newScrollTop;
                shouldPreventDefault = true;
                scrolled = true;
            }
        }
        
        // 处理水平滚动（降低阈值以支持触摸板）
        if (Math.abs(deltaX) > 0.1) {
            if (!(atLeft && deltaX < 0) && !(atRight && deltaX > 0)) {
                // 在边界内，允许表格滚动
                const newScrollLeft = Math.max(0, Math.min(this.scrollLeft + deltaX, scrollWidth - clientWidth));
                this.scrollLeft = newScrollLeft;
                shouldPreventDefault = true;
                scrolled = true;
                
                // 水平滚动专用反馈
                console.log(`水平滚动: deltaX=${deltaX.toFixed(2)}, scrollLeft=${newScrollLeft}`);
            }
        }
        
        // 如果处理了任一方向的滚动，阻止默认行为
        if (shouldPreventDefault) {
            e.preventDefault();
        }
        
        // 添加滚动反馈
        if (scrolled) {
            // 短暂改变边框颜色提供视觉反馈
            const originalBoxShadow = this.style.boxShadow;
            this.style.boxShadow = '0 0 0 2px rgba(59, 130, 246, 0.3)';
            setTimeout(() => {
                this.style.boxShadow = originalBoxShadow;
            }, 150);
        }
    }, { passive: false });
    
    // 鼠标进入表格区域时添加视觉提示
    tableWrapper.addEventListener('mouseenter', function() {
        this.style.boxShadow = '0 0 0 2px rgba(59, 130, 246, 0.1)';
        document.body.style.overflow = 'hidden'; // 临时禁用页面滚动
    });
    
    // 鼠标离开表格区域时移除视觉提示
    tableWrapper.addEventListener('mouseleave', function() {
        this.style.boxShadow = '';
        document.body.style.overflow = ''; // 恢复页面滚动
    });
    
    // 添加键盘滚动支持（当表格获得焦点时）
    tableWrapper.addEventListener('keydown', function(e) {
        const scrollAmount = 50;
        const horizontalScrollAmount = 100;
        
        switch(e.key) {
            case 'ArrowUp':
                e.preventDefault();
                this.scrollTop -= scrollAmount;
                break;
            case 'ArrowDown':
                e.preventDefault();
                this.scrollTop += scrollAmount;
                break;
            case 'ArrowLeft':
                e.preventDefault();
                this.scrollLeft -= horizontalScrollAmount;
                break;
            case 'ArrowRight':
                e.preventDefault();
                this.scrollLeft += horizontalScrollAmount;
                break;
            case 'PageUp':
                e.preventDefault();
                this.scrollTop -= this.clientHeight;
                break;
            case 'PageDown':
                e.preventDefault();
                this.scrollTop += this.clientHeight;
                break;
            case 'Home':
                if (e.ctrlKey) {
                    e.preventDefault();
                    this.scrollTop = 0;
                    this.scrollLeft = 0;  // 同时回到左上角
                } else {
                    e.preventDefault();
                    this.scrollLeft = 0;  // 只回到行首
                }
                break;
            case 'End':
                if (e.ctrlKey) {
                    e.preventDefault();
                    this.scrollTop = this.scrollHeight;
                    this.scrollLeft = this.scrollWidth;  // 同时到右下角
                } else {
                    e.preventDefault();
                    this.scrollLeft = this.scrollWidth;  // 只到行尾
                }
                break;
        }
    });
    
    // 使表格容器可获得焦点
    tableWrapper.setAttribute('tabindex', '0');
    tableWrapper.style.outline = 'none';
    
    // 添加触摸屏和触摸板滑动支持
    let touchStartX = 0;
    let touchStartY = 0;
    let touchStartScrollLeft = 0;
    let touchStartScrollTop = 0;
    
    tableWrapper.addEventListener('touchstart', function(e) {
        if (e.touches.length === 1) {
            const touch = e.touches[0];
            touchStartX = touch.clientX;
            touchStartY = touch.clientY;
            touchStartScrollLeft = this.scrollLeft;
            touchStartScrollTop = this.scrollTop;
        }
    }, { passive: true });
    
    tableWrapper.addEventListener('touchmove', function(e) {
        if (e.touches.length === 1) {
            e.preventDefault(); // 阻止页面滚动
            e.stopPropagation();
            
            const touch = e.touches[0];
            const deltaX = touchStartX - touch.clientX;
            const deltaY = touchStartY - touch.clientY;
            
            // 应用滚动，但限制在边界内
            const newScrollLeft = Math.max(0, Math.min(
                this.scrollWidth - this.clientWidth,
                touchStartScrollLeft + deltaX
            ));
            const newScrollTop = Math.max(0, Math.min(
                this.scrollHeight - this.clientHeight,
                touchStartScrollTop + deltaY
            ));
            
            this.scrollLeft = newScrollLeft;
            this.scrollTop = newScrollTop;
        }
    }, { passive: false });
    
    // 添加滚动条样式优化（让滚动条更明显）
    tableWrapper.style.overflowX = 'auto';
    tableWrapper.style.overflowY = 'auto';
    
    addConsoleLog('表格独立滚动功能已启用（支持垂直和水平滚动）', 'system');
}

// ==================== 文件管理功能 ====================

// 分页相关变量
let filesData = [];
let currentPage = 1;
const filesPerPage = 3;  // 每页显示3个文件

// 加载已上传文件列表
async function loadTablesList() {
    const loadingElement = document.getElementById('loadingTables');
    const noTablesMessage = document.getElementById('noTablesMessage');
    const tablesListContainer = document.getElementById('tablesListContainer');
    
    try {
        loadingElement.style.display = 'block';
        noTablesMessage.style.display = 'none';
        tablesListContainer.style.display = 'none';
        
        addConsoleLog('正在加载文件列表...', 'system');
        
        const response = await fetch('/uploaded-files');
        const data = await response.json();
        
        if (data.success && data.files && data.files.length > 0) {
            filesData = data.files;  // 保存所有文件数据
            currentPage = 1;  // 重置到第一页
            renderFilesList();
            tablesListContainer.style.display = 'block';
            noTablesMessage.style.display = 'none';
            addConsoleLog(`成功加载 ${data.files.length} 个文件`, 'system');
        } else {
            filesData = [];
            noTablesMessage.style.display = 'block';
            tablesListContainer.style.display = 'none';
            addConsoleLog('暂无上传的文件', 'system');
        }
    } catch (error) {
        console.error('加载文件列表失败:', error);
        addConsoleLog('加载文件列表失败: ' + error.message, 'error');
        noTablesMessage.style.display = 'block';
        tablesListContainer.style.display = 'none';
    } finally {
        loadingElement.style.display = 'none';
    }
}

// 渲染文件列表
function renderFilesList() {
    const tablesList = document.getElementById('tablesList');
    const paginationContainer = document.getElementById('paginationContainer');
    
    if (!filesData || filesData.length === 0) {
        tablesList.innerHTML = '';
        paginationContainer.style.display = 'none';
        return;
    }
    
    // 计算分页
    const totalFiles = filesData.length;
    const totalPages = Math.ceil(totalFiles / filesPerPage);
    const startIndex = (currentPage - 1) * filesPerPage;
    const endIndex = Math.min(startIndex + filesPerPage, totalFiles);
    const currentFiles = filesData.slice(startIndex, endIndex);
    
    // 渲染当前页的文件
    let filesHtml = currentFiles.map(file => {
        const displayName = file.filename || `文件-${file.id}`;
        const recordCount = file.current_records || 0;
        const uploadTime = file.upload_time ? new Date(file.upload_time).toLocaleDateString() : '--';
        const hasData = file.has_data;
        
        // 如果没有数据了，显示不同的样式
        const itemStyle = hasData ? '' : 'opacity: 0.6; border-color: #f87171;';
        const statusText = hasData ? '' : ' (数据已删除)';
        
        return `
            <div class="table-item" data-filename="${file.filename}" style="${itemStyle}">
                <div class="table-item-header">
                    <div class="table-item-title" onclick="viewFile('${file.filename}')" title="点击查看文件内容">
                        ${displayName}${statusText}
                    </div>
                    <div class="table-item-info">
                        记录: ${recordCount} 条 | 上传: ${uploadTime}
                    </div>
                </div>
                <div class="table-item-actions">
                    <button class="table-action-btn table-action-btn-view" 
                            onclick="viewFile('${file.filename}')" 
                            title="查看" 
                            ${!hasData ? 'disabled style="opacity:0.5;"' : ''}>
                        查看
                    </button>
                    <button class="table-action-btn table-action-btn-export" 
                            onclick="exportFile('${file.filename}')" 
                            title="导出"
                            ${!hasData ? 'disabled style="opacity:0.5;"' : ''}>
                        导出
                    </button>
                    <button class="table-action-btn table-action-btn-delete" 
                            onclick="deleteFile('${file.filename}')" 
                            title="删除">
                        删除
                    </button>
                </div>
            </div>
        `;
    }).join('');
    
    // 在最后一页的末尾添加上传新文件的卡片
    if (currentPage === totalPages) {
        filesHtml += `
            <div class="table-item" style="border: 2px dashed #d1d5db; background: #f9fafb;">
                <div class="table-item-header" style="background: #f9fafb; border-bottom: none;">
                    <div class="table-item-title" style="color: #6b7280; text-align: center;">
                        上传新文件
                    </div>
                </div>
                <div class="table-item-actions" style="padding: 12px 16px;">
                    <button class="table-action-btn table-action-btn-upload" onclick="document.getElementById('fileInput').click()" title="上传新文件">
                        选择文件上传
                    </button>
                </div>
            </div>
        `;
    }
    
    tablesList.innerHTML = filesHtml;
    
    // 更新分页控件
    updatePaginationControls(currentPage, totalPages, totalFiles);
}

// 更新分页控件
function updatePaginationControls(page, totalPages, totalFiles) {
    const paginationContainer = document.getElementById('paginationContainer');
    const prevBtn = document.getElementById('prevPageBtn');
    const nextBtn = document.getElementById('nextPageBtn');
    const pageInfo = document.getElementById('pageInfo');
    
    if (totalPages <= 1) {
        paginationContainer.style.display = 'none';
        return;
    }
    
    paginationContainer.style.display = 'block';
    pageInfo.textContent = `${page}/${totalPages} (共${totalFiles}个文件)`;
    
    // 更新按钮状态
    prevBtn.disabled = page <= 1;
    nextBtn.disabled = page >= totalPages;
    
    if (page <= 1) {
        prevBtn.style.opacity = '0.5';
        prevBtn.style.cursor = 'not-allowed';
    } else {
        prevBtn.style.opacity = '1';
        prevBtn.style.cursor = 'pointer';
    }
    
    if (page >= totalPages) {
        nextBtn.style.opacity = '0.5';
        nextBtn.style.cursor = 'not-allowed';
    } else {
        nextBtn.style.opacity = '1';
        nextBtn.style.cursor = 'pointer';
    }
}

// 切换页码
function changePage(direction) {
    const totalPages = Math.ceil(filesData.length / filesPerPage);
    const newPage = currentPage + direction;
    
    if (newPage >= 1 && newPage <= totalPages) {
        currentPage = newPage;
        renderFilesList();
        addConsoleLog(`切换到第 ${currentPage} 页`, 'system');
    }
}

// 查看文件内容
async function viewFile(filename) {
    if (!filename) {
        showNotification('文件名无效', 'error');
        return;
    }
    
    try {
        addConsoleLog(`正在加载文件 ${filename} 的数据...`, 'system');
        showNotification('正在加载文件数据...', 'info');
        
        const response = await fetch(`/uploaded-files/${encodeURIComponent(filename)}/data`);
        const data = await response.json();
        
        if (data.success) {
            // 更新全局数据
            currentData = data.data || [];
            filteredData = [...currentData];
            currentSchema = data.schema || [];
            
            // 渲染表格
            renderTable();
            
            // 更新左侧表格标题显示当前查看的文件名
            updateTableSelectorForSingleFile(filename, currentData.length);
            
            // 滚动到表格区域
            const tableContainer = document.querySelector('.table-container');
            if (tableContainer) {
                tableContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
            
            addConsoleLog(`文件 ${filename} 加载成功，共 ${currentData.length} 条记录`, 'system');
            showNotification(`文件加载成功，共 ${currentData.length} 条记录`, 'success');
        } else {
            throw new Error(data.message || '加载失败');
        }
        
    } catch (error) {
        console.error('查看文件失败:', error);
        addConsoleLog('查看文件失败: ' + error.message, 'error');
        showNotification('查看文件失败: ' + error.message, 'error');
    }
}

// 导出单个文件
async function exportFile(filename) {
    if (!filename) {
        showNotification('文件名无效', 'error');
        return;
    }
    
    try {
        addConsoleLog(`正在导出文件 ${filename}...`, 'system');
        showNotification('正在导出文件...', 'info');
        
        const response = await fetch(`/uploaded-files/${encodeURIComponent(filename)}/export`);
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            // 从响应头获取文件名
            const contentDisposition = response.headers.get('Content-Disposition');
            let exportFilename = 'export.xlsx';
            if (contentDisposition) {
                const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                if (filenameMatch && filenameMatch[1]) {
                    exportFilename = filenameMatch[1].replace(/['"]/g, '');
                }
            }
            
            a.download = exportFilename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            addConsoleLog(`文件 ${filename} 导出成功`, 'system');
            showNotification('文件导出成功', 'success');
        } else {
            const errorData = await response.json();
            throw new Error(errorData.message || '导出失败');
        }
    } catch (error) {
        console.error('导出文件失败:', error);
        addConsoleLog('导出文件失败: ' + error.message, 'error');
        showNotification('导出文件失败: ' + error.message, 'error');
    }
}

// 删除文件
async function deleteFile(filename) {
    if (!filename) {
        showNotification('文件名无效', 'error');
        return;
    }
    
    // 确认删除
    const confirmed = confirm(`确定要删除文件 "${filename}" 吗？\n\n此操作将永久删除该文件的所有数据，无法恢复。`);
    if (!confirmed) {
        return;
    }
    
    try {
        addConsoleLog(`正在删除文件 "${filename}"...`, 'system');
        showNotification('正在删除文件...', 'info');
        
        const response = await fetch(`/uploaded-files/${encodeURIComponent(filename)}/delete`, {
            method: 'DELETE'
        });
        
        const data = await response.json();
        
        if (data.success) {
            addConsoleLog(`文件 "${filename}" 删除成功`, 'system');
            showNotification('文件删除成功', 'success');
            
            // 如果当前正在查看被删除的文件，需要清空表格显示
            const currentFilename = getCurrentDisplayedFilename();
            if (currentFilename === filename) {
                currentData = [];
                filteredData = [];
                currentSchema = [];
                renderTable();
                // 不需要更新统计信息，因为数据概览部分已被删除
            }
            
            // 重新加载文件列表和表格选择器
            await Promise.all([
                loadTablesList(),
                loadTableList()
            ]);
            
            // 如果当前页没有文件了，回到上一页
            const totalPages = Math.ceil(filesData.length / filesPerPage);
            if (currentPage > totalPages && totalPages > 0) {
                currentPage = totalPages;
                renderFilesList();
            }
        } else {
            throw new Error(data.message || '删除失败');
        }
    } catch (error) {
        console.error('删除文件失败:', error);
        addConsoleLog('删除文件失败: ' + error.message, 'error');
        showNotification('删除文件失败: ' + error.message, 'error');
    }
}

// 获取当前显示的文件名（辅助函数）
function getCurrentDisplayedFilename() {
    // 从当前数据中获取文件名，假设所有记录都来自同一个文件
    if (currentData && currentData.length > 0) {
        return currentData[0].source_file;
    }
    return null;
}

// 刷新文件列表
async function refreshTablesList() {
    addConsoleLog('手动刷新文件列表...', 'system');
    await loadTablesList();
}

// 修改页面初始化，添加表格管理初始化
document.addEventListener('DOMContentLoaded', function() {
    addConsoleLog('DataMerge Pro 系统初始化完成', 'system');
    
    initializeEventListeners();
    initializeNavbar();
    initializeUploadArea(); // 初始化上传区域功能
    loadTableList();
    loadTablesList(); // 添加表格管理初始化
});

// 初始化上传区域功能
function initializeUploadArea() {
    // 检查是否有现有数据，如果有则自动折叠上传区域
    setTimeout(() => {
        if (tableGroups && tableGroups.length > 0) {
            collapseUploadArea(false); // 静默折叠，不显示动画
        }
    }, 1000);
    
    addConsoleLog('上传区域功能初始化完成', 'system');
}

// 切换上传区域展开/折叠状态
function toggleUploadArea() {
    const uploadCard = document.getElementById('uploadCard');
    const isCollapsed = uploadCard.classList.contains('collapsed');
    
    if (isCollapsed) {
        expandUploadArea();
    } else {
        collapseUploadArea();
    }
}

// 展开上传区域
function expandUploadArea() {
    const uploadCard = document.getElementById('uploadCard');
    const uploadBody = document.getElementById('uploadBody');
    
    if (!uploadCard || uploadCard.classList.contains('expanding')) return;
    
    uploadCard.classList.add('expanding');
    uploadCard.classList.remove('collapsed');
    
    // 添加展开动画
    uploadBody.style.animation = 'expandUpload 0.4s cubic-bezier(0.4, 0, 0.2, 1) forwards';
    
    setTimeout(() => {
        uploadCard.classList.remove('expanding');
        uploadBody.style.animation = '';
        addConsoleLog('上传区域已展开', 'system');
    }, 400);
}

// 折叠上传区域
function collapseUploadArea(withAnimation = true) {
    const uploadCard = document.getElementById('uploadCard');
    const uploadBody = document.getElementById('uploadBody');
    
    if (!uploadCard || uploadCard.classList.contains('collapsing')) return;
    
    if (withAnimation) {
        uploadCard.classList.add('collapsing');
        
        // 添加折叠动画
        uploadBody.style.animation = 'collapseUpload 0.4s cubic-bezier(0.4, 0, 0.2, 1) forwards';
        
        setTimeout(() => {
            uploadCard.classList.add('collapsed');
            uploadCard.classList.remove('collapsing');
            uploadBody.style.animation = '';
            addConsoleLog('上传区域已折叠', 'system');
        }, 400);
    } else {
        // 静默折叠，无动画
        uploadCard.classList.add('collapsed');
    }
}

// 自动折叠上传区域（在文件上传成功后调用）
function autoCollapseUploadArea() {
    setTimeout(() => {
        collapseUploadArea(true);
        showNotification('上传完成', '上传区域已自动收起，为表格预留更多空间', 'info');
    }, 2000); // 上传完成2秒后自动折叠
}