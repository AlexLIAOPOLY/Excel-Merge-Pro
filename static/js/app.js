// 全局变量
let currentData = [];
let filteredData = [];
let currentSchema = [];
let tableGroups = [];
let currentGroupId = null;

// 全局搜索状态变量
let isGlobalSearchActive = false;
let currentGlobalSearchTerm = '';
let currentStats = null;

// 新上传文件跟踪变量
let newUploadedFiles = new Set();


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
    
    // 检查URL参数
    const urlParams = new URLSearchParams(window.location.search);
    const fileParam = urlParams.get('file');
    if (fileParam) {
        // 如果是数字，说明是旧的group ID方式
        if (/^\d+$/.test(fileParam)) {
            currentGroupId = parseInt(fileParam);
            addConsoleLog(`从工作台跳转，将加载分组ID: ${currentGroupId}`, 'system');
        } else {
            // 如果不是数字，说明是文件名方式
            window.currentFileName = decodeURIComponent(fileParam);
            addConsoleLog(`从工作台跳转，将加载文件: ${window.currentFileName}`, 'system');
        }
    }
    
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
    
    // 调试：检查元素是否存在
    if (!uploadArea) {
        console.error('uploadArea元素未找到');
        return;
    }
    console.log('uploadArea元素找到，开始添加拖拽事件监听器');
    
    // 添加点击触发文件选择事件
    uploadArea.addEventListener('click', function(e) {
        // 避免按钮点击时触发
        if (!e.target.classList.contains('upload-action-btn')) {
            console.log('uploadArea被点击了，触发文件选择');
            addConsoleLog('上传区域被点击，打开文件选择', 'system');
            document.getElementById('fileInput').click();
        }
    });
    
    // 添加更多拖拽相关的事件监听器用于调试
    uploadArea.addEventListener('dragenter', function(e) {
        console.log('dragenter事件触发');
        addConsoleLog('文件进入拖拽区域', 'system');
        e.preventDefault();
    });
    
    // 文件选择事件 - 选择后自动上传
    fileInput.addEventListener('change', function() {
        const files = Array.from(this.files);
        
        if (files.length > 0) {
            addConsoleLog(`已选择 ${files.length} 个文件，开始自动上传`, 'system');
            // 直接自动上传，不显示文件列表
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
    
    // 搜索模式切换事件
    const searchMode = document.getElementById('searchMode');
    searchMode.addEventListener('change', function() {
        // 当切换搜索模式时，重新执行搜索
        if (searchInput.value.trim()) {
            searchData();
        }
        updateSearchPlaceholder();
    });
    
    // 初始化搜索框占位符
    updateSearchPlaceholder();
    
    // 增强的拖拽上传功能
    uploadArea.addEventListener('dragover', function(e) {
        console.log('dragover事件触发');
        addConsoleLog('检测到文件拖拽', 'system');
        e.preventDefault();
        e.stopPropagation();
        this.classList.add('dragover');
        
        // 显示拖拽提示
        const title = this.querySelector('.upload-title');
        const desc = this.querySelector('.upload-desc');
        if (title && desc) {
            title.textContent = '释放文件以开始上传';
            desc.innerHTML = '支持批量拖拽多个Excel文件<br/>系统将自动处理并合并数据';
        }
    });
    
    uploadArea.addEventListener('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // 检查是否真的离开了拖拽区域（避免子元素触发）
        if (!this.contains(e.relatedTarget)) {
        this.classList.remove('dragover');
            
            // 恢复原始文本
            const title = this.querySelector('.upload-title');
            const desc = this.querySelector('.upload-desc');
            if (title && desc) {
                title.textContent = '拖拽Excel文件到此处上传';
                desc.innerHTML = '支持 .xlsx 和 .xls 格式，单文件最大 32MB<br/>支持批量拖拽上传多个文件';
            }
        }
    });
    
    uploadArea.addEventListener('drop', function(e) {
        console.log('drop事件触发');
        addConsoleLog('文件被释放，开始处理', 'system');
        e.preventDefault();
        e.stopPropagation();
        this.classList.remove('dragover');
        
        // 恢复原始文本  
        const title = this.querySelector('.upload-title');
        const desc = this.querySelector('.upload-desc');
        if (title && desc) {
            title.textContent = '拖拽Excel文件到此处上传';
            desc.innerHTML = '支持 .xlsx 和 .xls 格式，单文件最大 32MB<br/>支持批量拖拽上传多个文件';
        }
        
        const files = Array.from(e.dataTransfer.files);
        const excelFiles = files.filter(file => 
            file.name.toLowerCase().endsWith('.xlsx') || 
            file.name.toLowerCase().endsWith('.xls')
        );
        
        if (excelFiles.length > 0) {
            addConsoleLog(`通过拖拽选择了 ${excelFiles.length} 个Excel文件，开始自动上传`, 'system');
            showNotification('开始上传', `正在处理 ${excelFiles.length} 个Excel文件...`, 'info');
            
            // 直接上传拖拽的文件
            uploadDraggedFiles(excelFiles);
        } else {
            addConsoleLog('请选择Excel文件 (.xlsx 或 .xls)', 'warning');
            showNotification('文件格式错误', '请拖拽Excel文件（.xlsx 或 .xls格式）', 'error');
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
    
    // 显示进度条并开始动画
    const progressController = animateUploadProgress(4000);
    addConsoleLog('文件上传中...', 'system');
    
    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        // 完成上传，清理定时器并更新到100%
        progressController.clearAll();
        hideVPNTip(); // 隐藏VPN提示
        updateUploadProgress(100, '上传完成！');
        
        const result = await response.json();
        displayUploadResults(result.results);
        
        // 只刷新左侧的表格列表，右侧文件管理列表已在每个文件成功时实时刷新
        await loadTableList();
        
        // 清空文件选择
        fileInput.value = '';
        document.getElementById('fileList').style.display = 'none';
        
    } catch (error) {
        // 上传失败，检查是否需要显示VPN提示
        progressController.incrementFailure(error, error.message);
        progressController.clearAll();
        hideVPNTip(); // 隐藏VPN提示
        hideUploadProgress();
        addConsoleLog(`上传失败: ${error.message}`, 'error');
    } finally {
        addConsoleLog('文件上传完成', 'system');
        // 延迟隐藏进度条
        setTimeout(() => {
            hideUploadProgress();
        }, 800);
    }
}

// 处理拖拽文件上传
async function uploadDraggedFiles(files) {
    if (!files || files.length === 0) {
        addConsoleLog('没有可上传的文件', 'warning');
        return;
    }
    
    addConsoleLog(`开始上传 ${files.length} 个拖拽文件...`, 'system');
    
    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
        formData.append('files[]', files[i]);
    }
    
    // 显示进度条并开始动画
    const progressController = animateUploadProgress(4000);
    addConsoleLog('文件上传中...', 'system');
    
    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });
        
        // 完成上传，清理定时器并更新到100%
        progressController.clearAll();
        hideVPNTip(); // 隐藏VPN提示
        updateUploadProgress(100, '上传完成！');
        
        const result = await response.json();
        displayUploadResults(result.results);
        
        // 只刷新左侧的表格列表，右侧文件管理列表已在每个文件成功时实时刷新
        await loadTableList();
        
    } catch (error) {
        // 拖拽上传失败，检查是否需要显示VPN提示
        progressController.incrementFailure(error, error.message);
        progressController.clearAll();
        hideVPNTip(); // 隐藏VPN提示
        hideUploadProgress();
        addConsoleLog(`拖拽上传失败: ${error.message}`, 'error');
        showNotification('上传失败', error.message, 'error');
    } finally {
        addConsoleLog('拖拽文件上传完成', 'system');
        // 延迟隐藏进度条
        setTimeout(() => {
            hideUploadProgress();
        }, 800);
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
                
                // 记录新上传的文件
                newUploadedFiles.add(result.filename);
                
                // 立即刷新文件管理列表，让用户看到新上传的文件
                loadTablesList().then(() => {
                    // 刷新完成后等待一下让动画效果显示
                    console.log(`文件 ${result.filename} 已添加到文件管理列表`);
                }).catch(error => {
                    console.error('刷新文件列表失败:', error);
                });
                
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

// 从工作台导入文件
async function openWorkspaceUpload() {
    try {
        // 获取工作台文件列表
        const response = await fetch('/api/workspace/files');
        const data = await response.json();
        
        if (!data.success) {
            showNotification('获取文件失败', data.message, 'error');
            return;
        }

        const files = data.files;
        if (files.length === 0) {
            showNotification('工作台为空', '工作台中没有可导入的文件，请先上传文件到工作台', 'info');
            return;
        }

        // 创建文件选择模态框
        showWorkspaceFileModal(files);
    } catch (error) {
        console.error('获取工作台文件失败:', error);
        showNotification('获取失败', '无法连接到工作台，请稍后重试', 'error');
    }
}

// 显示工作台文件选择模态框
function showWorkspaceFileModal(files) {
    // 对文件进行排序：已导入的文件(has_data=true)排在前面
    const sortedFiles = files.sort((a, b) => {
        // 优先级：已导入 > 未导入
        if (a.has_data && !b.has_data) return -1;
        if (!a.has_data && b.has_data) return 1;
        
        // 如果状态相同，按记录数排序（多的在前）
        if (a.records !== b.records) {
            return b.records - a.records;
        }
        
        // 最后按文件名排序
        return a.name.localeCompare(b.name);
    });

    // 创建模态框HTML
    const modal = document.createElement('div');
    modal.className = 'workspace-modal-overlay';
    modal.innerHTML = `
        <div class="workspace-modal">
            <div class="workspace-modal-header">
                <h3>从工作台导入文件</h3>
                <button class="workspace-modal-close" onclick="closeWorkspaceModal()">&times;</button>
            </div>
            <div class="workspace-modal-body">
                <div class="workspace-file-list">
                    ${sortedFiles.map((file, index) => `
                        <div class="workspace-file-item" 
                             data-file-name="${file.name}" 
                             data-file-path="${file.path}"
                             data-index="${index}">
                            <input type="checkbox" class="workspace-file-checkbox" id="file-${index}">
                            <label for="file-${index}" class="workspace-file-label">
                                <div class="workspace-file-info">
                                    <div class="workspace-file-name">
                                        ${file.name}
                                    </div>
                                    <div class="workspace-file-details">
                                        <span class="file-folder">${file.folder}</span> | 
                                        <span class="file-size">${file.size_formatted}</span> | 
                                        <span class="file-records">${file.records}条记录</span> | 
                                        <span class="file-status ${file.has_data ? 'has-data' : 'no-data'}">${file.has_data ? '已导入' : '未导入'}</span>
                                    </div>
                                </div>
                            </label>
                        </div>
                    `).join('')}
                </div>
            </div>
            <div class="workspace-modal-footer">
                <span id="selectedFileCount">已选择 0 个文件</span>
                <div class="workspace-modal-actions">
                    <button class="workspace-action-btn secondary" onclick="selectAllFiles()">全选</button>
                    <button class="workspace-action-btn secondary" onclick="clearAllSelections()">清空</button>
                    <button class="workspace-action-btn secondary" onclick="closeWorkspaceModal()">取消</button>
                    <button class="workspace-action-btn primary" onclick="importSelectedFiles()">导入选中文件</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // 绑定复选框事件（支持Ctrl键多选）
    bindWorkspaceCheckboxEvents();
    
    // 显示动画
    setTimeout(() => modal.classList.add('show'), 10);
}

// 绑定工作台文件复选框事件
function bindWorkspaceCheckboxEvents() {
    const checkboxes = document.querySelectorAll('.workspace-file-checkbox');
    const countElement = document.getElementById('selectedFileCount');
    let lastSelectedIndex = -1;
    
    checkboxes.forEach((checkbox, index) => {
        const fileItem = checkbox.closest('.workspace-file-item');
        
        // 复选框变化事件
        checkbox.addEventListener('change', () => {
            updateFileSelectionCount();
        });
        
        // 点击文件项支持Ctrl键和Shift键多选
        fileItem.addEventListener('click', (e) => {
            // 如果点击的是复选框本身，让其正常处理
            if (e.target.type === 'checkbox' || e.target.tagName === 'LABEL') {
                lastSelectedIndex = index;
                return;
            }
            
            e.preventDefault();
            
            if (e.ctrlKey || e.metaKey) {
                // Ctrl键多选：切换当前项的选中状态
                checkbox.checked = !checkbox.checked;
                lastSelectedIndex = index;
            } else if (e.shiftKey && lastSelectedIndex !== -1) {
                // Shift键范围选择：选择从lastSelectedIndex到当前index的所有项
                const start = Math.min(lastSelectedIndex, index);
                const end = Math.max(lastSelectedIndex, index);
                
                for (let i = start; i <= end; i++) {
                    if (checkboxes[i]) {
                        checkboxes[i].checked = true;
                    }
                }
            } else {
                // 普通点击：清空其他选择，只选择当前项
                checkboxes.forEach(cb => cb.checked = false);
                checkbox.checked = true;
                lastSelectedIndex = index;
            }
            
            updateFileSelectionCount();
        });
        
        // 双击文件项只选择，不自动导入
        fileItem.addEventListener('dblclick', (e) => {
            e.preventDefault();
            // 清空其他选择，只选择当前项
            checkboxes.forEach(cb => cb.checked = false);
            checkbox.checked = true;
            updateFileSelectionCount();
        });
    });
    
    // 初始化计数
    updateFileSelectionCount();
}

// 更新文件选择计数
function updateFileSelectionCount() {
            const selectedCount = document.querySelectorAll('.workspace-file-checkbox:checked').length;
    const totalCount = document.querySelectorAll('.workspace-file-checkbox').length;
    const countElement = document.getElementById('selectedFileCount');
    
    if (countElement) {
        countElement.textContent = `已选择 ${selectedCount} / ${totalCount} 个文件`;
    }
    
    // 更新导入按钮状态
    const importBtn = document.querySelector('.workspace-action-btn.primary');
    if (importBtn) {
        importBtn.disabled = selectedCount === 0;
        if (selectedCount === 0) {
            importBtn.textContent = '请选择文件';
        } else {
            importBtn.textContent = `导入选中文件 (${selectedCount})`;
        }
    }
}

// 全选文件
function selectAllFiles() {
    const checkboxes = document.querySelectorAll('.workspace-file-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = true;
    });
    updateFileSelectionCount();
}

// 清空所有选择
function clearAllSelections() {
    const checkboxes = document.querySelectorAll('.workspace-file-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
    updateFileSelectionCount();
}

// 关闭工作台模态框
function closeWorkspaceModal() {
    const modal = document.querySelector('.workspace-modal-overlay');
    if (modal) {
        modal.classList.add('closing');
        setTimeout(() => modal.remove(), 300);
    }
}

// 导入选中的文件
async function importSelectedFiles() {
    const selectedCheckboxes = document.querySelectorAll('.workspace-file-checkbox:checked');
    
    if (selectedCheckboxes.length === 0) {
        showNotification('未选择文件', '请至少选择一个文件进行导入', 'warning');
        return;
    }

    const selectedFiles = Array.from(selectedCheckboxes).map(checkbox => {
        const fileItem = checkbox.closest('.workspace-file-item');
        return {
            name: fileItem.dataset.fileName,
            path: fileItem.dataset.filePath
        };
    });

    // 关闭模态框
    closeWorkspaceModal();

    // 显示进度条
    showUploadProgress();
    hideVPNTip(); // 确保开始时隐藏VPN提示
    updateUploadProgress(0, `准备导入 ${selectedFiles.length} 个文件...`);

    // API失败计数和VPN提示逻辑
    let apiFailureCount = 0;
    let vpnTipShown = false;
    
    // 长时间等待后检查是否需要显示VPN提示
    let vpnTipTimer = setTimeout(() => {
        if (!vpnTipShown && apiFailureCount >= 3) {
            showVPNTip();
            vpnTipShown = true;
        }
    }, 15000); // 15秒后检查是否需要显示VPN提示

    // 导入文件
    addConsoleLog(`开始从工作台导入 ${selectedFiles.length} 个文件...`, 'system');
    
    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        const progress = ((i + 1) / selectedFiles.length) * 100;
        
        try {
            // 确保进度不会为负数
            const displayProgress = Math.max(5, Math.min(95, progress * 0.8));
            updateUploadProgress(displayProgress, `正在导入: ${file.name}`);
            addConsoleLog(`正在导入: ${file.name}`, 'info');
            
            const response = await fetch('/api/workspace/files/import', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    file_name: file.name,
                    file_path: file.path
                })
            });

            const result = await response.json();
            
            if (result.success) {
                // 确保最终进度显示合理
                const finalProgress = Math.max(10, Math.min(95, progress * 0.85));
                updateUploadProgress(finalProgress, `${file.name} 导入成功`);
                addConsoleLog(`${file.name} 导入成功，共 ${result.count || 0} 条记录`, 'success');
                showNotification('导入成功', `${file.name} 已成功导入`, 'success');
                
                // 记录新导入的文件
                newUploadedFiles.add(file.name);
                
                // 立即刷新文件管理列表，让用户看到新导入的文件
                loadTablesList().then(() => {
                    console.log(`文件 ${file.name} 已添加到文件管理列表`);
                }).catch(error => {
                    console.error('刷新文件列表失败:', error);
                });
            } else {
                // API调用失败，检查是否是网络相关错误
                if (isNetworkError(null, result.message)) {
                    apiFailureCount++;
                    if (apiFailureCount >= 3 && !vpnTipShown) {
                        showVPNTip();
                        vpnTipShown = true;
                    }
                }
                addConsoleLog(`${file.name} 导入失败: ${result.message}`, 'error');
                showNotification('导入失败', `${file.name}: ${result.message}`, 'error');
            }
        } catch (error) {
            // 检查是否是网络错误
            if (isNetworkError(error, error.message)) {
                apiFailureCount++;
                if (apiFailureCount >= 3 && !vpnTipShown) {
                    showVPNTip();
                    vpnTipShown = true;
                }
            }
            console.error('导入文件失败:', error);
            addConsoleLog(`${file.name} 导入失败: 网络错误`, 'error');
            showNotification('导入失败', `${file.name}: 网络错误`, 'error');
        }
        
        // 每个文件之间添加短暂延迟
        await new Promise(resolve => setTimeout(resolve, 300));
    }

    // 完成导入，清理定时器
    clearTimeout(vpnTipTimer);
    hideVPNTip(); // 隐藏VPN提示
    updateUploadProgress(100, '导入完成！');

    // 导入完成后刷新表格组列表并自动收起上传框
    setTimeout(() => {
        // 检查 loadTableGroups 函数是否存在
        if (typeof loadTableGroups === 'function') {
            loadTableGroups();
        } else {
            // 如果函数不存在，刷新页面数据
            loadData();
        }
        
        // 自动收起文件上传区域
        collapseUploadArea(true); // 传入true表示静默折叠
        
        addConsoleLog('工作台文件导入完成，正在刷新数据...', 'system');
        showNotification('导入完成', '文件上传区域已自动收起', 'success');
        
        // 隐藏进度条
        hideUploadProgress();
    }, 500);
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
    
    // 添加到容器顶部（新通知在上方）
    container.insertBefore(notification, container.firstChild);
    
    // 用于控制自动隐藏的变量
    let autoHideTimer;
    
    // 鼠标悬停时暂停自动隐藏
    notification.addEventListener('mouseenter', () => {
        clearTimeout(autoHideTimer);
    });
    
    // 鼠标离开时重新开始自动隐藏倒计时
    notification.addEventListener('mouseleave', () => {
        autoHideTimer = setTimeout(() => {
            hideNotification(notification, container);
        }, 1500);
    });
    
    // 触发显示动画
    setTimeout(() => {
        notification.classList.add('show');
    }, 50);
    
    // 根据消息类型设置不同的显示时长
    const displayDuration = type === 'error' ? 4000 : type === 'success' ? 3000 : 2500;
    
    // 自动隐藏
    autoHideTimer = setTimeout(() => {
        hideNotification(notification, container);
    }, displayDuration);
}

// 隐藏通知 - 两阶段动画
function hideNotification(notification, container) {
    // 防止重复触发隐藏
    if (notification.classList.contains('hide') || notification.classList.contains('collapse')) {
        return;
    }
    
    // 第一阶段：向右滑出（保持高度）
    notification.classList.add('hide');
    
    // 第二阶段：在滑出动画进行中途开始收缩高度
    setTimeout(() => {
        notification.classList.add('collapse');
        
        // 动画完成后移除DOM元素
        setTimeout(() => {
            if (container.contains(notification)) {
                container.removeChild(notification);
            }
        }, 400); // 等待收缩动画完成
    }, 200); // 稍微提前开始收缩，让过渡更自然
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
    `;
    
    // 隐藏工具栏
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
    
    // 保持工具栏显示
    showToolbarButtons();
}

// 控制工具栏的显示/隐藏
function hideToolbarButtons() {
    const toolbar = document.querySelector('.toolbar');
    if (toolbar) {
        toolbar.style.display = 'none';
    }
}

function showToolbarButtons() {
    const toolbar = document.querySelector('.toolbar');
    if (toolbar) {
        toolbar.style.display = 'flex';
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
    
    // 显示工具栏
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

// 更新搜索框占位符
function updateSearchPlaceholder() {
    const searchInput = document.getElementById('searchInput');
    const searchMode = document.getElementById('searchMode').value;
    
    if (searchMode === 'global') {
        searchInput.placeholder = '全局搜索所有表格数据...（跨表格搜索）';
    } else {
        searchInput.placeholder = '搜索当前表格内容...（实时搜索）';
    }
}

// 搜索数据
function searchData() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
    const searchMode = document.getElementById('searchMode').value;
    
    if (!searchTerm) {
        filteredData = [...currentData];
        addConsoleLog('已清除搜索条件', 'system');
        renderTable();
    } else {
        if (searchMode === 'global') {
            // 全局搜索 - 搜索所有表格组
            performGlobalSearch(searchTerm);
        } else {
            // 当前表格搜索
            filteredData = currentData.filter(row => {
                return searchInRow(row, searchTerm);
            });
            
            addConsoleLog(`搜索 "${searchTerm}" 找到 ${filteredData.length} 条记录`, 'system');
            renderTable();
            highlightSearchResults(searchTerm);
        }
    }
}

// 全局搜索函数 - 修改版
async function performGlobalSearch(searchTerm) {
    try {
        addConsoleLog(`开始全局搜索 "${searchTerm}"...`, 'system');
        
        // 显示加载状态
        showLoadingState();
        
        // 调用新的全局搜索API
        const response = await fetch('/api/global-search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ search_term: searchTerm })
        });
        
        const result = await response.json();
        
        if (result.success) {
            // 直接显示搜索结果表格
            displayGlobalSearchResults(result);
        } else {
            addConsoleLog('全局搜索失败: ' + result.message, 'error');
            showNotification('搜索错误', result.message, 'error');
            hideLoadingState();
        }
        
    } catch (error) {
        console.error('全局搜索失败:', error);
        addConsoleLog('全局搜索失败: ' + error.message, 'error');
        showNotification('搜索错误', '全局搜索请求失败，请检查网络连接', 'error');
        hideLoadingState();
    }
}

// 全局搜索中的行匹配函数
function searchInRowGlobal(row, searchTerm, schema) {
    // 确保搜索词是小写
    const lowerSearchTerm = searchTerm.toLowerCase();
    
    // 搜索列名
    for (let columnName of schema) {
        if (columnName.toLowerCase().includes(lowerSearchTerm)) {
            const value = row[columnName];
            if (value != null && String(value).trim() !== '') {
                return true;
            }
        }
    }
    
    // 搜索数据值
    for (let key in row) {
        if (key === 'id' || key.startsWith('_')) continue;
        
        const value = row[key];
        if (value != null) {
            const valueStr = String(value).toLowerCase();
            if (valueStr.includes(lowerSearchTerm)) {
                return true;
            }
        }
    }
    
    return false;
}

// 显示全局搜索结果 - 新版本，直接渲染表格
function displayGlobalSearchResults(searchResult) {
    const { data, schema, search_term, total_matches, matched_groups, stats } = searchResult;
    
    if (total_matches === 0) {
        addConsoleLog(`全局搜索 "${search_term}" 未找到匹配结果`, 'system');
        showNotification('搜索结果', '未找到匹配的数据', 'info');
        hideLoadingState();
        return;
    }
    
    // 记录当前搜索状态为全局搜索
    isGlobalSearchActive = true;
    currentGlobalSearchTerm = search_term;
    
    addConsoleLog(`全局搜索 "${search_term}" 在 ${matched_groups} 个表格中找到 ${total_matches} 条记录`, 'success');
    
    // 设置当前数据为搜索结果
    currentData = data;
    filteredData = [...data];
    currentSchema = schema;
    currentStats = stats;
    
    // 更新表格选择器显示为搜索结果
    updateTableSelector(`全局搜索结果: "${search_term}" (${total_matches}条记录)`);
    
    // 更新导出按钮文本
    updateExportButtonForGlobalSearch(total_matches);
    
    // 渲染表格
    renderTable();
    hideLoadingState();
    
    // 高亮搜索关键词
    highlightSearchResults(search_term);
    
    // 为来源信息列添加特殊样式
    highlightSourceColumns();
    
    // 显示搜索结果通知
    showNotification(
        '全局搜索完成', 
        `在 ${matched_groups} 个表格中找到 ${total_matches} 条匹配记录，已合并显示`, 
        'success'
    );
    
    // 添加全局搜索结果的特殊样式标识
    addGlobalSearchIndicator();
}

// 更新表格选择器显示
function updateTableSelector(displayText) {
    const selector = document.getElementById('tableSelector');
    if (selector) {
        // 清空现有选项
        selector.innerHTML = '';
        
        // 添加搜索结果选项
        const option = document.createElement('option');
        option.value = 'global_search_result';
        option.textContent = displayText;
        option.selected = true;
        selector.appendChild(option);
    }
}

// 重置搜索状态
function resetGlobalSearchState() {
    isGlobalSearchActive = false;
    currentGlobalSearchTerm = '';
    // 重置导出按钮
    resetExportButton();
    // 移除全局搜索标识
    removeGlobalSearchIndicator();
}

// 更新导出按钮为全局搜索模式
function updateExportButtonForGlobalSearch(matchCount) {
    const exportBtn = document.getElementById('exportBtn');
    if (exportBtn) {
        exportBtn.textContent = `导出搜索结果 (${matchCount}条)`;
        exportBtn.style.background = 'linear-gradient(135deg, #10b981 0%, #059669 100%)';
        exportBtn.title = '导出当前全局搜索结果到Excel';
    }
}

// 重置导出按钮
function resetExportButton() {
    const exportBtn = document.getElementById('exportBtn');
    if (exportBtn) {
        exportBtn.textContent = '导出Excel';
        exportBtn.style.background = '';
        exportBtn.title = '导出所有表格数据到Excel';
    }
}

// 添加全局搜索结果标识
function addGlobalSearchIndicator() {
    removeGlobalSearchIndicator(); // 先移除已有的
    
    // 不再添加任何标识，保持界面简洁
    return;
}

// 移除全局搜索结果标识
function removeGlobalSearchIndicator() {
    const indicator = document.getElementById('globalSearchIndicator');
    if (indicator) {
        indicator.remove();
    }
}

// 为来源信息列添加特殊样式
function highlightSourceColumns() {
    if (!isGlobalSearchActive) return;
    
    const table = document.getElementById('dataTable');
    if (!table) return;
    
    // 找到来源信息列的索引
    const headerCells = table.querySelectorAll('thead th');
    const sourceColumnIndices = [];
    
    headerCells.forEach((th, index) => {
        const columnTitle = th.querySelector('.column-title');
        if (columnTitle && columnTitle.textContent.startsWith('_source_')) {
            sourceColumnIndices.push(index);
            
            // 为表头添加特殊样式
            th.style.background = 'linear-gradient(135deg, #e0f2fe 0%, #b3e5fc 100%)';
            th.style.borderLeft = '3px solid #0288d1';
            
            // 更新列标题显示
            if (columnTitle.textContent === '_source_table') {
                columnTitle.textContent = '来源表格';
            } else if (columnTitle.textContent === '_source_file') {
                columnTitle.textContent = '来源文件';
            }
        }
    });
    
    // 为数据行的来源列添加特殊样式
    const dataRows = table.querySelectorAll('tbody tr');
    dataRows.forEach(row => {
        sourceColumnIndices.forEach(colIndex => {
            const cell = row.children[colIndex];
            if (cell) {
                cell.style.background = 'rgba(224, 242, 254, 0.5)';
                cell.style.borderLeft = '2px solid #0288d1';
                cell.style.fontWeight = '500';
                cell.style.color = '#0277bd';
            }
        });
    });
}

// 显示加载状态
function showLoadingState() {
    const loadingState = document.getElementById('loadingState');
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    
    if (loadingState) {
        loadingState.style.display = 'flex';
    }
    if (dataTable) {
        dataTable.style.display = 'none';
    }
    if (emptyState) {
        emptyState.style.display = 'none';
    }
}

// 隐藏加载状态
function hideLoadingState() {
    const loadingState = document.getElementById('loadingState');
    const dataTable = document.getElementById('dataTable');
    const emptyState = document.getElementById('emptyState');
    
    if (loadingState) {
        loadingState.style.display = 'none';
    }
    
    // 根据是否有数据决定显示表格还是空状态
    if (currentData && currentData.length > 0) {
        if (dataTable) {
            dataTable.style.display = 'table';
        }
        if (emptyState) {
            emptyState.style.display = 'none';
        }
    } else {
        if (dataTable) {
            dataTable.style.display = 'none';
        }
        if (emptyState) {
            emptyState.style.display = 'block';
        }
    }
}

// 增强搜索输入框的用户体验
function enhanceSearchInput() {
    const searchInput = document.getElementById('searchInput');
    const searchMode = document.getElementById('searchMode');
    
    if (searchInput && searchMode) {
        // 监听搜索模式变化
        searchMode.addEventListener('change', function() {
            updateSearchPlaceholder();
            
            // 如果当前是全局搜索状态，但切换到了当前表格模式，需要重置
            if (isGlobalSearchActive && this.value === 'current') {
                resetSearch();
            }
        });
        
        // 在搜索输入框添加搜索状态提示
        searchInput.addEventListener('focus', function() {
            if (isGlobalSearchActive) {
                showNotification(
                    '搜索提示', 
                    '当前显示全局搜索结果，可以继续在结果中筛选', 
                    'info'
                );
            }
        });
        
        // 支持回车键触发搜索
        searchInput.addEventListener('keypress', function(event) {
            if (event.key === 'Enter') {
                event.preventDefault();
                const searchTerm = searchInput.value.trim();
                if (searchTerm) {
                    searchData();
                    addConsoleLog(`回车键执行搜索: "${searchTerm}"`, 'system');
                }
            }
        });
    }
}

// 切换到指定的表格组
async function switchToTableGroup(groupId) {
    closeGlobalSearchModal();
    
    // 更新URL和页面状态
    const newUrl = new URL(window.location);
    newUrl.searchParams.set('group_id', groupId);
    window.history.pushState({}, '', newUrl);
    
    // 加载指定的表格组
    await loadTableData(groupId);
    
    // 保持搜索词，但切换到当前表格模式
    document.getElementById('searchMode').value = 'current';
    updateSearchPlaceholder();
    
    // 重新执行当前表格搜索
    const searchTerm = document.getElementById('searchInput').value.trim();
    if (searchTerm) {
        setTimeout(() => searchData(), 300);
    }
    
    addConsoleLog(`已切换到表格: ${groupId}`, 'system');
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
    
    // 如果当前是全局搜索状态，需要切换回普通表格
    if (isGlobalSearchActive) {
        resetGlobalSearchState();
        // 重新加载原来的表格数据
        loadData();
        // 切换搜索模式为当前表格
        const searchMode = document.getElementById('searchMode');
        if (searchMode) {
            searchMode.value = 'current';
            updateSearchPlaceholder();
        }
        addConsoleLog('已退出全局搜索模式并重置搜索条件', 'system');
    } else {
        // 普通的重置搜索
        filteredData = [...currentData];
        clearHighlights(); // 清除高亮
        renderTable();
        addConsoleLog('已重置搜索条件', 'system');
    }
    
    searchInput.focus(); // 聚焦到搜索框
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



// 处理导出按钮点击
async function handleExportClick() {
    if (isGlobalSearchActive) {
        // 导出全局搜索结果
        await exportGlobalSearchResults();
    } else {
        // 导出所有表格分组
        await exportAllGroups();
    }
}

// 导出全局搜索结果
async function exportGlobalSearchResults() {
    if (!isGlobalSearchActive || !currentData || currentData.length === 0) {
        addConsoleLog('没有全局搜索结果可导出', 'warning');
        showNotification('导出错误', '没有搜索结果可导出', 'error');
        return;
    }
    
    addConsoleLog(`开始导出全局搜索结果 "${currentGlobalSearchTerm}"...`, 'system');
    
    try {
        const response = await fetch('/api/export-global-search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                search_term: currentGlobalSearchTerm,
                data: currentData,
                schema: currentSchema
            })
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            // 构造文件名
            const safeSearchTerm = currentGlobalSearchTerm.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_');
            const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '_');
            a.download = `全局搜索结果_${safeSearchTerm}_${currentData.length}条_${timestamp}.xlsx`;
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            addConsoleLog(`全局搜索结果导出成功，共 ${currentData.length} 条记录`, 'success');
            showNotification('导出成功', `全局搜索结果已成功导出，共 ${currentData.length} 条记录`, 'success');
        } else {
            const errorData = await response.json().catch(() => ({}));
            const errorMessage = errorData.message || '导出失败';
            addConsoleLog(`导出全局搜索结果失败: ${errorMessage}`, 'error');
            showNotification('导出失败', errorMessage, 'error');
        }
        
    } catch (error) {
        console.error('导出全局搜索结果时发生错误:', error);
        addConsoleLog(`导出全局搜索结果时发生错误: ${error.message}`, 'error');
        showNotification('导出错误', '导出过程中发生错误，请检查网络连接', 'error');
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
            
            addConsoleLog(`所有表格导出成功: ${filename}`, 'system');
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
            } else if (window.currentFileName) {
                // 如果是从工作台跳转的文件名，直接加载该文件数据
                loadFileData(window.currentFileName);
            } else {
                // 没有表格时显示空状态
                showEmptyState();
                updateStats({
                    total_records: 0,
                    source_files: 0,
                    total_columns: 0,
                    last_update: '--'
                });
                // 确保隐藏工具栏
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
        
        // 添加记录数显示
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
    
    addConsoleLog(`表格标题已更新为: ${filename}`, 'system');
}

// 返回表格组模式
async function backToTableGroups() {
    try {
        addConsoleLog('正在返回总表...', 'system');
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
        
        addConsoleLog(`已返回总表，当前显示: ${tableGroups.length > 0 ? tableGroups[0].group_name : '无表格'}`, 'system');
        showNotification('已返回总表', 'success');
        
    } catch (error) {
        console.error('返回总表失败:', error);
        addConsoleLog('返回总表失败: ' + error.message, 'error');
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
            
            // 同时刷新右侧文件管理列表
            await loadTablesList();
            
            addConsoleLog('所有数据已清空', 'system');
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
    
    // 去掉时间戳等后缀，保留主要名称
    currentName = currentName.replace(/_\d+$/, '');
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

// AI生成表格名称
async function generateAIName() {
    if (!currentGroupId) {
        addConsoleLog('请先选择一个表格', 'warning');
        return;
    }

    const aiBtn = document.getElementById('aiRenameBtn');
    const aiLoading = document.getElementById('aiRenameLoading');
    const newNameInput = document.getElementById('newTableName');

    // 显示加载状态
    aiBtn.disabled = true;
    aiBtn.textContent = '生成中...';
    aiLoading.style.display = 'flex';

    try {
        addConsoleLog('正在使用DeepSeek AI生成智能表格名称...', 'info');
        
        const response = await fetch('/api/ai-rename-table', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                group_id: currentGroupId
            })
        });

        const result = await response.json();

        if (result.success) {
            // 将AI生成的名称填入输入框
            newNameInput.value = result.new_name;
            addConsoleLog(`AI成功生成表格名称: ${result.new_name}`, 'success');
            
            // 立即应用重命名并刷新表格列表
            setTimeout(async () => {
                await renameTableConfirm();
                await loadTableList();
            }, 500);
        } else {
            addConsoleLog(`AI生成名称失败: ${result.message}`, 'error');
        }

    } catch (error) {
        addConsoleLog(`AI生成名称时发生错误: ${error.message}`, 'error');
    } finally {
        // 恢复按钮状态
        aiBtn.disabled = false;
        aiBtn.textContent = 'AI生成';
        aiLoading.style.display = 'none';
    }
}

// 兼容原有的loadData函数 - 重新加载表格列表
// 加载特定文件的数据（从工作台跳转）
async function loadFileData(fileName) {
    addConsoleLog(`正在加载文件数据: ${fileName}`, 'system');
    
    try {
        const response = await fetch(`/api/workspace/files/${encodeURIComponent(fileName)}/data`);
        const result = await response.json();
        
        if (result.success) {
            currentData = result.data || [];
            filteredData = [...currentData];
            currentSchema = result.schema || [];
            
            // 查找该文件对应的表格分组（通过API直接查询）
            let groupInfo = null;
            try {
                const groupResponse = await fetch(`/api/workspace/files/${encodeURIComponent(fileName)}/group`);
                const groupResult = await groupResponse.json();
                if (groupResult.success && groupResult.group) {
                    groupInfo = groupResult.group;
                    currentGroupId = groupInfo.id;
                }
            } catch (e) {
                addConsoleLog(`查找文件分组时出错: ${e.message}`, 'error');
            }
            
            // 更新表格选择器
            const selector = document.getElementById('tableSelector');
            if (groupInfo) {
                // 重新加载表格列表以确保选择器有正确的选项
                await loadTableList();
                selector.value = groupInfo.id;
                addConsoleLog(`已切换到分组: ${groupInfo.group_name}`, 'system');
            } else {
                // 如果没有找到对应分组，显示文件名
                selector.innerHTML = `<option value="file:${fileName}" selected>${fileName}</option>`;
            }
            
            // 渲染表格
            renderTable();
            updateStats(result.stats);
            
            // 显示表格，隐藏空状态
            document.getElementById('dataTable').style.display = 'table';
            document.getElementById('emptyState').style.display = 'none';
            document.getElementById('loadingState').style.display = 'none';
            showToolbarButtons();
            
            addConsoleLog(`文件 ${fileName} 加载完成，共 ${currentData.length} 条记录`, 'system');
        } else {
            addConsoleLog(`加载文件数据失败: ${result.message}`, 'error');
            showEmptyState();
        }
    } catch (error) {
        addConsoleLog(`加载文件数据时发生错误: ${error.message}`, 'error');
        showEmptyState();
    }
}

async function loadData() {
    if (window.currentFileName) {
        // 如果当前是文件模式，重新加载文件数据
        await loadFileData(window.currentFileName);
    } else if (currentGroupId) {
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

// 文件列表数据
let filesData = [];

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
        
        loadingElement.style.display = 'none';
    } catch (error) {
        console.error('加载文件列表失败:', error);
        addConsoleLog('加载文件列表失败: ' + error.message, 'error');
        noTablesMessage.style.display = 'block';
        tablesListContainer.style.display = 'none';
        loadingElement.style.display = 'none';
    } finally {
        loadingElement.style.display = 'none';
    }
}

// 渲染文件列表
function renderFilesList() {
    const tablesList = document.getElementById('tablesList');
    
    if (!filesData || filesData.length === 0) {
        tablesList.innerHTML = '';
        return;
    }
    
    // 显示所有文件，不再分页
    const currentFiles = filesData;
    
    // 渲染当前页的文件
    let filesHtml = currentFiles.map(file => {
        const displayName = file.filename || `文件-${file.id}`;
        const recordCount = file.current_records || 0;
        const uploadTime = file.upload_time ? new Date(file.upload_time).toLocaleDateString() : '--';
        const hasData = file.has_data;
        
        // 检查是否是新上传的文件
        const isNewFile = newUploadedFiles.has(file.filename);
        let itemClass = 'table-item';
        if (isNewFile) {
            itemClass += ' new-file';
        }
        
        // 如果没有数据了，显示不同的样式
        const itemStyle = hasData ? '' : 'opacity: 0.6; border-color: #f87171;';
        const statusText = hasData ? '' : ' (数据已删除)';
        
        return `
            <div class="${itemClass}" data-filename="${file.filename}" style="${itemStyle}">
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
    
    tablesList.innerHTML = filesHtml;
    
    // 处理新文件动画
    if (newUploadedFiles.size > 0) {
        // 延迟一帧以确保DOM已更新，然后添加高亮动画
        setTimeout(() => {
            const filesToProcess = Array.from(newUploadedFiles);
            filesToProcess.forEach(filename => {
                const fileElement = document.querySelector(`.table-item[data-filename="${filename}"]`);
                if (fileElement) {
                    // 动画完成后移除new-file类，添加高亮效果
                    setTimeout(() => {
                        fileElement.classList.remove('new-file');
                        fileElement.classList.add('new-file-highlight');
                        
                        // 高亮动画完成后清理
                        setTimeout(() => {
                            fileElement.classList.remove('new-file-highlight');
                            // 从新文件记录中移除这个文件
                            newUploadedFiles.delete(filename);
                        }, 2000);
                    }, 600);
                }
            });
        }, 16);
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
    
    // 延迟加载表格列表，确保DOM元素已完全加载
    setTimeout(() => {
        addConsoleLog('开始加载文件管理列表...', 'system');
        loadTablesList(); // 添加表格管理初始化
    }, 500);
    
    // 添加调试函数到全局
    window.debugRefreshFiles = async function() {
        console.log('强制刷新文件列表...');
        await loadTablesList();
    };
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

// 进度条控制函数
function showUploadProgress() {
    const progressContainer = document.getElementById('uploadProgress');
    const progressText = document.getElementById('uploadProgressText');
    const progressPercent = document.getElementById('uploadProgressPercent');
    const progressFill = document.getElementById('uploadProgressFill');
    const uploadTip = document.getElementById('uploadTip');
    
    if (progressContainer) {
        progressContainer.style.display = 'block';
        progressText.textContent = '正在上传文件...';
        progressPercent.textContent = '0%';
        progressFill.style.width = '0%';
        if (uploadTip) {
            uploadTip.style.display = 'none';
        }
    }
}

function updateUploadProgress(percent, text = null) {
    const progressText = document.getElementById('uploadProgressText');
    const progressPercent = document.getElementById('uploadProgressPercent');
    const progressFill = document.getElementById('uploadProgressFill');
    
    if (progressPercent && progressFill) {
        // 确保百分比在0-100之间
        const safePercent = Math.max(0, Math.min(100, percent));
        progressPercent.textContent = Math.round(safePercent) + '%';
        progressFill.style.width = safePercent + '%';
    }
    
    if (text && progressText) {
        progressText.textContent = text;
    }
}

function hideUploadProgress() {
    const progressContainer = document.getElementById('uploadProgress');
    
    if (progressContainer) {
        setTimeout(() => {
            progressContainer.style.display = 'none';
        }, 1000); // 延迟隐藏，让用户看到完成状态
    }
}

// 判断是否是网络连接相关的错误
function isNetworkError(error, errorMessage) {
    // 检查JavaScript的网络错误类型
    if (error instanceof TypeError && error.message.includes('fetch')) {
        return true;
    }
    
    // 检查常见的网络相关错误消息
    const networkErrorPatterns = [
        'network',
        'connection',
        'timeout',
        'fetch',
        'deepseek',
        'api',
        '网络',
        '连接',
        '超时',
        '请求失败',
        'Failed to fetch',
        'NetworkError',
        'ERR_NETWORK',
        'ERR_INTERNET_DISCONNECTED'
    ];
    
    const message = (errorMessage || error?.message || '').toLowerCase();
    return networkErrorPatterns.some(pattern => 
        message.includes(pattern.toLowerCase())
    );
}

// 显示VPN提示
function showVPNTip() {
    const uploadTip = document.getElementById('uploadTip');
    if (uploadTip) {
        uploadTip.style.display = 'block';
    }
}

// 隐藏VPN提示
function hideVPNTip() {
    const uploadTip = document.getElementById('uploadTip');
    if (uploadTip) {
        uploadTip.style.display = 'none';
    }
}

// 模拟进度条动画（用于没有真实进度的API请求）
function animateUploadProgress(duration = 3000) {
    showUploadProgress();
    hideVPNTip(); // 确保开始时隐藏VPN提示
    
    let progress = 0;
    const increment = 100 / (duration / 50); // 每50ms增加的百分比
    let vpnTipShown = false;
    let apiFailureCount = 0;
    
    // 长时间等待或多次API失败后显示VPN提示
    const vpnTipTimer = setTimeout(() => {
        if (!vpnTipShown && apiFailureCount >= 3) {
            showVPNTip();
            vpnTipShown = true;
        }
    }, 15000); // 15秒后检查是否需要显示VPN提示
    
    const interval = setInterval(() => {
        progress += increment;
        
        if (progress >= 95) {
            // 在95%停止，等待实际请求完成
            clearInterval(interval);
            updateUploadProgress(95, '正在处理数据...');
        } else {
            let text = '正在上传文件...';
            if (progress > 30) text = '正在解析表格...';
            if (progress > 60) text = '正在处理数据...';
            
            updateUploadProgress(progress, text);
        }
    }, 50);
    
    // 返回对象包含interval和timer，便于清理
    const controller = {
        interval: interval,
        vpnTipTimer: vpnTipTimer,
        apiFailureCount: 0,
        incrementFailure: function(error, errorMessage) {
            // 只有在网络错误时才增加失败计数
            if (isNetworkError(error, errorMessage)) {
                this.apiFailureCount++;
                // 如果失败次数达到3次，立即显示VPN提示
                if (this.apiFailureCount >= 3 && !vpnTipShown) {
                    showVPNTip();
                    vpnTipShown = true;
                }
            }
        },
        clearAll: function() {
            clearInterval(this.interval);
            clearTimeout(this.vpnTipTimer);
        }
    };
    
    return controller;
}

// ========================= API配置管理 =========================

// API配置预设
const API_PRESETS = {
    deepseek: {
        url: 'https://api.deepseek.com',
        model: 'deepseek-chat',
        name: 'DeepSeek',
        models: [
            { value: 'deepseek-chat', label: 'deepseek-chat (通用对话模型)' },
            { value: 'deepseek-coder', label: 'deepseek-coder (代码专用模型)' },
            { value: 'deepseek-reasoner', label: 'deepseek-reasoner (推理增强模型)' }
        ]
    },
    openai: {
        url: 'https://api.openai.com/v1',
        model: 'gpt-3.5-turbo',
        name: 'OpenAI',
        models: [
            { value: 'gpt-3.5-turbo', label: 'GPT-3.5 Turbo (快速响应)' },
            { value: 'gpt-4', label: 'GPT-4 (高质量推理)' },
            { value: 'gpt-4-turbo', label: 'GPT-4 Turbo (性能平衡)' },
            { value: 'gpt-4o', label: 'GPT-4o (多模态)' }
        ]
    },
    anthropic: {
        url: 'https://api.anthropic.com/v1',
        model: 'claude-3-haiku-20240307',
        name: 'Anthropic',
        models: [
            { value: 'claude-3-haiku-20240307', label: 'Claude 3 Haiku (快速)' },
            { value: 'claude-3-sonnet-20240229', label: 'Claude 3 Sonnet (平衡)' },
            { value: 'claude-3-opus-20240229', label: 'Claude 3 Opus (顶级)' }
        ]
    },
    gemini: {
        url: 'https://generativelanguage.googleapis.com/v1beta',
        model: 'gemini-pro',
        name: 'Google Gemini',
        models: [
            { value: 'gemini-pro', label: 'Gemini Pro (通用)' },
            { value: 'gemini-pro-vision', label: 'Gemini Pro Vision (视觉)' },
            { value: 'gemini-1.5-pro', label: 'Gemini 1.5 Pro (最新)' }
        ]
    },
    zhipu: {
        url: 'https://open.bigmodel.cn/api/paas/v4',
        model: 'glm-4',
        name: '智谱AI',
        models: [
            { value: 'glm-4', label: 'GLM-4 (通用模型)' },
            { value: 'glm-4v', label: 'GLM-4V (视觉理解)' },
            { value: 'glm-3-turbo', label: 'GLM-3 Turbo (快速)' }
        ]
    },
    hunyuan: {
        url: 'https://api.hunyuan.cloud.tencent.com/v1',
        model: 'hunyuan-pro',
        name: '腾讯混元',
        models: [
            { value: 'hunyuan-pro', label: '混元-Pro (高性能)' },
            { value: 'hunyuan-standard', label: '混元-Standard (标准)' },
            { value: 'hunyuan-lite', label: '混元-Lite (轻量)' }
        ]
    },
    qwen: {
        url: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
        model: 'qwen-max',
        name: '阿里通义千问',
        models: [
            { value: 'qwen-max', label: '通义千问-Max (最强能力)' },
            { value: 'qwen-plus', label: '通义千问-Plus (平衡)' },
            { value: 'qwen-turbo', label: '通义千问-Turbo (快速)' }
        ]
    },
    doubao: {
        url: 'https://ark.cn-beijing.volces.com/api/v3',
        model: 'doubao-pro-4k',
        name: '字节豆包',
        models: [
            { value: 'doubao-pro-4k', label: '豆包-Pro-4K (高性能)' },
            { value: 'doubao-lite-4k', label: '豆包-Lite-4K (轻量)' },
            { value: 'doubao-pro-32k', label: '豆包-Pro-32K (长文本)' }
        ]
    },
    moonshot: {
        url: 'https://api.moonshot.cn/v1',
        model: 'moonshot-v1-8k',
        name: '月之暗面',
        models: [
            { value: 'moonshot-v1-8k', label: 'Moonshot v1 8K (8K上下文)' },
            { value: 'moonshot-v1-32k', label: 'Moonshot v1 32K (32K上下文)' },
            { value: 'moonshot-v1-128k', label: 'Moonshot v1 128K (128K上下文)' }
        ]
    },

    custom: {
        url: '',
        model: '',
        name: '自定义',
        models: [
            { value: '', label: '请手动输入模型名称' }
        ]
    }
};

// 显示API配置模态框
function showAPIConfigModal() {
    const modal = document.getElementById('apiConfigModal');
    if (modal) {
        modal.style.display = 'flex';
        loadAPIConfig();
    }
}

// 隐藏API配置模态框
function hideAPIConfigModal() {
    const modal = document.getElementById('apiConfigModal');
    if (modal) {
        modal.style.display = 'none';
        // 清空测试结果
        const testResult = document.getElementById('apiTestResult');
        if (testResult) {
            testResult.style.display = 'none';
        }
    }
}

// 加载当前API配置
function loadAPIConfig() {
    try {
        const savedConfig = localStorage.getItem('apiConfig');
        if (savedConfig) {
            const config = JSON.parse(savedConfig);
            
            // 先设置提供商，这会触发模型选项的更新
            document.getElementById('apiProvider').value = config.provider || 'deepseek';
            
            // 更新提供商配置（包括模型选项）
            updateAPIProviderConfig();
            
            // 然后设置其他配置
            document.getElementById('apiUrl').value = config.url || 'https://api.deepseek.com';
            document.getElementById('apiKey').value = config.key || '';
            
            // 最后设置模型（确保模型选项已经加载）
            setTimeout(() => {
                const modelElement = document.getElementById('apiModel');
                if (modelElement) {
                    modelElement.value = config.model || 'deepseek-chat';
                    
                    // 如果是自定义配置且模型不在预设列表中，切换到文本输入模式
                    if (config.provider === 'custom') {
                        const preset = API_PRESETS[config.provider];
                        const hasPresetModel = preset && preset.models && 
                                              preset.models.some(m => m.value === config.model);
                        
                        if (!hasPresetModel && config.model) {
                            // 模型不在预设中，切换到文本输入模式
                            updateCustomModelInput();
                            // 重新设置值
                            setTimeout(() => {
                                const newModelElement = document.getElementById('apiModel');
                                if (newModelElement) {
                                    newModelElement.value = config.model;
                                }
                            }, 10);
                        }
                    }
                }
            }, 50); // 短暂延迟确保DOM更新完成
            
        } else {
            // 使用默认配置
            updateAPIProviderConfig();
        }
    } catch (error) {
        console.error('加载API配置失败:', error);
        updateAPIProviderConfig();
    }
}

// 更新API提供商配置
function updateAPIProviderConfig() {
    const provider = document.getElementById('apiProvider').value;
    const preset = API_PRESETS[provider];
    const modelSelect = document.getElementById('apiModel');
    
    if (preset) {
        // 更新API地址
        document.getElementById('apiUrl').value = preset.url;
        
        // 清空现有模型选项
        modelSelect.innerHTML = '';
        
        // 动态添加模型选项
        if (preset.models && preset.models.length > 0) {
            preset.models.forEach(model => {
                const option = document.createElement('option');
                option.value = model.value;
                option.textContent = model.label;
                modelSelect.appendChild(option);
            });
            
            // 设置默认选择的模型
            modelSelect.value = preset.model;
        } else {
            // 如果没有预设模型，创建一个默认选项
            const option = document.createElement('option');
            option.value = preset.model;
            option.textContent = preset.model || '默认模型';
            modelSelect.appendChild(option);
            modelSelect.value = preset.model;
        }
        
        // 如果是自定义配置，允许手动输入
        if (provider === 'custom') {
            // 为自定义配置添加手动输入功能
            updateCustomModelInput();
        }
    }
    
    // 清空测试结果
    const testResult = document.getElementById('apiTestResult');
    if (testResult) {
        testResult.style.display = 'none';
        testResult.className = 'api-test-result';
        testResult.textContent = '';
    }
}

// 处理自定义模型输入
function updateCustomModelInput() {
    const modelElement = document.getElementById('apiModel');
    const provider = document.getElementById('apiProvider').value;
    
    if (provider === 'custom') {
        // 如果当前不是input，创建input
        if (modelElement && modelElement.tagName !== 'INPUT') {
            const currentValue = modelElement.value;
            const parent = modelElement.parentNode;
            
            // 创建输入框
            const input = document.createElement('input');
            input.type = 'text';
            input.id = 'apiModel';
            input.className = 'form-input';
            input.placeholder = '请输入模型名称，如：gpt-3.5-turbo';
            input.value = currentValue;
            
            // 替换select为input
            parent.replaceChild(input, modelElement);
        }
    } else {
        // 如果当前是input，切换回select
        if (modelElement && modelElement.tagName === 'INPUT') {
            const currentValue = modelElement.value;
            const parent = modelElement.parentNode;
            
            // 重新创建select
            const select = document.createElement('select');
            select.id = 'apiModel';
            select.className = 'form-select';
            
            // 替换input为select
            parent.replaceChild(select, modelElement);
            
            // 重新填充选项，但不递归调用updateAPIProviderConfig()
            const preset = API_PRESETS[provider];
            if (preset && preset.models) {
                preset.models.forEach(model => {
                    const option = document.createElement('option');
                    option.value = model.value;
                    option.textContent = model.label;
                    select.appendChild(option);
                });
                
                // 尝试恢复之前的值，如果不存在则用默认值
                const targetValue = currentValue && 
                                   preset.models.some(m => m.value === currentValue) ? 
                                   currentValue : preset.model;
                select.value = targetValue;
            }
        }
    }
}

// 切换API Key可见性
function toggleAPIKeyVisibility() {
    const apiKeyInput = document.getElementById('apiKey');
    const toggleBtn = document.querySelector('.toggle-visibility');
    
    if (apiKeyInput.type === 'password') {
        apiKeyInput.type = 'text';
        toggleBtn.textContent = '🙈';
    } else {
        apiKeyInput.type = 'password';
        toggleBtn.textContent = '👁️';
    }
}

// 测试API连接
async function testAPIConnection() {
    const testResult = document.getElementById('apiTestResult');
    const provider = document.getElementById('apiProvider').value;
    const url = document.getElementById('apiUrl').value;
    const key = document.getElementById('apiKey').value;
    const model = document.getElementById('apiModel').value;
    
    // 调试输出
    console.log('testResult element:', testResult);
    console.log('API config:', { provider, url: url ? 'set' : 'empty', key: key ? 'set' : 'empty', model });
    
    if (!url || !key) {
        testResult.className = 'api-test-result error';
        testResult.textContent = '请输入API地址和API Key';
        testResult.style.display = 'block'; // 强制显示
        return;
    }
    
    testResult.className = 'api-test-result loading';
    testResult.textContent = '正在测试连接...';
    testResult.style.display = 'block'; // 强制显示
    
    try {
        // 发送测试请求到后端
        const response = await fetch('/test-api-connection', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                provider: provider,
                url: url,
                key: key,
                model: model
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            testResult.className = 'api-test-result success';
            testResult.textContent = '连接成功，API配置正常';
            testResult.style.display = 'block'; // 强制显示
        } else {
            testResult.className = 'api-test-result error';
            testResult.textContent = '连接失败：' + (result.error || '未知错误');
            testResult.style.display = 'block'; // 强制显示
        }
    } catch (error) {
        testResult.className = 'api-test-result error';
        testResult.textContent = '测试失败：' + error.message;
        testResult.style.display = 'block'; // 强制显示
    }
}

// 保存API配置
function saveAPIConfig() {
    const provider = document.getElementById('apiProvider').value;
    const url = document.getElementById('apiUrl').value;
    const key = document.getElementById('apiKey').value;
    const model = document.getElementById('apiModel').value;
    
    if (!url || !key || !model) {
        showNotification('配置不完整', '请填写完整的API配置信息', 'error');
        return;
    }
    
    const config = {
        provider: provider,
        url: url,
        key: key,
        model: model,
        updated: new Date().toISOString()
    };
    
    try {
        localStorage.setItem('apiConfig', JSON.stringify(config));
        showNotification('配置保存成功', 'API配置已保存到本地', 'success');
        hideAPIConfigModal();
        addConsoleLog(`API配置已更新: ${API_PRESETS[provider]?.name || provider}`, 'system');
    } catch (error) {
        showNotification('保存失败', '配置保存失败：' + error.message, 'error');
    }
}

// 关闭移动端菜单
function closeNavMenu() {
    const navMenuMobile = document.getElementById('navMenuMobile');
    const navToggle = document.getElementById('navToggle');
    if (navMenuMobile) {
        navMenuMobile.classList.remove('active');
    }
    if (navToggle) {
        navToggle.classList.remove('active');
    }
}

// 初始化API配置 - 页面加载时执行
document.addEventListener('DOMContentLoaded', function() {
    // 检查是否有保存的API配置，如果没有则设置默认配置
    const savedConfig = localStorage.getItem('apiConfig');
    if (!savedConfig) {
        const defaultConfig = {
            provider: 'deepseek',
            url: 'https://api.deepseek.com',
            key: '', // 需要用户自行设置
            model: 'deepseek-chat',
            updated: new Date().toISOString()
        };
        localStorage.setItem('apiConfig', JSON.stringify(defaultConfig));
        console.log('已设置默认API配置：DeepSeek');
    }
});