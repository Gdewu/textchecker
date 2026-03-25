// HTML实体编码函数 - 防止XSS攻击
const escapeHtml = (unsafe) => {
    if (typeof unsafe !== 'string') return unsafe;
    return unsafe
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
};

// 清理危险HTML标签和事件处理器 - 防止XSS
const sanitizeHtml = (html) => {
    if (typeof html !== 'string') return html;
    // 移除 script、iframe、object、embed 等危险标签
    let sanitized = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
    sanitized = sanitized.replace(/<iframe[^>]*>[\s\S]*?<\/iframe>/gi, '');
    sanitized = sanitized.replace(/<object[^>]*>[\s\S]*?<\/object>/gi, '');
    sanitized = sanitized.replace(/<embed[^>]*>/gi, '');
    sanitized = sanitized.replace(/<form[^>]*>[\s\S]*?<\/form>/gi, '');
    // 移除事件处理器属性 (onerror, onclick, onload 等)
    sanitized = sanitized.replace(/\s*on\w+\s*=\s*["'][^"']*["']/gi, '');
    // 移除 javascript: 伪协议
    sanitized = sanitized.replace(/javascript:/gi, '');
    // 移除 data: URI (可能包含恶意代码)
    sanitized = sanitized.replace(/data:text\/html[^;]*;base64,/gi, '');
    return sanitized;
};

// 从文件名中提取姓名
// 支持格式: "姓名.docx", "姓名_其他.docx", "姓名-其他.docx", "姓名 其他.docx"
const extractNameFromFileName = (fileName) => {
    if (!fileName || typeof fileName !== 'string') return '';
    
    // 移除扩展名
    const nameWithoutExt = fileName.replace(/\.[^/.]+$/, '');
    
    // 尝试提取姓名（假设姓名在分隔符前）
    // 匹配中文姓名（2-4个汉字）或英文姓名
    const patterns = [
        // 匹配开头的2-4个汉字
        /^([\u4e00-\u9fa5]{2,4})/,
        // 匹配开头的中文姓名，后跟分隔符
        /^([\u4e00-\u9fa5]{2,4})[_\-\s]/,
        // 匹配开头的英文姓名
        /^([a-zA-Z\s]{2,20})/,
        // 匹配开头的英文姓名，后跟分隔符
        /^([a-zA-Z\s]{2,20})[_\-\s]/
    ];
    
    for (const pattern of patterns) {
        const match = nameWithoutExt.match(pattern);
        if (match && match[1]) {
            // 去除首尾空格
            const name = match[1].trim();
            if (name.length >= 2) {
                return name;
            }
        }
    }
    
    // 如果没有匹配到，返回文件名（不含扩展名）
    return nameWithoutExt;
};

// 从localStorage加载数据，实现数据持久化
const loadFromStorage = () => {
    const savedResults = localStorage.getItem('textCheckerResults');
    const savedCommonErrors = localStorage.getItem('textCheckerCommonErrors');
    
    return {
        results: savedResults ? JSON.parse(savedResults) : [],
        commonErrors: savedCommonErrors ? JSON.parse(savedCommonErrors) : {}
    };
};

// 保存数据到localStorage
const saveToStorage = () => {
    localStorage.setItem('textCheckerResults', JSON.stringify(results));
    localStorage.setItem('textCheckerCommonErrors', JSON.stringify(commonErrors));
};

// 清空所有数据
const clearAllData = () => {
    if (confirm('确定要清空所有数据吗？此操作不可恢复。')) {
        // 清空全局变量
        results.length = 0;
        Object.keys(commonErrors).forEach(key => delete commonErrors[key]);
        // 保存到localStorage
        saveToStorage();
        // 重新渲染表格
        renderStatsTable();
        alert('数据已清空');
    }
};

// 初始化数据
const { results, commonErrors } = loadFromStorage();

let fileName = '';
let lastModified = new Date().toLocaleString();
let currentSortField = 'timestamp'; // 默认按时间排序
let currentSortOrder = 'desc'; // 默认降序
document.addEventListener('DOMContentLoaded', function() {
    const originalText = document.getElementById('originalText');
    const copiedText = document.getElementById('copiedText');
    const checkBtn = document.getElementById('checkBtn');
    const pasteBtn = document.getElementById('pasteBtn');
    const accuracyElement = document.getElementById('accuracy');
    const replaceCount = document.getElementById('replaceCount');
    const deleteCount = document.getElementById('deleteCount');
    const insertCount = document.getElementById('insertCount');
    const originalComparison = document.getElementById('originalComparison');
    const copiedComparison = document.getElementById('copiedComparison');
    const clearDataBtn = document.getElementById('clearDataBtn');
    const reverseTableBtn = document.getElementById('reverseTableBtn');
    
    // 为比较文本区域添加样式
    originalComparison.style.overflow = 'auto';
    copiedComparison.style.overflow = 'auto';
    
    // 清空数据按钮事件
    clearDataBtn.addEventListener('click', clearAllData);
    
    // 表头排序功能
    const sortableHeaders = document.querySelectorAll('.sortable');
    sortableHeaders.forEach(header => {
        header.addEventListener('click', function() {
            const field = this.dataset.sort;
            
            // 如果点击的是当前排序字段，切换排序方向
            if (field === currentSortField) {
                currentSortOrder = currentSortOrder === 'asc' ? 'desc' : 'asc';
            } else {
                // 否则，设置新的排序字段和默认排序方向
                currentSortField = field;
                currentSortOrder = 'asc';
            }
            
            // 更新排序指示器
            updateSortIndicators();
            // 重新渲染表格
            renderStatsTable();
        });
    });
    
    // 更新排序指示器
    function updateSortIndicators() {
        sortableHeaders.forEach(header => {
            const indicator = header.querySelector('.sort-indicator');
            if (header.dataset.sort === currentSortField) {
                indicator.textContent = currentSortOrder === 'asc' ? '↑' : '↓';
            } else {
                indicator.textContent = '';
            }
        });
    }
    
    // 初始化排序指示器
    updateSortIndicators();
    
    // 列宽调整功能
    let resizing = false;
    let startX = 0;
    let startWidth = 0;
    let currentTh = null;
    
    const statsTable = document.getElementById('statsTable');
    const tableHeaders = statsTable.querySelectorAll('th');
    
    // 保存列宽的对象
    const columnWidths = {};
    
    // 初始化列宽
    function initColumnWidths() {
        tableHeaders.forEach((th, index) => {
            // 如果没有保存的宽度，使用当前计算宽度
            if (!columnWidths[th.dataset.sort || index]) {
                columnWidths[th.dataset.sort || index] = th.offsetWidth;
            }
        });
    }
    
    // 应用保存的列宽
    function applyColumnWidths() {
        tableHeaders.forEach((th, index) => {
            const key = th.dataset.sort || index;
            if (columnWidths[key]) {
                th.style.width = `${columnWidths[key]}px`;
            }
        });
    }
    
    // 鼠标按下事件 - 开始调整列宽
    statsTable.addEventListener('mousedown', function(e) {
        if (e.target.tagName === 'TH') {
            const th = e.target;
            // 检查是否点击在调整手柄上（右侧5px区域）
            const rect = th.getBoundingClientRect();
            const clickX = e.clientX - rect.left;
            
            if (clickX > rect.width - 5) {
                resizing = true;
                currentTh = th;
                startX = e.clientX;
                startWidth = th.offsetWidth;
                
                // 添加调整中的样式
                th.classList.add('resizing');
                document.body.classList.add('resizing');
                
                // 防止文本选择
                e.preventDefault();
            }
        }
    });
    
    // 鼠标移动事件 - 调整列宽
    document.addEventListener('mousemove', function(e) {
        if (resizing && currentTh) {
            const deltaX = e.clientX - startX;
            const newWidth = Math.max(50, startWidth + deltaX); // 最小宽度50px
            
            // 更新当前列宽
            currentTh.style.width = `${newWidth}px`;
            
            // 保存列宽
            const key = currentTh.dataset.sort || Array.from(tableHeaders).indexOf(currentTh);
            columnWidths[key] = newWidth;
        }
    });
    
    // 鼠标释放事件 - 结束调整列宽
    document.addEventListener('mouseup', function() {
        if (resizing) {
            if (currentTh) {
                currentTh.classList.remove('resizing');
            }
            document.body.classList.remove('resizing');
            resizing = false;
            currentTh = null;
        }
    });
    
    // 初始化列宽
    initColumnWidths();
    // 应用列宽
    applyColumnWidths();
    
    // 在表格重新渲染后应用列宽
    const originalRenderStatsTable = renderStatsTable;
    renderStatsTable = function() {
        originalRenderStatsTable();
        applyColumnWidths();
    };
    
    // 实现同步滚动功能 - 优化版
    function setupSyncScroll(element1, element2) {
        let isSyncing = false;
        let scrollTimeout = null;
        
        // 防抖函数
        function debounce(func, wait) {
            return function(...args) {
                clearTimeout(scrollTimeout);
                scrollTimeout = setTimeout(() => func.apply(this, args), wait);
            };
        }
        
        // 优化的滚动处理函数
        function handleScroll(source, target) {
            if (!isSyncing) {
                isSyncing = true;
                
                // 使用requestAnimationFrame使滚动更平滑
                requestAnimationFrame(() => {
                    const scrollRatio = target.scrollHeight / source.scrollHeight;
                    target.scrollTop = source.scrollTop * scrollRatio;
                    target.scrollLeft = source.scrollLeft;
                });
                
                setTimeout(() => { isSyncing = false; }, 30);
            }
        }
        
        // 添加防抖处理的滚动事件监听
        const debouncedScroll1 = debounce(() => handleScroll(element1, element2), 10);
        const debouncedScroll2 = debounce(() => handleScroll(element2, element1), 10);
        
        element1.addEventListener('scroll', debouncedScroll1);
        element2.addEventListener('scroll', debouncedScroll2);
    }
    
    // 设置输入文本区域的同步滚动
    setupSyncScroll(originalText, copiedText);
    // 设置对比文本区域的同步滚动
    setupSyncScroll(originalComparison, copiedComparison);
    
    // 检测按钮
    checkBtn.addEventListener('click', function() {
        const original = originalText.value;
        const copied = copiedText.value;
        
        if (!original || !copied) {
            alert('请确保目标文本和抄写文本都不为空');
            return;
        }
        
        const result = calculateAccuracy(original, copied);
        
        accuracyElement.textContent = `${result.accuracy.toFixed(2)}`;
        replaceCount.textContent = result.replaceCount;
        deleteCount.textContent = result.deleteCount;
        insertCount.textContent = result.insertCount;
        
        originalComparison.innerHTML = result.markedOriginal;
        copiedComparison.innerHTML = result.markedCopied;
        renderStatsTable();
    });
    // 粘贴按钮
    pasteBtn.addEventListener('click', async function() {
        try {
            const text = await navigator.clipboard.readText();
            copiedText.value = text;
        } catch (err) {
            alert('无法读取剪贴板: ' + err.message);
        }
    });
    // 处理单个DOCX文件导入的通用函数（返回Promise）
    function handleDocxFile(file) {
        return new Promise((resolve, reject) => {
            if (!file || !file.name.endsWith('.docx')) {
                reject(new Error('请选择DOCX格式的文件！'));
                return;
            }
            
            lastModified = new Date(file.lastModified);
            fileName = file.name.replace(/\.[^/.]+$/, "");
            
            const reader = new FileReader();
            reader.onload = function(e) {
                const arrayBuffer = e.target.result;
                mammoth.convertToHtml({arrayBuffer: arrayBuffer})
                    .then(function(result) {
                        // 获取转换后的HTML内容并清理危险标签
                        const htmlContent = sanitizeHtml(result.value);
                        
                        // 从HTML中提取纯文本内容，并保留换行符
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(htmlContent, 'text/html');
                        
                        // 简单直接的方式：将HTML转换为文本，保留换行符
                        // 首先获取所有文本节点
                        function extractTextWithNewlines(element) {
                            let text = '';
                            for (let child of element.childNodes) {
                                if (child.nodeType === Node.TEXT_NODE) {
                                    text += child.textContent;
                                } else if (child.nodeType === Node.ELEMENT_NODE) {
                                    // 对于块级元素，只在前面添加一个换行符
                                    const isBlockElement = ['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'table', 'tr', 'td', 'th', 'hr'].includes(child.tagName.toLowerCase());
                                    if (isBlockElement) {
                                        text += '\n';
                                    }
                                    text += extractTextWithNewlines(child);
                                    // 对于br标签，添加换行符
                                    if (child.tagName.toLowerCase() === 'br') {
                                        text += '\n';
                                    }
                                }
                            }
                            return text;
                        }
                        
                        let plainText = extractTextWithNewlines(doc.body);
                        
                        // 处理多余的换行符，保留最多2个连续换行符
                        plainText = plainText.replace(/\n{3,}/g, '\n\n');
                        // 移除首尾的空白字符
                        plainText = plainText.trim();
                        
                        // 保存原始文本（假设目标文本区域已有内容）
                        const original = originalText.value;
                        
                        if (!original) {
                            reject(new Error('目标文本区域不能为空！'));
                            return;
                        }
                        
                        // 将纯文本内容设置到抄写文本区域
                        document.getElementById('copiedText').value = plainText;
                        
                        // 运行检测
                        const detectionResult = calculateAccuracy(original, plainText);
                        
                        // 返回结果
                        resolve({
                            file: file,
                            result: detectionResult
                        });
                    })
                    .catch(function(error) {
                        console.error('转换错误:', error);
                        reject(new Error('导入DOCX文件时出错: ' + error.message));
                    });
            }

            reader.onerror = function() {
                reject(new Error('文件读取失败！'));
            };

            reader.readAsArrayBuffer(file);
        });
    }
    
    // 导入DOCX文件 - 点击方式（支持多文件）
    document.getElementById('importDocxBtn').addEventListener('click', function() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.docx';
        input.multiple = true; // 允许选择多个文件
        input.onchange = function(e) {
            const files = Array.from(e.target.files);
            if (files.length > 0) {
                processMultipleFiles(files);
            }
        };
        input.click();
    });
    
    // 导入DOCX文件 - 拖放方式
    const importDocxBtn = document.getElementById('importDocxBtn');
    
    // 防止浏览器默认拖放行为
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        importDocxBtn.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    // 添加拖放视觉反馈
    ['dragenter', 'dragover'].forEach(eventName => {
        importDocxBtn.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        importDocxBtn.addEventListener(eventName, unhighlight, false);
    });
    
    function highlight() {
        importDocxBtn.classList.add('drag-over');
    }
    
    function unhighlight() {
        importDocxBtn.classList.remove('drag-over');
    }
    
    // 处理拖放事件（支持多文件）
    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = Array.from(dt.files);
        
        if (files.length > 0) {
            processMultipleFiles(files);
        }
    }
    
    importDocxBtn.addEventListener('drop', handleDrop, false);
    
    // 处理多个文件的函数
    async function processMultipleFiles(files) {
        // 显示文件信息
        const fileInfo = document.getElementById('fileInfo');
        fileInfo.innerHTML = `正在处理 <strong>${files.length}</strong> 个文件...`;
        fileInfo.style.display = 'block';
        
        // 显示进度条
        const progressContainer = document.getElementById('progressContainer');
        const progressFill = document.getElementById('progressFill');
        const resultsList = document.getElementById('resultsList');
        
        progressContainer.style.display = 'block';
        resultsList.innerHTML = '';
        resultsList.style.display = 'block';
        
        let processedCount = 0;
        const results = [];
        
        // 逐个处理文件
        for (const file of files) {
            try {
                // 更新进度
                processedCount++;
                const progress = Math.round((processedCount / files.length) * 100);
                progressFill.style.width = `${progress}%`;
                progressFill.textContent = `${progress}%`;
                
                // 添加当前处理的文件到结果列表
                const resultItem = document.createElement('div');
                resultItem.className = 'result-item processing-status';
                resultItem.innerHTML = `
                    <div>
                        <strong>${escapeHtml(file.name)}</strong>
                        <span> - 处理中...</span>
                    </div>
                    <div class="status">处理中</div>
                `;
                resultItem.id = `result-${file.name}`;
                resultsList.appendChild(resultItem);
                
                // 处理单个文件
                const result = await handleDocxFile(file);
                results.push(result);
                
                // 更新结果项为成功
                const successItem = document.getElementById(`result-${file.name}`);
                successItem.className = 'result-item success';
                successItem.innerHTML = `
                    <div>
                        <strong>${escapeHtml(file.name)}</strong>
                        <span> - 正确率: ${result.result.accuracy.toFixed(2)}%</span>
                        <div class="file-results">
                            <div>替换错误: ${result.result.replaceCount}个</div>
                            <div>删除错误: ${result.result.deleteCount}个</div>
                            <div>插入错误: ${result.result.insertCount}个</div>
                        </div>
                    </div>
                    <div class="status processing-complete">完成</div>
                `;
                
            } catch (error) {
                console.error(`处理文件 ${file.name} 时出错:`, error);
                
                // 更新结果项为错误
                const errorItem = document.getElementById(`result-${file.name}`);
                if (errorItem) {
                    errorItem.className = 'result-item error';
                    errorItem.innerHTML = `
                        <div>
                            <strong>${escapeHtml(file.name)}</strong>
                            <span> - 错误: ${escapeHtml(error.message)}</span>
                        </div>
                        <div class="status processing-error">失败</div>
                    `;
                }
            }
        }
        
        // 所有文件处理完成
        const fileResults = [];
        fileInfo.innerHTML = `已完成处理 <strong>${files.length}</strong> 个文件，成功 <strong>${fileResults.length}</strong> 个，失败 <strong>${files.length - fileResults.length}</strong> 个。`;
        progressFill.style.width = '100%';
        progressFill.textContent = '100%';
        
        // 短暂延迟后隐藏进度条
        setTimeout(() => {
            progressContainer.style.display = 'none';
        }, 1000);
        
        // 重新渲染统计表格
        renderStatsTable();
        
        console.log('所有文件处理完成，结果:', fileResults);
    }

    // 使用Levenshtein距离算法计算文本差异
    function calculateAccuracy(original, copied) {
        const m = original.length;
        const n = copied.length;
        
        // 创建编辑距离矩阵
        const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));
        
        // 创建临时数组收集本次检测的错误文字
        const currentErrors = [];
        
        // 初始化边界条件
        for (let i = 0; i <= m; i++) {
            dp[i][0] = i;
        }
        for (let j = 0; j <= n; j++) {
            dp[0][j] = j;
        }
        
        // 填充编辑距离矩阵
        for (let i = 1; i <= m; i++) {
            for (let j = 1; j <= n; j++) {
                if (original[i - 1] === copied[j - 1]) {
                    dp[i][j] = dp[i - 1][j - 1];
                } else {
                    dp[i][j] = Math.min(
                        dp[i - 1][j] + 1,     // 删除
                        dp[i][j - 1] + 1,     // 插入
                        dp[i - 1][j - 1] + 1  // 替换
                    );
                }
            }
        }
        
        // 回溯操作序列
        let i = m;
        let j = n;
        const operations = [];
        
        while (i > 0 || j > 0) {
            if (i > 0 && j > 0 && original[i - 1] === copied[j - 1]) {
                operations.unshift('match');
                i--;
                j--;
            } else if (i > 0 && (j === 0 || dp[i][j] === dp[i - 1][j] + 1)) {
                operations.unshift('delete');
                i--;
            } else if (j > 0 && (i === 0 || dp[i][j] === dp[i][j - 1] + 1)) {
                operations.unshift('insert');
                j--;
            } else {
                operations.unshift('replace');
                i--;
                j--;
            }
        }
         // === 新增：合并连续操作 ===
        const mergedOps = [];
        let currentOp = null;
        let count = 0;

        for (const op of operations) {
            if (op === 'match') {
                if (currentOp && currentOp.type !== 'match') {
                    mergedOps.push({ ...currentOp });
                    currentOp = null;
               }
                mergedOps.push({ type: 'match' });
            } else {
                if (currentOp && currentOp.type === op) {
                    currentOp.count++;
                } else {
                    if (currentOp) mergedOps.push(currentOp);
                    currentOp = { type: op, count: 1 };
                }
            }
        }
        if (currentOp) mergedOps.push(currentOp);
        // === 生成标记文本 ===
        let markedOriginal = '';
        let markedCopied = '';
        let replaceCount = 0; // 实际上是 delete + insert 的组合，但这里我们按 insert/delete 统计
        let deleteCount = 0;
        let insertCount = 0;

        let originalIndex = 0; // 原始文本索引
        let copiedIndex = 0; // 抄写文本索引

        for (const op of mergedOps) {
            if (op.type === 'match') {
                markedOriginal += escapeHtml(original[originalIndex]);
                markedCopied += escapeHtml(copied[copiedIndex]);
                originalIndex++;
                copiedIndex++;
            } else if (op.type === 'delete') {
                // 连续少字
                const deletedText = original.slice(originalIndex, originalIndex + op.count);
                markedOriginal += `<span class="delete">${escapeHtml(deletedText)}</span>`;
                markedCopied += `<span class="delete">${escapeHtml(deletedText)}</span>`;
                deleteCount += op.count;
                originalIndex += op.count;
                
                // 收集删除错误到临时数组
                const errorKey = `[删除]${deletedText}`;
                currentErrors.push(errorKey);
            } else if (op.type === 'insert') {
                // 连续多字
                const insertedText = copied.slice(copiedIndex, copiedIndex + op.count);
                markedOriginal += `<span class="insert">${escapeHtml(insertedText)}</span>`;
                markedCopied += `<span class="insert">${escapeHtml(insertedText)}</span>`;
                insertCount += op.count;
                copiedIndex += op.count;
                
                // 收集插入错误到临时数组
                const errorKey = `[插入]${insertedText}`;
                currentErrors.push(errorKey);
            } else if (op.type === 'replace') {
                const originalText = original.slice(originalIndex, originalIndex + op.count);
                const replacedText = copied.slice(copiedIndex, copiedIndex + op.count);
                markedOriginal += `<span class="replace">${escapeHtml(originalText)}</span>`;
                markedCopied += `<span class="replace">${escapeHtml(replacedText)}</span>`;
                replaceCount += op.count;  // 单独统计替换
                originalIndex += op.count;
                copiedIndex += op.count;
                
                // 收集替换错误到临时数组
                const errorKey = `${originalText}→${replacedText}`;
                currentErrors.push(errorKey);
            }
        }
        // 计算正确率
        const totalErrors = replaceCount + deleteCount + insertCount;
        const accuracy = (( 1 - totalErrors / m )* 100);

        // 保存结果到数组 
        const result = {
            fileName: fileName,
            accuracy: accuracy.toFixed(2),
            replaceCount: replaceCount,
            deleteCount: deleteCount,
            insertCount: insertCount,
            timestamp: lastModified
        };
        
        // 检查是否已存在相同的记录（文件名、准确率、替换数、删除数、插入数都相同）
        const isDuplicate = results.some(existingResult => 
            existingResult.fileName === result.fileName &&
            existingResult.accuracy === result.accuracy &&
            existingResult.replaceCount === result.replaceCount &&
            existingResult.deleteCount === result.deleteCount &&
            existingResult.insertCount === result.insertCount
        );
        
        // 仅当不是重复记录时才添加到数组并更新错误文字统计
        if (!isDuplicate) {
            results.push(result);
            
            // 将本次检测的错误文字更新到commonErrors
            currentErrors.forEach(errorKey => {
                commonErrors[errorKey] = (commonErrors[errorKey] || 0) + 1;
            });
            
            // 保存数据到localStorage
            saveToStorage();
        }
        // 添加错误位置联动高亮功能 - 优化版
        function setupErrorHighlighting() {
            // 获取所有错误标记元素
            const originalErrors = Array.from(originalComparison.querySelectorAll('.replace, .delete, .insert'));
            const copiedErrors = Array.from(copiedComparison.querySelectorAll('.replace, .delete, .insert'));
            
            // 优化错误元素匹配逻辑
            if (originalErrors.length !== copiedErrors.length) {
                console.warn('错误元素数量不匹配，可能影响联动高亮效果');
            }
            
            // 存储当前高亮的错误元素
            let currentlyHighlighted = null;
            
            // 为每个错误元素添加事件监听器
            for (let i = 0; i < Math.min(originalErrors.length, copiedErrors.length); i++) {
                const originalError = originalErrors[i];
                const copiedError = copiedErrors[i];
                
                // 保存配对信息
                originalError.dataset.pairId = i;
                copiedError.dataset.pairId = i;
                
                // 创建共享的高亮函数
                function highlightPair(sourceError, targetError) {
                    // 移除之前的高亮
                    if (currentlyHighlighted) {
                        const [prevOrig, prevCopy] = currentlyHighlighted;
                        prevOrig.classList.remove('active');
                        prevCopy.classList.remove('active');
                    }
                    
                    // 添加新的高亮
                    sourceError.classList.add('active');
                    targetError.classList.add('active');
                    
                    // 保存当前高亮的元素对
                    currentlyHighlighted = [originalError, copiedError];
                    
                    // 优化滚动行为，避免频繁滚动
                    if (!isElementInViewport(targetError)) {
                        targetError.scrollIntoView({ 
                            behavior: 'smooth', 
                            block: 'center',
                            inline: 'nearest' 
                        });
                    }
                }
                
                // 为原始文本中的错误添加鼠标事件
                originalError.addEventListener('mouseenter', function() {
                    highlightPair(originalError, copiedError);
                });
                
                // 为抄写文本中的错误添加鼠标事件
                copiedError.addEventListener('mouseenter', function() {
                    highlightPair(copiedError, originalError);
                });
            }
            
            // 添加鼠标离开对比区域时清除高亮
            function clearHighlights() {
                if (currentlyHighlighted) {
                    const [orig, copy] = currentlyHighlighted;
                    orig.classList.remove('active');
                    copy.classList.remove('active');
                    currentlyHighlighted = null;
                }
            }
            
            // 添加鼠标离开事件监听器
            originalComparison.addEventListener('mouseleave', clearHighlights);
            copiedComparison.addEventListener('mouseleave', clearHighlights);
        }
        
        // 检测元素是否在视口中的函数
        function isElementInViewport(element) {
            const rect = element.getBoundingClientRect();
            const parentRect = element.parentElement.getBoundingClientRect();
            return (
                rect.top >= parentRect.top &&
                rect.bottom <= parentRect.bottom
            );
        };
        
        // 延迟设置错误高亮功能，确保DOM已更新
        setTimeout(setupErrorHighlighting, 100);
        
        return {
            accuracy,
            markedOriginal,
            markedCopied,
            replaceCount,
            deleteCount,
            insertCount
        };
    }
    
    // 获取最常见的错误文字
    function getMostCommonErrors(limit = 5) {
        // 将错误对象转换为数组并按出现次数排序
        const sortedErrors = Object.entries(commonErrors)
            .sort((a, b) => b[1] - a[1])
            .slice(0, limit);
        
        return sortedErrors.map(([error, count]) => `${error}: ${count}次`).join(', ');
    }
    
    function renderStatsTable(){
        const statsTableBody = document.getElementById('statsTableBody');
        const commonErrorsContainer = document.getElementById('commonErrorsContainer');
        
        // 清空现有内容
        statsTableBody.innerHTML = '';
        commonErrorsContainer.innerHTML = '';
        
        // 显示常见错误文字统计
        if (Object.keys(commonErrors).length > 0) {
            commonErrorsContainer.innerHTML = `
                <div class="common-errors-stats">
                    <strong>常见错误文字：</strong>
                    <span>${escapeHtml(getMostCommonErrors())}</span>
                </div>
            `;
        }

        // 根据当前排序字段和方向排序结果
        const displayResults = [...results].sort((a, b) => {
            let aVal, bVal;
            
            // 根据不同字段获取对应的值
            switch (currentSortField) {
                case 'fileName':
                    aVal = a.fileName.toLowerCase();
                    bVal = b.fileName.toLowerCase();
                    break;
                case 'accuracy':
                    // 准确率转换为数值进行比较
                    aVal = parseFloat(a.accuracy);
                    bVal = parseFloat(b.accuracy);
                    break;
                case 'timestamp':
                    // 时间戳直接比较
                    aVal = new Date(a.timestamp).getTime();
                    bVal = new Date(b.timestamp).getTime();
                    break;
                default:
                    aVal = a[currentSortField];
                    bVal = b[currentSortField];
            }
            
            // 比较值
            let comparison = 0;
            if (aVal < bVal) {
                comparison = -1;
            } else if (aVal > bVal) {
                comparison = 1;
            }
            
            // 根据排序方向调整结果
            return currentSortOrder === 'asc' ? comparison : comparison * -1;
        });

        displayResults.forEach((result, index) => {
                               const row = document.createElement('tr');
                               const name = extractNameFromFileName(result.fileName);
                               row.innerHTML = `
                                   <td>${escapeHtml(name)}</td>
                                   <td>${escapeHtml(result.accuracy)}</td>
                                   <td>${escapeHtml(String(result.replaceCount))}</td>
                                   <td>${escapeHtml(String(result.deleteCount))}</td>
                                   <td>${escapeHtml(String(result.insertCount))}</td>
                                   <td>${escapeHtml(formatTimestamp(result.timestamp))}</td>
                               `;

                               statsTableBody.appendChild(row);
                           });
    }
function formatTimestamp(timestamp) {
    if (!timestamp || isNaN(new Date(timestamp))) {
        return '无效时间';
    }
    return new Date(timestamp).toLocaleString();
}
    // 在页面加载完成后渲染统计表格
    window.addEventListener('load', renderStatsTable);
});