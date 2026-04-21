// 全局变量：存储 Word 拆分推荐的级别，防止分析后丢失
let wsplit_recommended = null; 

/**
 * 1. Toast 通知系统
 */
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    
    const icons = { success: '✓', error: '✕', info: 'ℹ', warning: '⚠' };
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span style="font-size:1.2rem;font-weight:bold">${icons[type] || icons.info}</span><span>${message}</span>`;
    
    container.appendChild(toast);
    setTimeout(() => { if (toast.parentNode) toast.remove(); }, 4000);
}

/**
 * 2. 终端日志输出
 */
eel.expose(update_terminal);
function update_terminal(msg) {
    const term = document.getElementById('terminal');
    if (!term) return;
    const div = document.createElement('div');
    div.innerText = `[${new Date().toLocaleTimeString()}] ${msg}`;
    term.appendChild(div);
    term.scrollTop = term.scrollHeight;
    
    // 成功/错误日志自动触发 Toast
    if (msg.includes('[√]') || msg.includes('成功')) {
        showToast(msg.replace('[√]', '').trim(), 'success');
    } else if (msg.includes('[x]') || msg.includes('失败') || msg.includes('错误')) {
        showToast(msg, 'error');
    }
}

/**
 * 3. 界面切换逻辑
 */
function switchView(viewId, element) {
    document.querySelectorAll('.tool-view').forEach(v => v.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    
    const targetView = document.getElementById('view-' + viewId);
    if (targetView) {
        targetView.classList.add('active');
        element.classList.add('active');
    }

    const titleMap = {
        'dashboard': ['运行概览', '系统状态与近期活动'],
        'rmblank': ['Word 空白页清理', '绝对物理切片，专治各种幽灵排版'],
        'pdfclean': ['扫描件去黑边', '基于 OpenCV 智能识别并遮盖扫描仪产生的黑边'], // ✨ 新增
        'p2w': ['PDF 提取 Word', '支持可编辑与纯图双模式'],
        'split': ['PDF 精准拆分', '定长、平均、全拆与范围提取'],
        'wsplit': ['Word 目录拆解', '按大纲级别一键拆分为独立文档'],
        'wmerge': ['Word 批量合并', '带 A 级剪贴板清理的静默合成'],
        'unlock': ['PDF 权限解密', '移除打印、复制及编辑限制'],
        'comp': ['文档极限瘦身', '二分法自动调参 PDF/Word 压缩引擎'],
        'ocr': ['PDF OCR 增强', '强制 OCR 重新扫描，生成透明文本层'],
        'i2p': ['图像转编 PDF', '自然排序递归打包'],
        'w2p': ['Word 转 PDF', 'Office 原生内核导出'],
        'p2i': ['PDF 转图片', 'Acrobat 换页对齐专用渲染'],
        'inv': ['发票自动提取', 'OCR 结构化对账提取'],
        'diff': ['文档差异比对', 'Word/Excel 双文本深度分析']
    };
    if (titleMap[viewId]) {
        document.getElementById('page-title').innerText = titleMap[viewId][0];
        document.getElementById('page-subtitle').innerText = titleMap[viewId][1];
    }
}

/**
 * 4. 基础文件/文件夹选择联动
 */
async function selectFile(id) {
    const p = await eel.ask_file()();
    if (p) {
        document.getElementById(id).value = p;
        showToast(`已选择: ${p.split('\\').pop()}`, 'info');
    }
}

async function selectFolder(id) {
    const p = await eel.ask_folder()();
    if (p) {
        document.getElementById(id).value = p;
        showToast(`已选择目录: ${p.split('\\').pop()}`, 'info');
    }
}

/**
 * 5. 核心联动：Word 拆分文件分析
 */
async function selectWordFileForSplit(id) {
    const p = await eel.ask_file()();
    if (p) {
        document.getElementById(id).value = p;
        
        const linkage = await eel.handle_file_selection(p)();
        if (linkage && linkage.out_dir) {
            document.getElementById('wsplit-out').value = linkage.out_dir;
        }

        const statusLbl = document.getElementById('wsplit-status');
        const levelCb = document.getElementById('wsplit-level');
        
        statusLbl.innerText = "🔍 正在分析大纲 (请稍候)...";
        statusLbl.style.color = "#8b949e";
        levelCb.innerHTML = '<option>分析中...</option>';
        
        try {
            const scanRes = await eel.get_word_outline(p)();
            if (scanRes && scanRes.options && scanRes.options.length > 0) {
                levelCb.innerHTML = "";
                scanRes.options.forEach(opt => {
                    const el = document.createElement("option");
                    el.value = opt; el.innerText = opt;
                    levelCb.appendChild(el);
                });
                statusLbl.innerText = scanRes.status_text;
                statusLbl.style.color = scanRes.status === "success" ? "#3fb950" : "#f85149";
                wsplit_recommended = scanRes.recommended;
            } else {
                statusLbl.innerText = "❌ 未识别到有效标题级别";
                levelCb.innerHTML = '<option value="1">手动尝试 1 级</option>';
            }
        } catch (err) {
            statusLbl.innerText = "❌ 引擎通讯失败";
            console.error(err);
        }
    }
}

/**
 * 6. 任务全量分发器
 */
async function execTask(type, btnElement) {
    let originalText = "执行任务";
    if (btnElement) {
        originalText = btnElement.innerText;
        btnElement.disabled = true;
        btnElement.innerHTML = `<span class="loading-spinner"></span> 引擎运转中...`;
        btnElement.style.cursor = "wait";
    }

    update_terminal(`[*] 正在启动 ${type} 核心引擎...`);
    let res;
    try {
        if (type === 'p2w') {
            res = await eel.run_pdf2word(
                document.getElementById('p2w-path').value, 
                document.getElementById('p2w-mode').value, 
                document.getElementById('p2w-dpi').value
            )();
        } else if (type === 'rmblank') {
            res = await eel.run_rm_blank(
                document.getElementById('rmblank-path').value
            )();
        } else if (type === 'pdfclean') {  // ✨ 新增执行分支
            res = await eel.run_pdf_cleaner(
                document.getElementById('pdfclean-path').value
            )();
        } else if (type === 'split') {
            const val = document.getElementById('split-val').value;
            res = await eel.run_split(
                document.getElementById('split-path').value, 
                document.getElementById('split-mode').value, 
                val, val, 1, 1
            )();
        } else if (type === 'wsplit') {
            res = await eel.run_word_split(
                document.getElementById('wsplit-path').value, 
                document.getElementById('wsplit-out').value, 
                document.getElementById('wsplit-level').value, 
                wsplit_recommended
            )();
        } else if (type === 'wmerge') {
            res = await eel.run_word_merge(
                document.getElementById('wmerge-dir').value, 
                document.getElementById('wmerge-name').value
            )();
        } else if (type === 'unlock') {
            res = await eel.run_unlock(
                document.getElementById('unlock-path').value, 
                document.getElementById('unlock-pwd').value, 
                document.getElementById('unlock-empty').checked, 
                document.getElementById('unlock-suffix').checked, 
                parseInt(document.getElementById('unlock-mode').value), 
                document.getElementById('unlock-retry').checked
            )();
        } else if (type === 'comp') {
            res = await eel.run_compress(
                document.getElementById('comp-path').value, 
                document.getElementById('comp-size').value, 
                document.getElementById('comp-unit').value
            )();
        } else if (type === 'ocr') {
            res = await eel.run_ocr(
                document.getElementById('ocr-path').value
            )();
        } else if (type === 'i2p') {
            res = await eel.run_img2pdf(
                document.getElementById('i2p-path').value, 
                true, true
            )();
        } else if (type === 'w2p') {
            res = await eel.run_word2pdf(
                document.getElementById('w2p-path').value, 
                true,
                document.getElementById('w2p-doc').checked, 
                document.getElementById('w2p-xls').checked, 
                "标题", 
                document.getElementById('w2p-pdfa').checked
            )();
        } else if (type === 'p2i') {
            res = await eel.run_pdf2img(
                document.getElementById('p2i-path').value, 
                document.getElementById('p2i-edge').value, 
                document.getElementById('p2i-q').value
            )();
        } else if (type === 'inv') {
            res = await eel.run_invoice(
                document.getElementById('inv-path').value, 
                document.getElementById('inv-rec').checked
            )();
        } else if (type === 'diff') {
            res = await eel.run_diff(
                document.getElementById('diff-f1').value, 
                document.getElementById('diff-f2').value, 
                document.getElementById('diff-strict').checked
            )();
        }

        if (res && res.status === "success") {
            update_terminal(`[√] 任务处理成功完成: ${res.msg || ''}`);
            showToast(res.msg || '任务执行成功', 'success');
        } else {
            update_terminal(`[!] 引擎返回失败: ${res ? res.msg : '未知错误'}`);
            showToast(res?.msg || '任务执行失败', 'error');
        }
    } catch (e) {
        const errorMsg = e.toString ? e.toString() : String(e);
        update_terminal(`[致命异常] ${errorMsg}`);
        showToast(errorMsg, 'error');
    } finally {
        if (btnElement) {
            btnElement.disabled = false;
            btnElement.innerText = originalText;
            btnElement.style.cursor = "pointer";
        }
    }
}