/**
 * 党政机关公文审核系统 — 前端逻辑
 * 流程：上传 → 审核 → 应用修订 → 下载
 */

(function () {
    "use strict";

    // ---- DOM 元素 ----
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const browseLink = document.getElementById("browse-link");
    const fileInfo = document.getElementById("file-info");
    const fileName = document.getElementById("file-name");
    const fileSize = document.getElementById("file-size");
    const removeFile = document.getElementById("remove-file");

    const enableRules = document.getElementById("enable-rules");
    const enableLlm = document.getElementById("enable-llm");
    const btnStart = document.getElementById("btn-start");
    const btnReset = document.getElementById("btn-reset");

    const progressSection = document.getElementById("progress-section");
    const progressFill = document.getElementById("progress-fill");
    const progressText = document.getElementById("progress-text");
    const statusDetail = document.getElementById("status-detail");

    const resultSection = document.getElementById("result-section");
    const statsRow = document.getElementById("stats-row");
    const issuesList = document.getElementById("issues-list");
    const downloadRow = document.getElementById("download-row");
    const btnDownload = document.getElementById("btn-download");

    const errorSection = document.getElementById("error-section");
    const errorMessage = document.getElementById("error-message");

    const queueSection = document.getElementById("queue-section");
    const queueTbody = document.getElementById("queue-tbody");

    // ---- 状态 ----
    let currentFile = null;
    let taskId = null;
    let pollTimer = null;
    let queuePollTimer = null;
    let taskStartTime = null;

    // ---- 工具函数 ----
    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
        return (bytes / (1024 * 1024)).toFixed(1) + " MB";
    }

    function formatElapsed(sec) {
        if (sec < 60) return sec + "秒";
        var m = Math.floor(sec / 60);
        var s = sec % 60;
        return m + "分" + s + "秒";
    }

    function elapsedText() {
        if (!taskStartTime) return "";
        var sec = Math.floor((Date.now() - taskStartTime) / 1000);
        return "（已用 " + formatElapsed(sec) + "）";
    }

    function show(el) { if (el) el.classList.remove("hidden"); }
    function hide(el) { if (el) el.classList.add("hidden"); }

    function resetAll() {
        currentFile = null;
        taskId = null;
        taskStartTime = null;
        if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
        if (queuePollTimer) { clearInterval(queuePollTimer); queuePollTimer = null; }
        hide(fileInfo);
        hide(progressSection);
        hide(resultSection);
        hide(errorSection);
        hide(btnReset);
        hide(queueSection);
        btnStart.disabled = true;
        progressFill.style.width = "0%";
        progressFill.classList.remove("indeterminate");
        progressText.textContent = "准备中...";
        statusDetail.textContent = "";
        fileInput.value = "";
    }

    function setProgress(value, text) {
        if (value < 0) {
            progressFill.classList.add("indeterminate");
            progressFill.style.width = "40%";
        } else {
            progressFill.classList.remove("indeterminate");
            progressFill.style.width = value + "%";
        }
        if (text) progressText.textContent = text;
    }

    // ---- 文件选择 ----
    function handleFile(file) {
        if (!file) return;
        if (!file.name.toLowerCase().endsWith(".docx")) {
            alert("仅支持 .docx 格式的文档");
            return;
        }
        currentFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        show(fileInfo);
        btnStart.disabled = false;
    }

    dropZone.addEventListener("click", function () {
        fileInput.click();
    });
    browseLink.addEventListener("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
    });
    fileInput.addEventListener("change", function () {
        handleFile(fileInput.files[0]);
    });
    removeFile.addEventListener("click", function (e) {
        e.preventDefault();
        currentFile = null;
        fileInput.value = "";
        hide(fileInfo);
        btnStart.disabled = true;
    });

    // 拖拽支持
    dropZone.addEventListener("dragover", function (e) {
        e.preventDefault();
        dropZone.classList.add("drag-over");
    });
    dropZone.addEventListener("dragleave", function (e) {
        e.preventDefault();
        dropZone.classList.remove("drag-over");
    });
    dropZone.addEventListener("drop", function (e) {
        e.preventDefault();
        dropZone.classList.remove("drag-over");
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // ---- 上传 & 审核 ----
    async function uploadFile() {
        const formData = new FormData();
        formData.append("file", currentFile);

        const resp = await fetch("/api/upload", {
            method: "POST",
            body: formData,
        });
        if (!resp.ok) {
            const err = await resp.json();
            throw new Error(err.detail || "上传失败");
        }
        return await resp.json();
    }

    async function startReviewAndApply() {
        if (!currentFile) return;

        try {
            btnStart.disabled = true;
            hide(errorSection);
            hide(resultSection);
            show(progressSection);
            setProgress(-1, "上传文档中...");

            // 1. 上传
            const uploadResult = await uploadFile();
            taskId = uploadResult.task_id;
            setProgress(10, "文档已上传，开始审核...");

            // 2. 审核 + 应用修订（一步完成）
            const resp = await fetch(`/api/review-and-apply/${taskId}`, {
                method: "POST",
            });
            if (!resp.ok) {
                const err = await resp.json();
                throw new Error(err.detail || "启动任务失败");
            }

            setProgress(20, "审核中...");
            startPolling();

        } catch (err) {
            showError(err.message);
        }
    }

    // ---- 轮询状态 ----
    function startPolling() {
        taskStartTime = Date.now();
        if (pollTimer) clearInterval(pollTimer);
        pollTimer = setInterval(pollStatus, 1000);
        // 同时轮询队列信息
        if (queuePollTimer) clearInterval(queuePollTimer);
        queuePollTimer = setInterval(pollQueue, 3000);
    }

    async function pollQueue() {
        try {
            const resp = await fetch("/api/queue");
            if (!resp.ok) return;
            const data = await resp.json();
            renderQueue(data);
        } catch (err) {
            // 静默
        }
    }

    function formatQueueTime(isoStr) {
        if (!isoStr) return "";
        const d = new Date(isoStr);
        const pad = function (n) { return n < 10 ? "0" + n : "" + n; };
        return pad(d.getHours()) + ":" + pad(d.getMinutes()) + ":" + pad(d.getSeconds());
    }

    function renderQueue(data) {
        var running = data.running;
        var queued = data.queued || [];
        var hasContent = running || queued.length > 0;

        if (!hasContent) {
            queueTbody.innerHTML = '<tr><td colspan="3" class="queue-empty">— 暂无排队 —</td></tr>';
            if (!running) hide(queueSection);
        } else {
            show(queueSection);
            var html = "";
            // 正在执行
            if (running) {
                var isMine = running.task_id === taskId;
                html += '<tr>';
                html += '<td class="queue-time">' + formatQueueTime(running.started_at || running.added_at) + '</td>';
                html += '<td class="queue-name">';
                html += '<span class="queue-badge running">执行中</span> ';
                html += escapeHtml(running.file_name);
                if (isMine) html += ' <em>（本任务）</em>';
                html += '</td>';
                html += '<td>—</td>';
                html += '</tr>';
            }
            // 排队中
            for (var i = 0; i < queued.length; i++) {
                var q = queued[i];
                var isMine = q.task_id === taskId;
                html += '<tr>';
                html += '<td class="queue-time">' + formatQueueTime(q.added_at) + '</td>';
                html += '<td class="queue-name">';
                html += '<span class="queue-badge">第' + q.position + '位</span> ';
                html += escapeHtml(q.file_name);
                if (isMine) html += ' <em>（本任务）</em>';
                html += '</td>';
                html += '<td>';
                if (isMine) {
                    html += '<button class="queue-cancel" onclick="cancelQueueTask(\'' + q.task_id + '\')">取消排队</button>';
                } else {
                    html += '—';
                }
                html += '</td>';
                html += '</tr>';
            }
            if (!running && queued.length === 0) {
                html = '<tr><td colspan="3" class="queue-empty">— 暂无排队 —</td></tr>';
            }
            queueTbody.innerHTML = html;
        }
    }

    window.cancelQueueTask = async function (id) {
        if (!confirm("确定要取消排队吗？任务将回退到审核完成状态。")) return;
        try {
            const resp = await fetch("/api/cancel/" + id, { method: "POST" });
            if (!resp.ok) {
                const err = await resp.json();
                alert(err.detail || "取消失败");
                return;
            }
            // 清掉 taskId 的 queued 状态，切到 reviewed，让用户重新发起
            // pollStatus 会自然拿到 reviewed 并展示
        } catch (e) {
            alert("网络错误，取消失败");
        }
    };

    async function pollStatus() {
        if (!taskId) return;

        try {
            const resp = await fetch(`/api/status/${taskId}`);
            if (!resp.ok) {
                throw new Error("查询状态失败");
            }
            const status = await resp.json();

            // 根据状态更新进度
            switch (status.status) {
                case "processing":
                    setProgress(40, "审核中...");
                    statusDetail.textContent = "正在执行格式规则检查和 LLM 内容审核..." +
                        elapsedText();
                    break;
                case "queued":
                    setProgress(45, "等待 WPS/Word 空闲...");
                    statusDetail.textContent = "修订任务已排队，等待 WPS/Word COM 资源空闲后自动执行..." +
                        elapsedText();
                    pollQueue();
                    break;
                case "applying":
                    if (status.apply_current > 0 && status.apply_total > 0) {
                        var ratio = status.apply_current / status.apply_total;
                        setProgress(Math.round(ratio * 100), "应用修订中...");
                        statusDetail.textContent =
                            status.apply_message +
                            " (" + status.apply_current + "/" + status.apply_total + ")" +
                            elapsedText();
                    } else {
                        setProgress(0, "应用修订中...");
                        statusDetail.textContent = "正在通过 WPS COM 以修订模式应用修改..." +
                            elapsedText();
                    }
                    break;
                case "reviewed":
                    setProgress(60, "审核完成，发现 " + status.total_issues + " 个问题");
                    statusDetail.textContent = status.error_message
                        ? status.error_message + " — 点击下方按钮重新应用修订"
                        : "点击「应用修订」生成修订后的文档";
                    if (queuePollTimer) { clearInterval(queuePollTimer); queuePollTimer = null; }
                    break;
                case "done":
                    setProgress(100, "完成！");
                    statusDetail.textContent = "修订后的文档已生成";
                    clearInterval(pollTimer);
                    pollTimer = null;
                    if (queuePollTimer) { clearInterval(queuePollTimer); queuePollTimer = null; }
                    hide(queueSection);
                    showResult(status);
                    break;
                case "error":
                    setProgress(0, "出错");
                    clearInterval(pollTimer);
                    pollTimer = null;
                    if (queuePollTimer) { clearInterval(queuePollTimer); queuePollTimer = null; }
                    hide(queueSection);
                    showError(status.error_message || "处理过程中出现未知错误");
                    break;
                default:
                    break;
            }
        } catch (err) {
            // 轮询失败不立即报错，继续尝试
            console.error("轮询失败:", err);
        }
    }

    // ---- 显示结果 ----
    function showResult(status) {
        show(resultSection);
        show(btnReset);
        hide(progressSection);

        // 统计卡片
        statsRow.innerHTML = "";
        const stats = [
            { label: "问题总数", value: status.total_issues, cls: "stat-total" },
            { label: "错误", value: status.error_count, cls: "stat-error" },
            { label: "警告", value: status.warning_count, cls: "stat-warning" },
            { label: "提示", value: status.info_count, cls: "stat-info" },
        ];
        stats.forEach(function (s) {
            const div = document.createElement("div");
            div.className = "stat-card " + s.cls;
            div.innerHTML = '<div class="stat-number">' + s.value + '</div>' +
                            '<div class="stat-label">' + s.label + '</div>';
            statsRow.appendChild(div);
        });

        // 问题列表
        issuesList.innerHTML = "";
        if (status.issues && status.issues.length > 0) {
            status.issues.forEach(function (issue) {
                issuesList.appendChild(renderIssue(issue));
            });
        } else {
            issuesList.innerHTML = '<p style="text-align:center;color:#666;">✅ 未发现问题</p>';
        }

        // 下载按钮
        if (status.download_url) {
            btnDownload.href = status.download_url;
            show(downloadRow);
        } else {
            hide(downloadRow);
        }
    }

    function renderIssue(issue) {
        const div = document.createElement("div");
        div.className = "issue-item severity-" + issue.severity;

        const badgeMap = { error: "错误", warning: "警告", info: "提示" };
        const badgeClsMap = { error: "badge-error", warning: "badge-warning", info: "badge-info" };
        const sourceMap = { rule: "格式规则", llm: "AI 审核" };

        let html = '<div class="issue-header">';
        html += '<span class="issue-badge ' + (badgeClsMap[issue.severity] || "badge-info") + '">' +
                (badgeMap[issue.severity] || issue.severity) + '</span>';
        html += '<span class="issue-title">' + escapeHtml(issue.title || "未命名问题") + '</span>';
        html += '<span class="issue-source">' + (sourceMap[issue.source] || issue.source) + '</span>';
        html += '</div>';

        if (issue.description) {
            html += '<div class="issue-desc">' + escapeHtml(issue.description) + '</div>';
        }
        if (issue.suggestion) {
            html += '<div class="issue-suggestion">建议：' + escapeHtml(issue.suggestion) + '</div>';
        }
        if (issue.original_text && issue.revised_text) {
            html += '<div class="issue-compare">' +
                    '<span class="orig">原文：' + escapeHtml(issue.original_text) + '</span>' +
                    '<span class="revised">修订：' + escapeHtml(issue.revised_text) + '</span>' +
                    '</div>';
        }

        div.innerHTML = html;
        return div;
    }

    function escapeHtml(str) {
        if (!str) return "";
        const div = document.createElement("div");
        div.textContent = str;
        return div.innerHTML;
    }

    function showError(msg) {
        hide(progressSection);
        show(errorSection);
        show(btnReset);
        errorMessage.textContent = msg;
        btnStart.disabled = false;
    }

    // ---- 事件绑定 ----
    btnStart.addEventListener("click", startReviewAndApply);
    btnReset.addEventListener("click", function () {
        resetAll();
    });

    // 初始状态
    resetAll();
})();