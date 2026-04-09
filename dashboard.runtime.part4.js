            graphic: [{
                type: "text",
                left: "center",
                top: "middle",
                style: {
                    text: "\u6682\u65e0\u53ef\u900f\u89c6\u7684\u57f9\u8bad\u6570\u636e",
                    fill: "#64748b",
                    fontSize: 16,
                    fontWeight: 500
                }
            }]
        };
    } else if (chartType === "pie") {
        option = {
            ...baseOption,
            tooltip: {
                trigger: "item"
            },
            legend: {
                type: "scroll",
                bottom: 0
            },
            series: [{
                name: pivotData.hasSplit
                    ? `${pivotData.xAxisLabel} / ${pivotData.yAxisLabel}`
                    : pivotData.xAxisLabel,
                type: "pie",
                radius: window.innerWidth < 768 ? ["30%", "62%"] : ["35%", "70%"],
                center: ["50%", "42%"],
                avoidLabelOverlap: true,
                itemStyle: {
                    borderRadius: 10,
                    borderColor: "#ffffff",
                    borderWidth: 2
                },
                label: {
                    formatter: "{b}\n{d}%"
                },
                data: pivotData.pieData
            }]
        };
    } else {
        option = {
            ...baseOption,
            tooltip: {
                trigger: "axis"
            },
            legend: pivotData.hasSplit ? {
                top: 0,
                type: "scroll"
            } : undefined,
            grid: {
                top: pivotData.hasSplit ? 56 : 24,
                right: 16,
                bottom: 24,
                left: 16,
                containLabel: true
            },
            xAxis: {
                type: "category",
                name: pivotData.xAxisLabel,
                boundaryGap: chartType === "bar",
                data: pivotData.categories,
                axisLabel: {
                    interval: window.innerWidth < 768 ? "auto" : 0,
                    hideOverlap: true
                }
            },
            yAxis: {
                type: "value",
                name: pivotData.metricLabel,
                minInterval: 1
            },
            series: pivotData.series.map((item) => ({
                name: pivotData.hasSplit ? item.name : pivotData.metricLabel,
                type: chartType,
                smooth: chartType === "line",
                symbolSize: chartType === "line" ? 8 : undefined,
                barMaxWidth: 44,
                emphasis: {
                    focus: "series"
                },
                data: item.data
            }))
        };
    }

    pivotChartInstance.setOption(option, true);
    pivotChartInstance.resize();
}

function renderAll() {
    updateFilterDropdowns();

    if (!validateSession()) {
        return;
    }

    const filteredRecords = getFilteredRecords();
    renderMetrics(filteredRecords);
    renderUploadList();
    renderAdminOverview();
    renderAdminList();
    renderLogs();
    renderFileTable();
    renderTrend(filteredRecords);
    applyScale();
    ui.loginBadge.textContent = state.session.isLoggedIn ? state.session.currentAdmin.slice(0, 1).toUpperCase() : "\u6559";

    if (activeAdminTab === "pivot") {
        window.setTimeout(() => {
            renderPivot();
            pivotChartInstance?.resize();
        }, 0);
    }
}

async function handleLogin(event) {
    event.preventDefault();
    ui.loginError.classList.add("hidden");

    const username = normalizeCredential(ui.loginUsername.value, true);
    const password = normalizeCredential(ui.loginPassword.value);

    try {
        const payload = await apiRequest("/login", {
            method: "POST",
            body: {
                username,
                password
            }
        });
        const matchedAdmin = normalizeAdminRecord(payload?.admin);

        state.session.isLoggedIn = true;
        state.session.currentAdmin = matchedAdmin?.username || normalizeCredential(payload?.admin?.username);
        state.session.authToken = toCellText(payload?.token);
        writeStoredSession();
        await loadState();
        closeLoginModal();
        renderAll();
        openAdminDrawer("overview");
        showToast("\u6559\u52a1\u540e\u53f0\u767b\u5f55\u6210\u529f");
    } catch (error) {
        ui.loginError.textContent = error.message || "\u4e91\u7aef\u767b\u5f55\u9a8c\u8bc1\u5931\u8d25\u3002";
        ui.loginError.classList.remove("hidden");
    }
}

function handleLogout() {
    state.session.isLoggedIn = false;
    state.session.currentAdmin = "";
    state.session.authToken = "";
    state.admins = [];
    state.logs = [];
    clearStoredSession();
    saveState();
    closeAdminDrawer();
    renderAll();

    if (suppressNextLogoutToast) {
        suppressNextLogoutToast = false;
        return;
    }

    showToast("\u5df2\u9000\u51fa\u6559\u52a1\u540e\u53f0");
}

async function handleCreateAdmin(event) {
    event.preventDefault();
    ui.adminCreateError.classList.add("hidden");

    const isSuperAdmin = normalizeCredential(state.session.currentAdmin, true) === normalizeCredential(DEFAULT_ADMIN_USERNAME, true);
    if (!isSuperAdmin) {
        ui.adminCreateError.textContent = "\u6743\u9650\u4e0d\u8db3\uff1a\u4ec5\u8d85\u7ea7\u7ba1\u7406\u5458\u53ef\u65b0\u589e\u8d26\u53f7\u3002";
        ui.adminCreateError.classList.remove("hidden");
        return;
    }

    const username = normalizeCredential(ui.newAdminUsername.value);
    const password = normalizeCredential(ui.newAdminPassword.value);

    if (!username || !password) {
        ui.adminCreateError.textContent = "\u8bf7\u5b8c\u6574\u8f93\u5165\u65b0\u7ba1\u7406\u5458\u8d26\u53f7\u548c\u5bc6\u7801\u3002";
        ui.adminCreateError.classList.remove("hidden");
        return;
    }

    const normalizedUsername = normalizeCredential(username, true);
    if (state.admins.some((admin) => admin.normalizedUsername === normalizedUsername)) {
        ui.adminCreateError.textContent = "\u8be5\u7ba1\u7406\u5458\u8d26\u53f7\u5df2\u5b58\u5728\u3002";
        ui.adminCreateError.classList.remove("hidden");
        return;
    }

    const snapshot = clonePersistedState();

    try {
        const payload = await apiRequest("/admins", {
            method: "POST",
            auth: true,
            body: {
                username,
                password
            }
        });
        const createdAdmin = normalizeAdminRecord(payload?.admin);
        state.admins = sortByDateDesc([
            createdAdmin,
            ...state.admins.filter((admin) => admin.id !== createdAdmin.id)
        ], "createdAt");
        await appendLogAndSync("\u65b0\u589e\u7ba1\u7406\u5458\u8d26\u53f7", `\u65b0\u589e\u7ba1\u7406\u5458 ${username}`, state.session.currentAdmin);
        ui.adminCreateForm.reset();
        renderAll();
        showToast("\u65b0\u589e\u7ba1\u7406\u5458\u8d26\u53f7\u6210\u529f");
    } catch (error) {
        restorePersistedState(snapshot);
        renderAll();
        ui.adminCreateError.textContent = error.message || "\u7ba1\u7406\u5458\u8d26\u53f7\u521b\u5efa\u5931\u8d25\u3002";
        ui.adminCreateError.classList.remove("hidden");
    }
}

function previewScale(nextValue) {
    if (state.scaleLocked) {
        ui.scaleSlider.value = String(state.uiScale);
        return;
    }

    applyScale(Number(nextValue) || 100);
}

async function handleScaleChange(nextValue) {
    if (state.scaleLocked) {
        ui.scaleSlider.value = String(state.uiScale);
        return;
    }

    const snapshot = clonePersistedState();
    state.uiScale = Number(nextValue) || 100;
    applyScale();

    try {
        await saveState();
    } catch (error) {
        restorePersistedState(snapshot);
        applyScale();
        showToast(error.message || "\u7f29\u653e\u8bbe\u7f6e\u4fdd\u5b58\u5931\u8d25", true);
    }
}

async function handleToggleScaleLock() {
    const snapshot = clonePersistedState();
    state.scaleLocked = !state.scaleLocked;
    applyScale();

    try {
        await appendLogAndSync(state.scaleLocked ? "\u9501\u5b9a\u754c\u9762\u7f29\u653e" : "\u89e3\u9501\u754c\u9762\u7f29\u653e", `\u5f53\u524d\u7f29\u653e ${state.uiScale}%`);
        renderAll();
    } catch (error) {
        restorePersistedState(snapshot);
        applyScale();
        showToast(error.message || "\u7f29\u653e\u9501\u5b9a\u4fdd\u5b58\u5931\u8d25", true);
    }
}

async function handleClearLogs() {
    if (!state.session.isLoggedIn) {
        showToast("\u8bf7\u5148\u767b\u5f55\u6559\u52a1\u540e\u53f0", true);
        return;
    }

    if (!window.confirm("\u786e\u8ba4\u6e05\u7a7a\u5168\u90e8\u53d8\u52a8\u65e5\u5fd7\uff1f")) {
        return;
    }

    const snapshot = clonePersistedState();
    state.logs = [];

    try {
        await apiRequest("/logs", {
            method: "DELETE",
            auth: true
        });
        renderAll();
        showToast("\u65e5\u5fd7\u5df2\u6e05\u7a7a");
    } catch (error) {
        restorePersistedState(snapshot);
        renderAll();
        showToast(error.message || "\u65e5\u5fd7\u6e05\u7406\u5931\u8d25", true);
    }
}

async function handleUploadFile(event) {
    const [file] = event.target.files || [];
    if (!file) {
        return;
    }

    if (!state.session.isLoggedIn) {
        showToast("\u8bf7\u5148\u767b\u5f55\u6559\u52a1\u540e\u53f0", true);
        event.target.value = "";
        return;
    }

    const snapshot = clonePersistedState();

    try {
        const records = await parseWorkbook(file);
        const fileId = createId("upload");
        records.forEach((record) => {
            record.uploadId = fileId;
            record.fileSize = file.size;
        });

        await apiRequest("/upload", {
            method: "POST",
            auth: true,
            body: {
                fileName: file.name,
                fileSize: file.size,
                records
            }
        });
        await loadState();
        renderAll();
        setAdminTab("data");
        showToast("\u57f9\u8bad\u6570\u636e\u5df2\u5bfc\u5165");
    } catch (error) {
        restorePersistedState(snapshot);
        renderAll();
        showToast(error.message || "\u57f9\u8bad\u6570\u636e\u5bfc\u5165\u5931\u8d25", true);
    } finally {
        event.target.value = "";
    }
}

function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
}

async function handleExportReport(exportType) {
    if (!state.session.isLoggedIn) {
        showToast("\u8bf7\u5148\u767b\u5f55\u6559\u52a1\u540e\u53f0", true);
        return;
    }

    if (!window.XLSX) {
        showToast("\u5bfc\u51fa\u7ec4\u4ef6\u672a\u52a0\u8f7d", true);
        return;
    }

    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    let sheetData = [];
    let sheetName = "";
    let fileSuffix = "";

    switch (exportType) {
    case "filtered": {
        const filteredRecords = getFilteredRecords();
        if (!filteredRecords.length) {
            showToast("\u5f53\u524d\u7b5b\u9009\u65e0\u6570\u636e", true);
            return;
        }
        sheetData = filteredRecords.map((record) => ({ ...record.detail }));
        sheetName = "\u7b5b\u9009\u6570\u636e";
        fileSuffix = "\u7b5b\u9009\u6570\u636e";
        break;
    }
    case "summary":
        sheetData = state.uploadedFiles.map((file) => ({
            "\u6587\u4ef6\u540d": file.fileName,
            "\u8bb0\u5f55\u6570": file.recordCount,
            "\u5bfc\u5165\u65f6\u95f4": formatDateTime(file.importedAt),
            "\u7ebf\u4e0b\u8bfe\u5802": file.counts?.["\u7ebf\u4e0b\u8bfe\u5802"] || 0,
            "\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b": file.counts?.["\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b"] || 0,
            "\u5546\u5b66\u9662": file.counts?.["\u5546\u5b66\u9662"] || 0,
            "\u672a\u5206\u7c7b": file.counts?.["\u672a\u5206\u7c7b"] || 0
        }));
        sheetName = "\u6c47\u603b";
        fileSuffix = "\u6c47\u603b";
        break;
    case "detail":
        sheetData = state.records.map((record) => ({
            "\u6587\u4ef6\u540d": record.sourceFile,
            "\u5de5\u4f5c\u8868": record.sheetName,
            "\u884c\u53f7": record.rowIndex,
            "\u5206\u7c7b": record.category,
            "\u5206\u7c7b\u539f\u503c": record.categorySource,
            "\u63d0\u53d6\u65f6\u95f4": record.extractedTime || "",
            "\u5bfc\u5165\u65f6\u95f4": formatDateTime(record.importedAt),
            ...record.detail
        }));
        sheetName = "\u660e\u7ec6";
        fileSuffix = "\u660e\u7ec6";
        break;
    case "log":
        sheetData = state.logs.map((log) => ({
            "\u64cd\u4f5c": log.action,
            "\u8be6\u60c5": log.detail,
            "\u6267\u884c\u4eba": log.actor,
            "\u65f6\u95f4": formatDateTime(log.createdAt)
        }));
        sheetName = "\u65e5\u5fd7";
        fileSuffix = "\u65e5\u5fd7";
        break;
    default:
        showToast("\u672a\u8bc6\u522b\u7684\u5bfc\u51fa\u7c7b\u578b", true);
        return;
    }

    try {
        const workbook = window.XLSX.utils.book_new();
        const normalizedSheetData = sheetData.length ? sheetData : [{ "\u63d0\u793a": "\u6682\u65e0\u6570\u636e" }];
        window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(normalizedSheetData), sheetName);
        window.XLSX.writeFile(workbook, `\u65f6\u4ee3\u8702\u65cf\u57f9\u8bad\u62a5\u8868_${fileSuffix}_${stamp}.xlsx`);

        await appendLogAndSync("\u5bfc\u51fa\u57f9\u8bad\u62a5\u8868", `\u5bfc\u51fa${sheetName} ${formatCount(sheetData.length)} \u6761`, state.session.currentAdmin);
        closeExportModal();
        renderAll();
        showToast(`${sheetName}\u5df2\u5bfc\u51fa`);
    } catch (error) {
        renderAll();
        showToast(error.message || "\u5bfc\u51fa\u62a5\u8868\u5931\u8d25", true);
    }
}

async function handleDeleteFile(fileId) {
    if (!state.session.isLoggedIn) {
        showToast("\u8bf7\u5148\u767b\u5f55\u6559\u52a1\u540e\u53f0", true);
        return;
    }

    const target = state.uploadedFiles.find((file) => file.id === fileId);
    if (!target) {
        showToast("\u672a\u627e\u5230\u5bf9\u5e94\u6587\u4ef6", true);
        return;
    }

    if (!window.confirm(`\u786e\u8ba4\u5220\u9664 ${target.fileName} \uff1f\u5220\u9664\u540e\u5173\u8054\u8bb0\u5f55\u4e5f\u4f1a\u4e00\u5e76\u79fb\u9664\u3002`)) {
        return;
    }

    const snapshot = clonePersistedState();

    try {
        await apiRequest(`/files/${encodeURIComponent(fileId)}`, {
            method: "DELETE",
            auth: true
        });
        await loadState();
        renderAll();
        showToast("\u57f9\u8bad\u6587\u4ef6\u5df2\u5220\u9664");
    } catch (error) {
        restorePersistedState(snapshot);
        renderAll();
        showToast(error.message || "\u6587\u4ef6\u5220\u9664\u5931\u8d25", true);
    }
}

async function handleDeleteAdmin(adminId) {
    const isSuperAdmin = normalizeCredential(state.session.currentAdmin, true) === normalizeCredential(DEFAULT_ADMIN_USERNAME, true);
    if (!isSuperAdmin) {
        showToast("\u6743\u9650\u4e0d\u8db3\uff1a\u4ec5\u8d85\u7ea7\u7ba1\u7406\u5458\u53ef\u5220\u9664\u8d26\u53f7\u3002", true);
        return;
    }

    const targetAdmin = state.admins.find((admin) => admin.id === adminId);
    if (!targetAdmin) {
        showToast("\u672a\u627e\u5230\u5bf9\u5e94\u7ba1\u7406\u5458\u8d26\u53f7", true);
        return;
    }

    if (targetAdmin.normalizedUsername === normalizeCredential(DEFAULT_ADMIN_USERNAME, true)) {
        showToast("\u8d85\u7ea7\u7ba1\u7406\u5458\u8d26\u53f7\u4e0d\u53ef\u5220\u9664", true);
        return;
    }

    const targetUsername = targetAdmin.username;
    if (!window.confirm(`\u786e\u5b9a\u8981\u5220\u9664\u7ba1\u7406\u5458 ${targetUsername} \u5417\uff1f\u6b64\u64cd\u4f5c\u4e0d\u53ef\u64a4\u9500\u3002`)) {
        return;
    }

    const snapshot = clonePersistedState();

    try {
        await apiRequest(`/admins/${encodeURIComponent(adminId)}`, {
            method: "DELETE",
            auth: true
        });
        state.admins = state.admins.filter((admin) => admin.id !== adminId);
        await loadState();
        renderAll();
        showToast("\u7ba1\u7406\u5458\u8d26\u53f7\u5df2\u79fb\u9664");
    } catch (error) {
        restorePersistedState(snapshot);
        renderAll();
        showToast(error.message || "\u7ba1\u7406\u5458\u8d26\u53f7\u5220\u9664\u5931\u8d25", true);
    }
}

function bindEvents() {
    let isAbsentCardScrolling = false;

    ui.openSidebar?.addEventListener("click", openSidebar);
    ui.closeSidebar?.addEventListener("click", closeSidebar);
    ui.sidebarBackdrop?.addEventListener("click", closeSidebar);
    ui.mobileOpenFilter?.addEventListener("click", openSidebar);
    ui.mobileOpenLogin?.addEventListener("click", () => {
        state.session.isLoggedIn ? openAdminDrawer("overview") : openLoginModal();
    });

    ui.resetFilters?.addEventListener("click", () => {
        ui.timeRange.value = DEFAULT_TIME_RANGE;
        ui.department.value = DEFAULT_DEPARTMENT;
        if (ui.filterBranch) {
            ui.filterBranch.value = DEFAULT_BRANCH;
        }
        if (ui.filterSalesDept) {
            ui.filterSalesDept.value = DEFAULT_SALES_DEPT;
        }
        if (ui.filterCourse) {
            ui.filterCourse.value = DEFAULT_COURSE;
        }
        renderAll();
    });

    ui.applyFilters?.addEventListener("click", () => {
        renderAll();
        if (window.innerWidth < 1024) {
            closeSidebar();
        }
    });

    ui.timeRange?.addEventListener("change", renderAll);
    ui.department?.addEventListener("change", renderAll);
    ui.filterBranch?.addEventListener("change", renderAll);
    ui.filterSalesDept?.addEventListener("change", renderAll);
    ui.filterCourse?.addEventListener("change", renderAll);
    ui.pivotXAxis?.addEventListener("change", renderPivot);
    ui.pivotYAxis?.addEventListener("change", renderPivot);
    ui.pivotChartType?.addEventListener("change", renderPivot);

    ui.openLogin?.addEventListener("click", () => {
        state.session.isLoggedIn ? openAdminDrawer("overview") : openLoginModal();
    });
    ui.closeLogin?.addEventListener("click", closeLoginModal);
    ui.loginModal?.addEventListener("click", (event) => {
        if (event.target === ui.loginModal) {
            closeLoginModal();
        }
    });
    ui.loginForm?.addEventListener("submit", handleLogin);

    ui.closeAdmin?.addEventListener("click", closeAdminDrawer);
    ui.logoutAdmin?.addEventListener("click", handleLogout);
    ui.adminTabs.forEach((button) => button.addEventListener("click", () => setAdminTab(button.dataset.tab)));

    ui.adminUploadTrigger?.addEventListener("click", () => ui.adminUploadInput?.click());
    ui.adminUploadInput?.addEventListener("change", handleUploadFile);
    ui.adminExportTrigger?.addEventListener("click", openExportModal);
    ui.closeExport?.addEventListener("click", closeExportModal);
    ui.exportModal?.addEventListener("click", async (event) => {
        if (event.target === ui.exportModal) {
            closeExportModal();
            return;
        }

        const button = event.target.closest(".export-option-btn");
        if (!button) {
            return;
        }

        await handleExportReport(button.dataset.type);
        closeExportModal();
    });

    ui.adminCreateForm?.addEventListener("submit", handleCreateAdmin);
    ui.adminList?.addEventListener("click", (event) => {
        const button = event.target.closest(".delete-admin-btn");
        if (button) {
            handleDeleteAdmin(button.dataset.adminId);
        }
    });
    ui.scaleSlider?.addEventListener("input", (event) => previewScale(event.target.value));
    ui.scaleSlider?.addEventListener("change", (event) => handleScaleChange(event.target.value));
    ui.toggleScaleLock?.addEventListener("click", handleToggleScaleLock);
    ui.clearLogs?.addEventListener("click", handleClearLogs);
    ui.fileTableBody?.addEventListener("click", (event) => {
        const button = event.target.closest(".delete-file-btn");
        if (button) {
            handleDeleteFile(button.dataset.fileId);
        }
    });
    ui.cardAbsent?.addEventListener("mouseenter", () => ui.absentTooltip?.classList.remove("hidden"));
    ui.cardAbsent?.addEventListener("mouseleave", () => ui.absentTooltip?.classList.add("hidden"));
    ui.cardAbsent?.addEventListener("touchmove", () => {
        isAbsentCardScrolling = true;
    }, { passive: true });
    ui.cardAbsent?.addEventListener("touchend", () => {
        if (!isAbsentCardScrolling) {
            ui.absentTooltip?.classList.toggle("hidden");
        }
        isAbsentCardScrolling = false;
    });
    document.addEventListener("click", (event) => {
        if (ui.cardAbsent && !ui.cardAbsent.contains(event.target)) {
            ui.absentTooltip?.classList.add("hidden");
        }
    });

    window.addEventListener("resize", () => {
        window.clearTimeout(resizeTimer);
        resizeTimer = window.setTimeout(() => {
            renderTrend(getFilteredRecords());
            if (activeAdminTab === "pivot") {
                renderPivot();
                pivotChartInstance?.resize();
            }
        }, 120);
    });

    document.addEventListener("visibilitychange", async () => {
        if (document.visibilityState !== "visible") {
            return;
        }

        try {
            await loadState();
            renderAll();
        } catch (error) {
            console.warn("接口数据刷新失败。", error);
        }
    });
}

async function bootstrap() {
    hydrateSessionFromStorage();
    bindEvents();
    try {
        await loadState();
    } catch (error) {
        showToast(error.message || "页面数据加载失败", true);
    }
    setAdminTab(activeAdminTab);
    renderAll();
}

bootstrap();
