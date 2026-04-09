        : buildUploadedFilesFromRecords(state.records);
    state.admins = Array.isArray(payload.admins) ? payload.admins.map(normalizeAdminRecord).filter(Boolean) : [];
    state.logs = sortByDateDesc(Array.isArray(payload.logs) ? payload.logs.map(normalizeLogRecord) : [], "createdAt").slice(0, 200);
}

async function loadState() {
    try {
        const payload = await apiRequest("/state", {
            auth: state.session.isLoggedIn
        });
        applyStatePayload(payload);
        validateSession();
    } catch (error) {
        if (error.status === 401 && state.session.isLoggedIn) {
            clearStoredSession();
            state.session = { isLoggedIn: false, currentAdmin: "", authToken: "" };
            const fallbackPayload = await apiRequest("/state");
            applyStatePayload(fallbackPayload);
            showToast("登录状态已失效，请重新登录", true);
            return;
        }

        throw error;
    }
}

async function saveState() {
    writeStoredSession();
    return buildPersistedState();
}

function appendLog(action, detail, actor = state.session.currentAdmin || "\u7cfb\u7edf") {
    state.logs.unshift({
        id: createId("log"),
        action,
        detail,
        actor,
        createdAt: new Date().toISOString()
    });
    state.logs = sortByDateDesc(state.logs, "createdAt").slice(0, 200);
}

async function appendLogAndSync(action, detail, actor = state.session.currentAdmin || "系统") {
    const latestLog = {
        id: createId("log"),
        action,
        detail,
        actor,
        createdAt: new Date().toISOString()
    };
    state.logs.unshift(latestLog);
    state.logs = sortByDateDesc(state.logs, "createdAt").slice(0, 200);

    try {
        const payload = await apiRequest("/logs", {
            method: "POST",
            auth: true,
            body: {
                action: latestLog.action,
                detail: latestLog.detail,
                actor
            }
        });
        latestLog.id = payload?.id || latestLog.id;
        latestLog.createdAt = payload?.createdAt || latestLog.createdAt;
    } catch (error) {
        console.warn("日志同步失败，已仅保留当前页面内存状态。", error);
    }

    return latestLog;
}

function validateSession() {
    if (!state.session.isLoggedIn) {
        return true;
    }

    const currentAdminKey = normalizeCredential(state.session.currentAdmin, true);
    const currentAdminExists = state.admins.some((admin) => normalizeCredential(admin.username, true) === currentAdminKey);

    if (!currentAdminExists) {
        suppressNextLogoutToast = true;
        handleLogout();
        showToast("\u8d26\u53f7\u6743\u9650\u5df2\u53d8\u66f4\uff0c\u8bf7\u91cd\u65b0\u767b\u5f55", true);
        return false;
    }

    return true;
}

function syncBodyLock() {
    const sidebarOpen = ui.sidebar?.classList.contains("translate-x-0");
    const loginOpen = ui.loginModal && !ui.loginModal.classList.contains("hidden");
    const exportOpen = ui.exportModal && !ui.exportModal.classList.contains("hidden");
    document.body.classList.toggle("drawer-open", Boolean(sidebarOpen || loginOpen || exportOpen));
}

function openSidebar() {
    ui.sidebar?.classList.remove("-translate-x-full");
    ui.sidebar?.classList.add("translate-x-0");
    ui.sidebarBackdrop?.classList.remove("hidden");
    syncBodyLock();
}

function closeSidebar() {
    ui.sidebar?.classList.add("-translate-x-full");
    ui.sidebar?.classList.remove("translate-x-0");
    ui.sidebarBackdrop?.classList.add("hidden");
    syncBodyLock();
}

function openLoginModal() {
    ui.loginError?.classList.add("hidden");
    ui.loginModal?.classList.remove("hidden");
    ui.loginModal?.classList.add("flex");
    syncBodyLock();
}

function closeLoginModal() {
    ui.loginModal?.classList.add("hidden");
    ui.loginModal?.classList.remove("flex");
    ui.loginForm?.reset();
    ui.loginError?.classList.add("hidden");
    syncBodyLock();
}

function openExportModal() {
    ui.exportModal?.classList.remove("hidden");
    ui.exportModal?.classList.add("flex");
    syncBodyLock();
}

function closeExportModal() {
    ui.exportModal?.classList.add("hidden");
    ui.exportModal?.classList.remove("flex");
    syncBodyLock();
}

function openAdminDrawer(tab = activeAdminTab) {
    activeAdminTab = tab;
    setAdminTab(tab);
    ui.adminDrawer?.classList.remove("translate-x-full");
    ui.adminDrawer?.classList.add("translate-x-0");
}

function closeAdminDrawer() {
    ui.adminDrawer?.classList.add("translate-x-full");
    ui.adminDrawer?.classList.remove("translate-x-0");
}

function setAdminTab(tab) {
    activeAdminTab = tab;
    ui.adminTabs.forEach((button) => {
        const active = button.dataset.tab === tab;
        button.classList.toggle("admin-tab-active", active);
        button.classList.toggle("border-transparent", active);
        button.classList.toggle("border-slate-200", !active);
        button.classList.toggle("bg-white", !active);
        button.classList.toggle("text-slate-700", !active);
    });

    ui.adminPanels.forEach((panel) => {
        panel.classList.toggle("hidden-panel", panel.dataset.panel !== tab);
    });

    if (tab === "pivot") {
        window.setTimeout(() => {
            renderPivot();
            pivotChartInstance?.resize();
        }, 60);
    }
}

function applyScale(scaleValue = state.uiScale) {
    const normalizedScale = Number(scaleValue) || 100;
    const baseFontSize = Math.round((normalizedScale / 100) * 16);
    document.documentElement.style.fontSize = `${baseFontSize}px`;

    if (ui.scaleValue) {
        ui.scaleValue.textContent = `${normalizedScale}%`;
    }

    if (ui.scaleSlider) {
        ui.scaleSlider.value = String(normalizedScale);
        ui.scaleSlider.disabled = state.scaleLocked;
        ui.scaleSlider.classList.toggle("opacity-50", state.scaleLocked);
        ui.scaleSlider.classList.toggle("cursor-not-allowed", state.scaleLocked);
    }

    if (ui.toggleScaleLock) {
        ui.toggleScaleLock.textContent = state.scaleLocked ? "\u89e3\u9501\u7f29\u653e" : "\u9501\u5b9a\u7f29\u653e";
    }

    if (ui.scaleHint) {
        ui.scaleHint.textContent = state.scaleLocked
            ? "\u5f53\u524d\u5df2\u9501\u5b9a\u7f29\u653e\u6bd4\u4f8b\uff0c\u9700\u89e3\u9501\u540e\u624d\u80fd\u518d\u6b21\u8c03\u6574\u3002"
            : "\u672a\u9501\u5b9a\u65f6\u53ef\u62d6\u52a8\u8c03\u6574\u9875\u9762\u6bd4\u4f8b\uff0c\u9501\u5b9a\u540e\u4f1a\u963b\u6b62\u8fdb\u4e00\u6b65\u4fee\u6539\u3002";
    }
}

function getReferenceTimestamp(record) {
    return toTimestamp(record?.extractedTime) || toTimestamp(record?.importedAt);
}

function updateFilterDropdowns() {
    const rebuildSelect = (select, defaultValue, values) => {
        if (!select) {
            return;
        }

        const currentValue = select.value || defaultValue;
        const options = Array.from(new Set(
            values
                .map((value) => toCellText(value))
                .filter(Boolean)
        )).sort((left, right) => left.localeCompare(right, "zh-CN"));

        select.innerHTML = [
            `<option value="${escapeHtml(defaultValue)}">${escapeHtml(defaultValue)}</option>`,
            ...options.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
        ].join("");

        select.value = options.includes(currentValue) ? currentValue : defaultValue;
    };

    rebuildSelect(
        ui.filterBranch,
        DEFAULT_BRANCH,
        state.records.map((record) => record?.detail?.["\u516c\u53f8\u540d\u79f0"])
    );
    const currentBranch = ui.filterBranch ? ui.filterBranch.value : DEFAULT_BRANCH;
    const branchFilteredRecords = currentBranch === DEFAULT_BRANCH
        ? state.records
        : state.records.filter((record) => toCellText(record?.detail?.["\u516c\u53f8\u540d\u79f0"]) === currentBranch);

    rebuildSelect(
        ui.filterSalesDept,
        DEFAULT_SALES_DEPT,
        branchFilteredRecords.map((record) => record?.detail?.["\u8425\u4e1a\u90e8"])
    );
    rebuildSelect(
        ui.filterCourse,
        DEFAULT_COURSE,
        state.records.map((record) => record?.detail?.["\u8bfe\u5802\u540d\u79f0"])
    );
}

function getSelectedFilters() {
    return {
        timeRange: ui.timeRange?.value || DEFAULT_TIME_RANGE,
        department: ui.department?.value || DEFAULT_DEPARTMENT,
        branch: ui.filterBranch?.value || DEFAULT_BRANCH,
        salesDept: ui.filterSalesDept?.value || DEFAULT_SALES_DEPT,
        course: ui.filterCourse?.value || DEFAULT_COURSE
    };
}

function matchesTimeRange(record, timeRange) {
    if (timeRange === DEFAULT_TIME_RANGE) {
        return true;
    }

    const selectedMonthMatch = String(timeRange).match(/^(\d{1,2})\u6708$/);
    if (!selectedMonthMatch) {
        return true;
    }

    const referenceTimestamp = getReferenceTimestamp(record);
    if (!referenceTimestamp) {
        return false;
    }

    const selectedMonth = Number(selectedMonthMatch[1]);
    const recordMonth = new Date(referenceTimestamp).getMonth() + 1;
    return recordMonth === selectedMonth;
}

function getFilteredRecords() {
    const {
        timeRange,
        department,
        branch,
        salesDept,
        course
    } = getSelectedFilters();

    return state.records.filter((record) => {
        const detail = record.detail || {};
        const branchMatched = branch === DEFAULT_BRANCH || toCellText(detail["\u516c\u53f8\u540d\u79f0"]) === branch;
        const salesDeptMatched = salesDept === DEFAULT_SALES_DEPT || toCellText(detail["\u8425\u4e1a\u90e8"]) === salesDept;
        const courseMatched = course === DEFAULT_COURSE || toCellText(detail["\u8bfe\u5802\u540d\u79f0"]) === course;
        const departmentMatched = department === DEFAULT_DEPARTMENT || record.category === department;

        return departmentMatched
            && branchMatched
            && salesDeptMatched
            && courseMatched
            && matchesTimeRange(record, timeRange);
    });
}

function getPivotDimensionLabel(key) {
    switch (key) {
    case "month":
        return "\u6708\u4efd";
    case "category":
        return "\u7c7b\u522b";
    case "company":
        return "\u516c\u53f8\u540d\u79f0";
    case "salesDept":
        return "\u8425\u4e1a\u90e8";
    case "course":
        return "\u8bfe\u5802\u540d\u79f0";
    case "none":
    case "\u65e0":
        return "\u65e0";
    default:
        return "\u8bb0\u5f55\u6570";
    }
}

function getPivotDimensionValue(record, key) {
    const detail = record?.detail || {};

    switch (key) {
    case "month": {
        const timestamp = toTimestamp(record?.extractedTime);
        if (!timestamp) {
            return "\u672a\u8bc6\u522b\u65f6\u95f4";
        }

        return `${new Date(timestamp).getMonth() + 1}\u6708`;
    }
    case "category":
        return record?.category || "\u672a\u5206\u7c7b";
    case "company":
        return toCellText(detail["\u516c\u53f8\u540d\u79f0"]) || "\u672a\u586b\u5199";
    case "salesDept":
        return toCellText(detail["\u8425\u4e1a\u90e8"]) || "\u672a\u586b\u5199";
    case "course":
        return toCellText(detail["\u8bfe\u5802\u540d\u79f0"]) || "\u672a\u586b\u5199";
    case "none":
    case "\u65e0":
        return "\u8bb0\u5f55\u6570";
    default:
        return "\u672a\u586b\u5199";
    }
}

function sortPivotValues(values, dimensionKey) {
    const categoryOrder = ["\u7ebf\u4e0b\u8bfe\u5802", "\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b", "\u5546\u5b66\u9662", "\u901a\u5173", "\u672a\u5206\u7c7b", "\u672a\u586b\u5199"];

    return [...values].sort((left, right) => {
        if (dimensionKey === "month") {
            const parseMonth = (value) => {
                const matched = String(value).match(/^(\d{1,2})\u6708$/);
                return matched ? Number(matched[1]) : Number.POSITIVE_INFINITY;
            };
            const monthDiff = parseMonth(left) - parseMonth(right);
            if (monthDiff !== 0) {
                return monthDiff;
            }
        }

        if (dimensionKey === "category") {
            const leftIndex = categoryOrder.indexOf(String(left));
            const rightIndex = categoryOrder.indexOf(String(right));
            if (leftIndex !== rightIndex) {
                if (leftIndex === -1) {
                    return 1;
                }
                if (rightIndex === -1) {
                    return -1;
                }
                return leftIndex - rightIndex;
            }
        }

        return String(left).localeCompare(String(right), "zh-CN", {
            numeric: true,
            sensitivity: "base"
        });
    });
}

function getPivotData(records, xAxisKey, yAxisKey) {
    const normalizedYAxisKey = yAxisKey === "\u65e0" ? "none" : (yAxisKey || "none");
    const metricLabel = "\u8bb0\u5f55\u6570";
    const bucketMap = new Map();
    const seriesKeySet = new Set();

    records.forEach((record) => {
        const xValue = getPivotDimensionValue(record, xAxisKey);
        const yValue = normalizedYAxisKey === "none"
            ? metricLabel
            : getPivotDimensionValue(record, normalizedYAxisKey);

        if (!bucketMap.has(xValue)) {
            bucketMap.set(xValue, new Map());
        }

        const seriesMap = bucketMap.get(xValue);
        seriesMap.set(yValue, (seriesMap.get(yValue) || 0) + 1);
        seriesKeySet.add(yValue);
    });

    const categories = sortPivotValues(Array.from(bucketMap.keys()), xAxisKey);
    const seriesKeys = normalizedYAxisKey === "none"
        ? [metricLabel]
        : sortPivotValues(Array.from(seriesKeySet), normalizedYAxisKey);

    const series = seriesKeys.map((name) => ({
        name,
        data: categories.map((category) => bucketMap.get(category)?.get(name) || 0)
    }));

    const rows = categories.map((category) => {
        const values = Object.fromEntries(seriesKeys.map((name) => [name, bucketMap.get(category)?.get(name) || 0]));
        const total = seriesKeys.reduce((sum, name) => sum + (values[name] || 0), 0);

        return {
            category,
            values,
            total
        };
    });

    const columnTotals = Object.fromEntries(seriesKeys.map((name) => [
        name,
        categories.reduce((sum, category) => sum + (bucketMap.get(category)?.get(name) || 0), 0)
    ]));

    const pieData = normalizedYAxisKey === "none"
        ? rows
            .map((row) => ({ name: row.category, value: row.total }))
            .filter((item) => item.value > 0)
        : rows.flatMap((row) => {
            return seriesKeys
                .filter((name) => row.values[name] > 0)
                .map((name) => ({
                    name: `${row.category} \u00b7 ${name}`,
                    value: row.values[name]
                }));
        });

    return {
        categories,
        seriesKeys,
        series,
        rows,
        pieData,
        hasSplit: normalizedYAxisKey !== "none",
        xAxisLabel: getPivotDimensionLabel(xAxisKey),
        yAxisLabel: normalizedYAxisKey === "none" ? metricLabel : getPivotDimensionLabel(normalizedYAxisKey),
        metricLabel,
        columnTotals,
        grandTotal: rows.reduce((sum, row) => sum + row.total, 0)
    };
}

function toCellText(value) {
    if (value instanceof Date) {
        return value.toISOString();
    }
    return String(value ?? "").trim();
}

function detectCategory(firstCellValue) {
    const normalized = normalizeCredential(firstCellValue);
    return TRAINING_RULES[normalized] || "\u672a\u5206\u7c7b";
}

function extractTimeFromRow(headers, row) {
    for (let index = 0; index < headers.length; index += 1) {
        const header = toCellText(headers[index]);
        const matched = TIME_HEADER_KEYWORDS.some((keyword) => header.includes(keyword));
        if (matched) {
            const candidate = toCellText(row[index]);
            if (candidate) {
                return candidate;
            }
        }
    }

    const joined = row.map(toCellText).join(" ");
    const match = joined.match(/(\d{4}[\/\-\.年]\d{1,2}[\/\-\.月]\d{1,2}(?:日)?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)/);
    return match ? match[1] : "";
}

async function parseWorkbook(file) {
    if (!window.XLSX) {
        throw new Error("Excel \u89e3\u6790\u5e93\u672a\u52a0\u8f7d");
    }

    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, {
        type: "array",
        cellDates: true
    });
