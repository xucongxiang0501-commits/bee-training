const TRAINING_RULES = {
    "\u7ebf\u4e0b\u8bfe\u5802": "\u7ebf\u4e0b\u8bfe\u5802",
    "\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b": "\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b",
    "\u5546\u5b66\u9662": "\u5546\u5b66\u9662",
    "\u901a\u5173": "\u901a\u5173"
};

const TIME_HEADER_KEYWORDS = [
    "\u65f6\u95f4",
    "\u65e5\u671f",
    "\u57f9\u8bad\u65f6\u95f4",
    "\u4e0a\u8bfe\u65f6\u95f4",
    "\u5f00\u59cb\u65f6\u95f4",
    "\u5f00\u8bfe\u65f6\u95f4",
    "\u521b\u5efa\u65f6\u95f4",
    "\u63d0\u4ea4\u65f6\u95f4"
];

const DEFAULT_TIME_RANGE = "\u5168\u90e8\u65f6\u95f4";
const DEFAULT_DEPARTMENT = "\u5168\u90e8\u7c7b\u522b";
const DEFAULT_BRANCH = "\u5168\u90e8\u516c\u53f8";
const DEFAULT_SALES_DEPT = "\u5168\u90e8\u8425\u4e1a\u90e8";
const DEFAULT_COURSE = "\u5168\u90e8\u8bfe\u5802";
const DEFAULT_ADMIN_USERNAME = "XUCONGXIANG";
const API_BASE = "/api";
const SESSION_STORAGE_KEY = "bee-training-dashboard-session-v1";

const ui = {
    sidebarBackdrop: document.getElementById("sidebar-backdrop"),
    sidebar: document.getElementById("sidebar"),
    openSidebar: document.getElementById("open-sidebar"),
    closeSidebar: document.getElementById("close-sidebar"),
    timeRange: document.getElementById("time-range"),
    department: document.getElementById("department"),
    filterBranch: document.getElementById("filter-branch"),
    filterSalesDept: document.getElementById("filter-sales-dept"),
    filterCourse: document.getElementById("filter-course"),
    resetFilters: document.getElementById("reset-filters"),
    applyFilters: document.getElementById("apply-filters"),
    healthAttendance: document.getElementById("health-attendance"),
    healthAlerts: document.getElementById("health-alerts"),
    openLogin: document.getElementById("open-login"),
    mobileOpenFilter: document.getElementById("mobile-open-filter"),
    mobileOpenLogin: document.getElementById("mobile-open-login"),
    loginBadge: document.getElementById("login-badge"),
    loginModal: document.getElementById("login-modal"),
    closeLogin: document.getElementById("close-login"),
    exportModal: document.getElementById("export-modal"),
    closeExport: document.getElementById("close-export"),
    loginForm: document.getElementById("login-form"),
    loginUsername: document.getElementById("login-username"),
    loginPassword: document.getElementById("login-password"),
    loginError: document.getElementById("login-error"),
    adminDrawer: document.getElementById("admin-drawer"),
    closeAdmin: document.getElementById("close-admin"),
    logoutAdmin: document.getElementById("logout-admin"),
    adminTabs: Array.from(document.querySelectorAll(".admin-tab")),
    adminPanels: Array.from(document.querySelectorAll(".admin-panel")),
    adminCurrentUser: document.getElementById("admin-current-user"),
    adminOverviewSummary: document.getElementById("admin-overview-summary"),
    adminOverviewTime: document.getElementById("admin-overview-time"),
    adminUploadTrigger: document.getElementById("admin-upload-trigger"),
    adminUploadInput: document.getElementById("admin-upload-input"),
    adminExportTrigger: document.getElementById("admin-export-trigger"),
    adminImportSummary: document.getElementById("admin-import-summary"),
    adminImportBadge: document.getElementById("admin-import-badge"),
    adminImportDetails: document.getElementById("admin-import-details"),
    fileManagementCount: document.getElementById("file-management-count"),
    fileTableBody: document.getElementById("file-table-body"),
    fileTableEmpty: document.getElementById("file-table-empty"),
    pivotXAxis: document.getElementById("pivot-x-axis"),
    pivotYAxis: document.getElementById("pivot-y-axis"),
    pivotChartType: document.getElementById("pivot-chart-type"),
    pivotChartContainer: document.getElementById("pivot-chart-container"),
    pivotTableContainer: document.getElementById("pivot-table-container"),
    adminCountLabel: document.getElementById("admin-count-label"),
    adminList: document.getElementById("admin-list"),
    adminCreateForm: document.getElementById("admin-create-form"),
    adminPermissionNotice: document.getElementById("admin-permission-notice"),
    newAdminUsername: document.getElementById("new-admin-username"),
    newAdminPassword: document.getElementById("new-admin-password"),
    adminCreateError: document.getElementById("admin-create-error"),
    scaleValue: document.getElementById("scale-value"),
    scaleSlider: document.getElementById("scale-slider"),
    scaleHint: document.getElementById("scale-hint"),
    toggleScaleLock: document.getElementById("toggle-scale-lock"),
    clearLogs: document.getElementById("clear-logs"),
    logList: document.getElementById("log-list"),
    uploadBadge: document.getElementById("upload-badge"),
    uploadList: document.getElementById("upload-list"),
    metricTotal: document.getElementById("metric-total"),
    metricTotalNote: document.getElementById("metric-total-note"),
    metricTotalBadge: document.getElementById("metric-total-badge"),
    metricTotalInsight: document.getElementById("metric-total-insight"),
    metricAttended: document.getElementById("metric-attended"),
    metricAttendedNote: document.getElementById("metric-attended-note"),
    metricAttendedBadge: document.getElementById("metric-attended-badge"),
    metricAttendedInsight: document.getElementById("metric-attended-insight"),
    metricAbsent: document.getElementById("metric-absent"),
    metricAbsentNote: document.getElementById("metric-absent-note"),
    metricAbsentBadge: document.getElementById("metric-absent-badge"),
    metricAbsentInsight: document.getElementById("metric-absent-insight"),
    cardAbsent: document.getElementById("card-absent"),
    absentTooltip: document.getElementById("absent-tooltip"),
    absentList: document.getElementById("absent-list"),
    metricRate: document.getElementById("metric-rate"),
    metricRateNote: document.getElementById("metric-rate-note"),
    metricRateBadge: document.getElementById("metric-rate-badge"),
    metricRateInsight: document.getElementById("metric-rate-insight"),
    trendSvg: document.getElementById("trend-svg"),
    toast: document.getElementById("toast")
};

const state = {
    admins: [],
    uploadedFiles: [],
    records: [],
    logs: [],
    uiScale: 100,
    scaleLocked: false,
    session: {
        isLoggedIn: false,
        currentAdmin: "",
        authToken: ""
    }
};

let activeAdminTab = "overview";
let toastTimer = null;
let resizeTimer = null;
let suppressNextLogoutToast = false;
let pivotChartInstance = null;

function sortByDateDesc(list, key) {
    return [...list].sort((left, right) => toTimestamp(right?.[key]) - toTimestamp(left?.[key]));
}

function createId(prefix) {
    return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function toTimestamp(value) {
    if (!value) {
        return 0;
    }

    const date = new Date(value);
    const timestamp = date.getTime();
    return Number.isNaN(timestamp) ? 0 : timestamp;
}

function formatDateTime(value) {
    const timestamp = toTimestamp(value);
    if (!timestamp) {
        return "\u672a\u77e5\u65f6\u95f4";
    }

    return new Intl.DateTimeFormat("zh-CN", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit"
    }).format(new Date(timestamp)).replace(/\//g, "-");
}

function formatMonthLabel(value) {
    const timestamp = toTimestamp(value);
    const date = timestamp ? new Date(timestamp) : new Date();
    return `${date.getMonth() + 1} \u6708`;
}

function formatCompactDate(value) {
    const timestamp = toTimestamp(value);
    if (!timestamp) {
        return "\u672a\u77e5";
    }

    return new Intl.DateTimeFormat("zh-CN", {
        month: "2-digit",
        day: "2-digit"
    }).format(new Date(timestamp)).replace(/\//g, "-");
}

function formatCount(value) {
    return new Intl.NumberFormat("zh-CN").format(value || 0);
}

function normalizeCredential(value, lowerCase = false) {
    const normalized = String(value ?? "").trim().normalize("NFKC");
    return lowerCase ? normalized.toLowerCase() : normalized;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function humanFileSize(bytes) {
    const size = Number(bytes) || 0;
    if (size < 1024) {
        return `${size} B`;
    }

    if (size < 1024 * 1024) {
        return `${(size / 1024).toFixed(1)} KB`;
    }

    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function showToast(message, isError = false) {
    if (!ui.toast) {
        return;
    }

    ui.toast.textContent = message;
    ui.toast.classList.remove("translate-y-4", "opacity-0", "bg-ink", "bg-rose-700");
    ui.toast.classList.add("translate-y-0", "opacity-100", isError ? "bg-rose-700" : "bg-ink");

    window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => {
        ui.toast.classList.remove("translate-y-0", "opacity-100");
        ui.toast.classList.add("translate-y-4", "opacity-0");
    }, 2600);
}

function readStoredSession() {
    try {
        const raw = window.sessionStorage.getItem(SESSION_STORAGE_KEY);
        if (!raw) {
            return null;
        }

        const parsed = JSON.parse(raw);
        if (!parsed?.authToken || !parsed?.currentAdmin) {
            return null;
        }

        return {
            authToken: String(parsed.authToken),
            currentAdmin: normalizeCredential(parsed.currentAdmin)
        };
    } catch (error) {
        return null;
    }
}

function writeStoredSession() {
    try {
        if (!state.session.isLoggedIn || !state.session.authToken || !state.session.currentAdmin) {
            window.sessionStorage.removeItem(SESSION_STORAGE_KEY);
            return;
        }

        window.sessionStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify({
            authToken: state.session.authToken,
            currentAdmin: state.session.currentAdmin
        }));
    } catch (error) {
        console.warn("会话信息写入失败。", error);
    }
}

function clearStoredSession() {
    try {
        window.sessionStorage.removeItem(SESSION_STORAGE_KEY);
    } catch (error) {
        console.warn("会话信息清理失败。", error);
    }
}

function hydrateSessionFromStorage() {
    const storedSession = readStoredSession();
    if (!storedSession) {
        return;
    }

    state.session.isLoggedIn = true;
    state.session.currentAdmin = storedSession.currentAdmin;
    state.session.authToken = storedSession.authToken;
}

function normalizeAdminRecord(admin) {
    const username = normalizeCredential(admin?.username);
    if (!username) {
        return null;
    }

    return {
        id: admin?.id || admin?.objectId || createId("admin"),
        username,
        normalizedUsername: normalizeCredential(admin?.normalizedUsername || username, true),
        createdAt: admin?.createdAt || new Date().toISOString(),
        lastLoginAt: admin?.lastLoginAt || ""
    };
}

function normalizeApiRecord(record) {
    const detail = record?.detail && Object.prototype.toString.call(record.detail) === "[object Object]"
        ? record.detail
        : (record?.recordData && Object.prototype.toString.call(record.recordData) === "[object Object]" ? record.recordData : {});
    const categorySource = toCellText(record?.categorySource) || toCellText(detail["课堂类别"]);

    return {
        id: record?.id || record?.objectId || createId("record"),
        uploadId: toCellText(record?.uploadId),
        sourceFile: toCellText(record?.sourceFile),
        fileSize: Number(record?.fileSize) || 0,
        sheetName: toCellText(record?.sheetName),
        rowIndex: Number(record?.rowIndex) || 0,
        category: toCellText(record?.category) || detectCategory(categorySource),
        categorySource,
        extractedTime: toCellText(record?.extractedTime),
        importedAt: toCellText(record?.importedAt) || toCellText(record?.createdAt) || new Date().toISOString(),
        detail
    };
}

function normalizeLogRecord(log) {
    return {
        id: log?.id || log?.objectId || createId("log"),
        action: toCellText(log?.action),
        detail: toCellText(log?.detail),
        actor: toCellText(log?.actor) || "系统",
        createdAt: toCellText(log?.createdAt) || new Date().toISOString()
    };
}

function normalizeUploadedFile(file) {
    return {
        id: toCellText(file?.id) || createId("upload"),
        fileName: toCellText(file?.fileName) || "未命名文件",
        fileSize: Number(file?.fileSize) || 0,
        importedAt: toCellText(file?.importedAt) || new Date().toISOString(),
        sheetCount: Number(file?.sheetCount) || 0,
        recordCount: Number(file?.recordCount) || 0,
        counts: file?.counts && Object.prototype.toString.call(file.counts) === "[object Object]" ? file.counts : {},
        timeSummary: file?.timeSummary && Object.prototype.toString.call(file.timeSummary) === "[object Object]" ? file.timeSummary : { start: "", end: "" }
    };
}

function buildPersistedState() {
    return {
        admins: state.admins,
        uploadedFiles: state.uploadedFiles,
        records: state.records,
        logs: state.logs,
        uiScale: state.uiScale,
        scaleLocked: state.scaleLocked,
        session: state.session
    };
}

function clonePersistedState() {
    return JSON.parse(JSON.stringify(buildPersistedState()));
}

function restorePersistedState(snapshot = {}) {
    state.admins = Array.isArray(snapshot.admins) ? snapshot.admins.map(normalizeAdminRecord).filter(Boolean) : [];
    state.uploadedFiles = sortByDateDesc(Array.isArray(snapshot.uploadedFiles) ? snapshot.uploadedFiles.map(normalizeUploadedFile) : [], "importedAt");
    state.records = sortByDateDesc(Array.isArray(snapshot.records) ? snapshot.records.map(normalizeApiRecord) : [], "importedAt");
    state.logs = sortByDateDesc(Array.isArray(snapshot.logs) ? snapshot.logs.map(normalizeLogRecord) : [], "createdAt");
    state.uiScale = Math.max(85, Math.min(120, Number(snapshot.uiScale) || 100));
    state.scaleLocked = Boolean(snapshot.scaleLocked);
    state.session = snapshot.session?.isLoggedIn
        ? {
            isLoggedIn: true,
            currentAdmin: normalizeCredential(snapshot.session.currentAdmin),
            authToken: toCellText(snapshot.session.authToken)
        }
        : { isLoggedIn: false, currentAdmin: "", authToken: "" };
}

function buildApiUrl(path, query = {}) {
    if (window.location.protocol === "file:") {
        throw new Error("当前版本不能直接通过 file:// 打开，请使用 Cloudflare Pages 或本地 Functions 开发服务");
    }

    const normalizedPath = path.startsWith("/") ? path : `/${path}`;
    const url = new URL(`${API_BASE}${normalizedPath}`, window.location.origin);
    Object.entries(query).forEach(([key, value]) => {
        if (value === undefined || value === null || value === "") {
            return;
        }
        url.searchParams.set(key, typeof value === "object" ? JSON.stringify(value) : String(value));
    });
    return url.toString();
}

function normalizeApiError(payload, status = 0, fallbackMessage = "接口请求失败") {
    const message = payload?.error || payload?.message || payload?.detail || fallbackMessage;
    const error = new Error(message);
    error.status = status;
    return error;
}

async function apiRequest(path, { method = "GET", body, auth = false, query = {} } = {}) {
    const headers = {
        "Content-Type": "application/json"
    };

    if (auth && state.session.authToken) {
        headers.Authorization = `Bearer ${state.session.authToken}`;
    }

    let response;

    try {
        response = await fetch(buildApiUrl(path, query), {
            method,
            headers,
            body: body === undefined ? undefined : JSON.stringify(body)
        });
    } catch (error) {
        throw new Error("无法连接站点接口，请确认 Cloudflare Functions 已部署或本地开发服务已启动");
    }

    const rawText = await response.text();
    let payload = {};

    if (rawText) {
        try {
            payload = JSON.parse(rawText);
        } catch (error) {
            payload = { message: rawText };
        }
    }

    if (!response.ok) {
        throw normalizeApiError(payload, response.status);
    }

    return payload;
}

function buildUploadedFilesFromRecords(records) {
    const summaryMap = new Map();

    records.forEach((record) => {
        const fileId = toCellText(record?.uploadId) || toCellText(record?.id) || createId("upload");
        if (!summaryMap.has(fileId)) {
            summaryMap.set(fileId, {
                id: fileId,
                fileName: toCellText(record?.sourceFile) || "云端导入",
                fileSize: Number(record?.fileSize) || 0,
                importedAt: toCellText(record?.importedAt) || new Date().toISOString(),
                sheetNames: new Set(),
                recordCount: 0,
                counts: {
                    "线下课堂": 0,
                    "自主培训课程": 0,
                    "商学院": 0,
                    "通关": 0,
                    "未分类": 0
                },
                timestamps: []
            });
        }

        const summary = summaryMap.get(fileId);
        summary.fileSize = Math.max(summary.fileSize, Number(record?.fileSize) || 0);
        summary.recordCount += 1;
        summary.sheetNames.add(toCellText(record?.sheetName) || "未命名工作表");
        summary.importedAt = toTimestamp(record?.importedAt) > toTimestamp(summary.importedAt) ? record.importedAt : summary.importedAt;
        summary.counts[record.category] = (summary.counts[record.category] || 0) + 1;

        const timestamp = toTimestamp(record?.extractedTime);
        if (timestamp) {
            summary.timestamps.push(timestamp);
        }
    });

    return sortByDateDesc(Array.from(summaryMap.values()).map((summary) => {
        summary.timestamps.sort((left, right) => left - right);
        return {
            id: summary.id,
            fileName: summary.fileName,
            fileSize: summary.fileSize,
            importedAt: summary.importedAt,
            sheetCount: summary.sheetNames.size,
            recordCount: summary.recordCount,
            counts: summary.counts,
            timeSummary: {
                start: summary.timestamps[0] ? new Date(summary.timestamps[0]).toISOString() : "",
                end: summary.timestamps[summary.timestamps.length - 1] ? new Date(summary.timestamps[summary.timestamps.length - 1]).toISOString() : ""
            }
        };
    }), "importedAt");
}

function applyStatePayload(payload = {}) {
    const records = Array.isArray(payload.records) ? payload.records.map(normalizeApiRecord) : [];
    state.records = sortByDateDesc(records, "importedAt");
    state.uploadedFiles = Array.isArray(payload.uploadedFiles) && payload.uploadedFiles.length
        ? sortByDateDesc(payload.uploadedFiles.map(normalizeUploadedFile), "importedAt")
