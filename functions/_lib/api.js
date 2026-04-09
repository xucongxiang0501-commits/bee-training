const BMOB_BASE_URLS = [
    "https://api2.bmob.cn/1",
    "https://api.bmobapp.com/1"
];

const BMOB_TABLES = {
    admins: "Admins",
    records: "TrainingRecords",
    logs: "TrainingLogs"
};

const BMOB_PAGE_SIZE = 100;
const BMOB_BATCH_SIZE = 50;
const TOKEN_TTL_MS = 12 * 60 * 60 * 1000;
const DEFAULT_SUPER_ADMIN_USERNAME = "XUCONGXIANG";

export function json(data, init = {}) {
    const headers = new Headers(init.headers || {});
    headers.set("Content-Type", "application/json; charset=UTF-8");
    return new Response(JSON.stringify(data), {
        ...init,
        headers
    });
}

export function errorResponse(message, status = 400) {
    return json({ error: message }, { status });
}

export function normalizeCredential(value, lowerCase = false) {
    const normalized = String(value ?? "").trim().normalize("NFKC");
    return lowerCase ? normalized.toLowerCase() : normalized;
}

export function toCellText(value) {
    if (value instanceof Date) {
        return value.toISOString();
    }
    return String(value ?? "").trim();
}

export function toTimestamp(value) {
    if (!value) {
        return 0;
    }

    const date = new Date(value);
    const timestamp = date.getTime();
    return Number.isNaN(timestamp) ? 0 : timestamp;
}

function getRequiredEnv(env, key) {
    const value = env?.[key];
    if (!value) {
        throw new Error(`缺少 Cloudflare 环境变量 ${key}`);
    }
    return value;
}

function buildBmobHeaders(env) {
    return {
        "Content-Type": "application/json",
        "X-Bmob-Application-Id": getRequiredEnv(env, "BMOB_APP_ID"),
        "X-Bmob-REST-API-Key": getRequiredEnv(env, "BMOB_REST_API_KEY")
    };
}

function buildBmobUrl(baseUrl, path, query = {}) {
    const normalizedPath = path.startsWith("/") ? path : `/${path}`;
    const url = new URL(`${baseUrl}${normalizedPath}`);
    url.protocol = "https:";

    Object.entries(query).forEach(([key, value]) => {
        if (value === undefined || value === null || value === "") {
            return;
        }
        url.searchParams.set(key, typeof value === "object" ? JSON.stringify(value) : String(value));
    });

    return url.toString();
}

function normalizeBmobError(payload, status = 0, fallbackMessage = "Bmob 请求失败") {
    const error = new Error(payload?.error || payload?.msg || payload?.message || fallbackMessage);
    error.status = status;
    return error;
}

function isRetryableFetchError(error) {
    return /Failed to fetch|fetch failed|NetworkError|ERR_/i.test(String(error?.message || ""));
}

export async function bmobRequest(env, path, { method = "GET", query = {}, body } = {}) {
    let lastError = null;

    for (const baseUrl of BMOB_BASE_URLS) {
        try {
            const response = await fetch(buildBmobUrl(baseUrl, path, query), {
                method,
                headers: buildBmobHeaders(env),
                body: body === undefined ? undefined : JSON.stringify(body)
            });
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
                throw normalizeBmobError(payload, response.status);
            }

            return payload;
        } catch (error) {
            lastError = error instanceof Error ? error : new Error(String(error));
            if (!isRetryableFetchError(lastError)) {
                break;
            }
        }
    }

    throw lastError || new Error("无法连接 Bmob 服务");
}

export async function queryClass(env, tableName, where = null, options = {}) {
    const response = await bmobRequest(env, `/classes/${tableName}`, {
        query: {
            where: where && Object.keys(where).length ? where : undefined,
            limit: options.limit ?? BMOB_PAGE_SIZE,
            skip: options.skip ?? 0,
            order: options.order ?? "-createdAt"
        }
    });

    return Array.isArray(response?.results) ? response.results : [];
}

export async function fetchAllRows(env, tableName, where = null, options = {}) {
    const rows = [];
    let skip = 0;

    while (true) {
        const batch = await queryClass(env, tableName, where, {
            ...options,
            limit: BMOB_PAGE_SIZE,
            skip
        });
        rows.push(...batch);

        if (batch.length < BMOB_PAGE_SIZE) {
            break;
        }

        skip += BMOB_PAGE_SIZE;
    }

    return rows;
}

function chunkArray(items, size) {
    const chunks = [];
    for (let index = 0; index < items.length; index += size) {
        chunks.push(items.slice(index, index + size));
    }
    return chunks;
}

export async function batchRequests(env, requests) {
    const chunks = chunkArray(requests, BMOB_BATCH_SIZE);
    for (const chunk of chunks) {
        await bmobRequest(env, "/batch", {
            method: "POST",
            body: { requests: chunk }
        });
    }
}

export function normalizeAdmin(admin) {
    const username = normalizeCredential(admin?.username);
    if (!username) {
        return null;
    }

    return {
        id: admin?.objectId || admin?.id || crypto.randomUUID(),
        username,
        normalizedUsername: normalizeCredential(admin?.normalizedUsername || username, true),
        createdAt: admin?.createdAt || admin?.createdAtManual || new Date().toISOString(),
        lastLoginAt: admin?.lastLoginAt || ""
    };
}

export function normalizeRecord(record) {
    const detail = record?.recordData && Object.prototype.toString.call(record.recordData) === "[object Object]"
        ? record.recordData
        : (record?.detail && Object.prototype.toString.call(record.detail) === "[object Object]" ? record.detail : {});
    const categorySource = toCellText(record?.categorySource) || toCellText(detail["课堂类别"]);

    return {
        id: record?.objectId || record?.id || crypto.randomUUID(),
        uploadId: toCellText(record?.uploadId),
        sourceFile: toCellText(record?.sourceFile),
        fileSize: Number(record?.fileSize) || 0,
        sheetName: toCellText(record?.sheetName),
        rowIndex: Number(record?.rowIndex) || 0,
        category: toCellText(record?.category) || toCellText(categorySource) || "未分类",
        categorySource,
        extractedTime: toCellText(record?.extractedTime),
        importedAt: toCellText(record?.importedAt) || toCellText(record?.createdAt) || new Date().toISOString(),
        detail
    };
}

export function normalizeLog(log) {
    return {
        id: log?.objectId || log?.id || crypto.randomUUID(),
        action: toCellText(log?.action),
        detail: toCellText(log?.detail),
        actor: toCellText(log?.actor) || "系统",
        createdAt: toCellText(log?.createdAtManual) || toCellText(log?.createdAt) || new Date().toISOString()
    };
}

export function buildUploadedFilesFromRecords(records) {
    const summaryMap = new Map();

    records.forEach((record) => {
        const fileId = toCellText(record?.uploadId) || toCellText(record?.id) || crypto.randomUUID();
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

    return Array.from(summaryMap.values())
        .map((summary) => {
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
        })
        .sort((left, right) => toTimestamp(right.importedAt) - toTimestamp(left.importedAt));
}

export function getSuperAdminUsername(env) {
    return normalizeCredential(env?.SUPER_ADMIN_USERNAME || DEFAULT_SUPER_ADMIN_USERNAME);
}

export async function ensureSuperAdmin(env) {
    const username = getSuperAdminUsername(env);
    const password = normalizeCredential(env?.SUPER_ADMIN_PASSWORD);

    if (!username || !password) {
        return;
    }

    const existing = await queryClass(env, BMOB_TABLES.admins, {
        normalizedUsername: normalizeCredential(username, true)
    }, {
        limit: 1
    });

    if (existing.length) {
        return;
    }

    await bmobRequest(env, `/classes/${BMOB_TABLES.admins}`, {
        method: "POST",
        body: {
            username,
            normalizedUsername: normalizeCredential(username, true),
            password,
            lastLoginAt: ""
        }
    });
}

export async function findAdminByCredentials(env, username, password) {
    await ensureSuperAdmin(env);
    const rows = await queryClass(env, BMOB_TABLES.admins, {
        normalizedUsername: normalizeCredential(username, true),
        password: normalizeCredential(password)
    }, {
        limit: 1
    });

    return rows[0] || null;
}

export async function findAdminById(env, adminId) {
    if (!adminId) {
        return null;
    }

    try {
        return await bmobRequest(env, `/classes/${BMOB_TABLES.admins}/${adminId}`);
    } catch (error) {
        if (error.status === 404) {
            return null;
        }
        throw error;
    }
}

export async function updateAdminLastLogin(env, adminId, lastLoginAt) {
    await bmobRequest(env, `/classes/${BMOB_TABLES.admins}/${adminId}`, {
        method: "PUT",
        body: { lastLoginAt }
    });
}

export async function createAdmin(env, username, password) {
    const normalizedUsername = normalizeCredential(username, true);
    const existing = await queryClass(env, BMOB_TABLES.admins, {
        normalizedUsername
    }, {
        limit: 1
    });

    if (existing.length) {
        const error = new Error("该管理员账号已存在。");
        error.status = 409;
        throw error;
    }

    const payload = await bmobRequest(env, `/classes/${BMOB_TABLES.admins}`, {
        method: "POST",
        body: {
            username: normalizeCredential(username),
            normalizedUsername,
            password: normalizeCredential(password),
            lastLoginAt: ""
        }
    });

    return normalizeAdmin({
        objectId: payload?.objectId,
        username,
        normalizedUsername,
        createdAt: payload?.createdAt,
        lastLoginAt: ""
    });
}

export async function deleteAdmin(env, adminId) {
    await bmobRequest(env, `/classes/${BMOB_TABLES.admins}/${adminId}`, {
        method: "DELETE"
    });
}

export async function fetchRecords(env) {
    const rows = await fetchAllRows(env, BMOB_TABLES.records);
    return rows.map(normalizeRecord);
}

export async function fetchAdmins(env) {
    const rows = await fetchAllRows(env, BMOB_TABLES.admins);
    return rows.map(normalizeAdmin).filter(Boolean);
}

export async function fetchLogs(env) {
    const rows = await fetchAllRows(env, BMOB_TABLES.logs);
    return rows.map(normalizeLog).filter(Boolean);
}

function prepareTrainingRecord(record, { fileName, fileSize, uploadId, importedAt }) {
    return {
        uploadId,
        sourceFile: toCellText(record?.sourceFile) || fileName,
        fileSize: Number(record?.fileSize) || Number(fileSize) || 0,
        sheetName: toCellText(record?.sheetName),
        rowIndex: Number(record?.rowIndex) || 0,
        category: toCellText(record?.category) || "未分类",
        categorySource: toCellText(record?.categorySource),
        extractedTime: toCellText(record?.extractedTime),
        importedAt,
        recordData: record?.detail && Object.prototype.toString.call(record.detail) === "[object Object]" ? record.detail : {}
    };
}

export async function insertTrainingRecords(env, payload) {
    const fileName = toCellText(payload?.fileName) || "未命名文件";
    const fileSize = Number(payload?.fileSize) || 0;
    const uploadId = toCellText(payload?.uploadId) || crypto.randomUUID();
    const importedAt = new Date().toISOString();
    const records = Array.isArray(payload?.records) ? payload.records : [];

    const preparedRecords = records.map((record) => prepareTrainingRecord(record, {
        fileName,
        fileSize,
        uploadId,
        importedAt
    }));

    await batchRequests(env, preparedRecords.map((record) => ({
        method: "POST",
        path: `/1/classes/${BMOB_TABLES.records}`,
        body: record
    })));

    return {
        uploadId,
        recordCount: preparedRecords.length,
        importedAt
    };
}

export async function deleteRecordsByUploadId(env, uploadId) {
    const records = await fetchAllRows(env, BMOB_TABLES.records, {
        uploadId: toCellText(uploadId)
    });

    if (!records.length) {
        return 0;
    }

    await batchRequests(env, records.map((record) => ({
        method: "DELETE",
        path: `/1/classes/${BMOB_TABLES.records}/${record.objectId}`
    })));

    return records.length;
}

export async function appendLog(env, { action, detail, actor }) {
    const createdAt = new Date().toISOString();
    const payload = await bmobRequest(env, `/classes/${BMOB_TABLES.logs}`, {
        method: "POST",
        body: {
            action: toCellText(action),
            detail: toCellText(detail),
            actor: toCellText(actor) || "系统",
            createdAtManual: createdAt
        }
    });

    return {
        id: payload?.objectId || crypto.randomUUID(),
        action: toCellText(action),
        detail: toCellText(detail),
        actor: toCellText(actor) || "系统",
        createdAt: payload?.createdAt || createdAt
    };
}

export async function clearLogs(env) {
    const logs = await fetchAllRows(env, BMOB_TABLES.logs);
    if (!logs.length) {
        return;
    }

    await batchRequests(env, logs.map((log) => ({
        method: "DELETE",
        path: `/1/classes/${BMOB_TABLES.logs}/${log.objectId}`
    })));
}

function base64UrlEncodeText(value) {
    const bytes = new TextEncoder().encode(value);
    let binary = "";
    bytes.forEach((byte) => {
        binary += String.fromCharCode(byte);
    });
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlDecodeText(value) {
    const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    const binary = atob(padded);
    const bytes = new Uint8Array(binary.length);

    for (let index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
    }

    return new TextDecoder().decode(bytes);
}

async function signToken(env, payloadSegment) {
    const key = await crypto.subtle.importKey(
        "raw",
        new TextEncoder().encode(getRequiredEnv(env, "AUTH_SECRET")),
        { name: "HMAC", hash: "SHA-256" },
        false,
        ["sign"]
    );
    const signatureBuffer = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payloadSegment));
    const bytes = new Uint8Array(signatureBuffer);
    let binary = "";
    bytes.forEach((byte) => {
        binary += String.fromCharCode(byte);
    });
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

export async function createSessionToken(env, admin) {
    const now = Date.now();
    const payloadSegment = base64UrlEncodeText(JSON.stringify({
        sub: admin.id,
        username: admin.username,
        exp: now + TOKEN_TTL_MS
    }));
    const signature = await signToken(env, payloadSegment);
    return `${payloadSegment}.${signature}`;
}

export async function verifySessionToken(env, token) {
    if (!token || !token.includes(".")) {
        return null;
    }

    const [payloadSegment, signature] = token.split(".");
    const expectedSignature = await signToken(env, payloadSegment);
    if (signature !== expectedSignature) {
        return null;
    }

    let payload;
    try {
        payload = JSON.parse(base64UrlDecodeText(payloadSegment));
    } catch (error) {
        return null;
    }

    if (!payload?.sub || Number(payload?.exp) < Date.now()) {
        return null;
    }

    return payload;
}

function readBearerToken(request) {
    const authorization = request.headers.get("Authorization") || "";
    const match = authorization.match(/^Bearer\s+(.+)$/i);
    return match ? match[1].trim() : "";
}

export async function requireAuth(request, env) {
    const token = readBearerToken(request);
    const payload = await verifySessionToken(env, token);

    if (!payload) {
        const error = new Error("未授权访问");
        error.status = 401;
        throw error;
    }

    const admin = normalizeAdmin(await findAdminById(env, payload.sub));
    if (!admin) {
        const error = new Error("登录状态已失效");
        error.status = 401;
        throw error;
    }

    return admin;
}

export async function getOptionalAuth(request, env) {
    const token = readBearerToken(request);
    if (!token) {
        return null;
    }

    return requireAuth(request, env);
}

export async function fetchDashboardState(env, includeAdminData = false) {
    const records = await fetchRecords(env);
    const payload = {
        records,
        uploadedFiles: buildUploadedFilesFromRecords(records)
    };

    if (includeAdminData) {
        payload.admins = await fetchAdmins(env);
        payload.logs = await fetchLogs(env);
    } else {
        payload.admins = [];
        payload.logs = [];
    }

    return payload;
}
