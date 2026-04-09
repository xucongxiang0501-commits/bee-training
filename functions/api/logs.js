import {
    appendLog,
    clearLogs,
    errorResponse,
    json,
    requireAuth,
    toCellText
} from "../_lib/api.js";

export async function onRequestPost(context) {
    try {
        const currentAdmin = await requireAuth(context.request, context.env);
        const body = await context.request.json();
        const log = await appendLog(context.env, {
            action: toCellText(body?.action),
            detail: toCellText(body?.detail),
            actor: currentAdmin.username
        });

        return json(log);
    } catch (error) {
        return errorResponse(error.message || "日志写入失败", error.status || 500);
    }
}

export async function onRequestDelete(context) {
    try {
        await requireAuth(context.request, context.env);
        await clearLogs(context.env);
        return json({ success: true });
    } catch (error) {
        return errorResponse(error.message || "日志清理失败", error.status || 500);
    }
}
