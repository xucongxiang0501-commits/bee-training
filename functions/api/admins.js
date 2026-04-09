import {
    appendLog,
    createAdmin,
    errorResponse,
    getSuperAdminUsername,
    json,
    normalizeCredential,
    requireAuth
} from "../_lib/api.js";

export async function onRequestPost(context) {
    try {
        const currentAdmin = await requireAuth(context.request, context.env);
        const isSuperAdmin = normalizeCredential(currentAdmin.username, true) === normalizeCredential(getSuperAdminUsername(context.env), true);

        if (!isSuperAdmin) {
            return errorResponse("权限不足：仅超级管理员可新增账号。", 403);
        }

        const body = await context.request.json();
        const username = normalizeCredential(body?.username);
        const password = normalizeCredential(body?.password);

        if (!username || !password) {
            return errorResponse("请完整输入新管理员账号和密码。", 400);
        }

        const admin = await createAdmin(context.env, username, password);
        await appendLog(context.env, {
            action: "新增管理员账号",
            detail: `新增管理员 ${admin.username}`,
            actor: currentAdmin.username
        });

        return json({
            success: true,
            admin
        });
    } catch (error) {
        return errorResponse(error.message || "新增管理员失败", error.status || 500);
    }
}
