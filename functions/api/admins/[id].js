import {
    appendLog,
    deleteAdmin,
    errorResponse,
    findAdminById,
    getSuperAdminUsername,
    json,
    normalizeAdmin,
    normalizeCredential,
    requireAuth,
    toCellText
} from "../../_lib/api.js";

export async function onRequestDelete(context) {
    try {
        const currentAdmin = await requireAuth(context.request, context.env);
        const isSuperAdmin = normalizeCredential(currentAdmin.username, true) === normalizeCredential(getSuperAdminUsername(context.env), true);

        if (!isSuperAdmin) {
            return errorResponse("权限不足：仅超级管理员可删除账号。", 403);
        }

        const adminId = toCellText(context.params.id);
        const targetAdmin = normalizeAdmin(await findAdminById(context.env, adminId));
        if (!targetAdmin) {
            return errorResponse("未找到对应管理员账号。", 404);
        }

        if (normalizeCredential(targetAdmin.username, true) === normalizeCredential(getSuperAdminUsername(context.env), true)) {
            return errorResponse("超级管理员账号不可删除。", 400);
        }

        await deleteAdmin(context.env, adminId);
        await appendLog(context.env, {
            action: "删除管理员账号",
            detail: `删除了账号: ${targetAdmin.username}`,
            actor: currentAdmin.username
        });

        return json({
            success: true
        });
    } catch (error) {
        return errorResponse(error.message || "删除管理员失败", error.status || 500);
    }
}
