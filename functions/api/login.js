import {
    createSessionToken,
    errorResponse,
    findAdminByCredentials,
    json,
    normalizeAdmin,
    normalizeCredential,
    toCellText,
    updateAdminLastLogin
} from "../_lib/api.js";

export async function onRequestPost(context) {
    try {
        const body = await context.request.json();
        const username = normalizeCredential(body?.username, true);
        const password = normalizeCredential(body?.password);

        if (!username || !password) {
            return errorResponse("请完整输入账号和密码。", 400);
        }

        const adminRecord = await findAdminByCredentials(context.env, username, password);
        if (!adminRecord) {
            return errorResponse("账号或密码错误。", 401);
        }

        const normalizedAdmin = normalizeAdmin(adminRecord);
        const lastLoginAt = new Date().toISOString();
        await updateAdminLastLogin(context.env, normalizedAdmin.id, lastLoginAt);
        normalizedAdmin.lastLoginAt = lastLoginAt;

        return json({
            token: await createSessionToken(context.env, normalizedAdmin),
            admin: {
                ...normalizedAdmin,
                username: toCellText(normalizedAdmin.username)
            }
        });
    } catch (error) {
        return errorResponse(error.message || "登录失败", error.status || 500);
    }
}
