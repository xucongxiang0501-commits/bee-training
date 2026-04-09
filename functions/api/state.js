import {
    errorResponse,
    fetchDashboardState,
    getOptionalAuth,
    json
} from "../_lib/api.js";

export async function onRequestGet(context) {
    try {
        const currentAdmin = await getOptionalAuth(context.request, context.env);
        const payload = await fetchDashboardState(context.env, Boolean(currentAdmin));
        return json(payload);
    } catch (error) {
        return errorResponse(error.message || "状态数据加载失败", error.status || 500);
    }
}
