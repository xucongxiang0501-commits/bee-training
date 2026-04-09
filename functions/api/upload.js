import {
    appendLog,
    errorResponse,
    insertTrainingRecords,
    json,
    requireAuth,
    toCellText
} from "../_lib/api.js";

export async function onRequestPost(context) {
    try {
        const currentAdmin = await requireAuth(context.request, context.env);
        const body = await context.request.json();
        const records = Array.isArray(body?.records) ? body.records : [];

        if (!records.length) {
            return errorResponse("没有可上传的培训记录。", 400);
        }

        const result = await insertTrainingRecords(context.env, body);
        await appendLog(context.env, {
            action: "上传培训数据",
            detail: `导入 ${toCellText(body?.fileName) || "未命名文件"}，新增 ${records.length} 条记录`,
            actor: currentAdmin.username
        });

        return json({
            success: true,
            uploadId: result.uploadId,
            recordCount: result.recordCount,
            importedAt: result.importedAt
        });
    } catch (error) {
        return errorResponse(error.message || "上传培训数据失败", error.status || 500);
    }
}
