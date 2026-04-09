import {
    appendLog,
    deleteRecordsByUploadId,
    errorResponse,
    json,
    requireAuth,
    toCellText
} from "../../_lib/api.js";

export async function onRequestDelete(context) {
    try {
        const currentAdmin = await requireAuth(context.request, context.env);
        const uploadId = toCellText(context.params.id);

        if (!uploadId) {
            return errorResponse("缺少待删除文件标识。", 400);
        }

        const removedCount = await deleteRecordsByUploadId(context.env, uploadId);
        await appendLog(context.env, {
            action: "删除培训文件",
            detail: `删除文件分组 ${uploadId}，移除 ${removedCount} 条记录`,
            actor: currentAdmin.username
        });

        return json({
            success: true,
            removedCount
        });
    } catch (error) {
        return errorResponse(error.message || "删除培训文件失败", error.status || 500);
    }
}
