    const records = [];

    workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const rows = window.XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
            blankrows: false,
            raw: false,
            dateNF: "yyyy-mm-dd hh:mm:ss"
        });

        if (!rows.length) {
            return;
        }

        const headers = rows[0].map((item, index) => toCellText(item) || `\u5217${index + 1}`);
        rows.slice(1).forEach((row, rowIndex) => {
            const hasData = row.some((cell) => toCellText(cell));
            if (!hasData) {
                return;
            }

            const detail = {};
            headers.forEach((header, index) => {
                detail[header] = toCellText(row[index]);
            });

            const categoryIndex = headers.indexOf("\u8bfe\u5802\u7c7b\u522b");
            const categoryCell = categoryIndex !== -1 ? toCellText(row[categoryIndex]) : "";
            records.push({
                id: createId("record"),
                uploadId: "",
                sourceFile: file.name,
                sheetName,
                rowIndex: rowIndex + 2,
                category: detectCategory(categoryCell),
                categorySource: categoryCell,
                extractedTime: extractTimeFromRow(headers, row),
                importedAt: new Date().toISOString(),
                detail
            });
        });
    });

    return records;
}

function buildFileSummary(fileId, file, records) {
    const counts = {
        "\u7ebf\u4e0b\u8bfe\u5802": 0,
        "\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b": 0,
        "\u5546\u5b66\u9662": 0,
        "\u672a\u5206\u7c7b": 0
    };

    const timestamps = [];
    records.forEach((record) => {
        counts[record.category] = (counts[record.category] || 0) + 1;
        const timestamp = toTimestamp(record.extractedTime);
        if (timestamp) {
            timestamps.push(timestamp);
        }
    });

    timestamps.sort((left, right) => left - right);
    const start = timestamps[0] ? new Date(timestamps[0]).toISOString() : "";
    const end = timestamps[timestamps.length - 1] ? new Date(timestamps[timestamps.length - 1]).toISOString() : "";

    return {
        id: fileId,
        fileName: file.name,
        fileSize: file.size,
        importedAt: new Date().toISOString(),
        sheetCount: 0,
        recordCount: records.length,
        counts,
        timeSummary: { start, end }
    };
}

function renderMetrics(records) {
    const total = records.length;
    const attendedCount = records.filter((record) => toCellText(record?.detail?.["\u7b7e\u5230"]) === "\u5df2\u7b7e\u5230").length;
    const absentCount = total - attendedCount;
    const attendanceRate = total > 0 ? Math.round((attendedCount / total) * 100) : 0;
    const offlineCount = records.filter((record) => record.category === "\u7ebf\u4e0b\u8bfe\u5802").length;
    const selfCount = records.filter((record) => record.category === "\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b").length;
    const businessCount = records.filter((record) => record.category === "\u5546\u5b66\u9662").length;
    const absentRecords = records.filter((record) => toCellText(record?.detail?.["\u7b7e\u5230"]) !== "\u5df2\u7b7e\u5230");
    const absentDepartmentCounts = new Map();

    absentRecords.forEach((record) => {
        const departmentName = toCellText(record?.detail?.["\u8425\u4e1a\u90e8"]);
        if (!departmentName) {
            return;
        }
        absentDepartmentCounts.set(departmentName, (absentDepartmentCounts.get(departmentName) || 0) + 1);
    });

    const [topAbsentDepartment = "", topAbsentCount = 0] = Array.from(absentDepartmentCounts.entries()).sort((left, right) => {
        if (right[1] !== left[1]) {
            return right[1] - left[1];
        }
        return left[0].localeCompare(right[0], "zh-CN");
    })[0] || [];
    const topAbsents = absentRecords.slice(0, 10);
    const targetAttendanceCount = Math.ceil(total * 0.95);
    const attendanceGap = targetAttendanceCount - attendedCount;

    if (ui.absentList) {
        if (!topAbsents.length) {
            ui.absentList.innerHTML = '<li class="text-xs text-slate-400">\u6682\u65e0\u672a\u7b7e\u5230\u6570\u636e</li>';
        } else {
            ui.absentList.innerHTML = topAbsents.map((record) => {
                const empName = toCellText(record?.detail?.["\u5458\u5de5"]) || "\u672a\u77e5";
                const empId = toCellText(record?.detail?.["\u5de5\u53f7"]) || "-";
                return `<li class="flex items-center justify-between text-sm"><span class="font-medium text-slate-700">${escapeHtml(empName)}</span><span class="text-xs text-slate-500">${escapeHtml(empId)}</span></li>`;
            }).join("");
        }
    }

    ui.metricTotal.textContent = formatCount(total);
    ui.metricTotalNote.textContent = total ? `共 ${formatCount(total)} 条排课记录` : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricTotalBadge.textContent = total ? `已汇总 ${formatCount(state.uploadedFiles.length)} 份文件` : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricTotalInsight.textContent = total
        ? `\u7ebf\u4e0b ${formatCount(offlineCount)} \u4eba \u00b7 \u81ea\u4e3b ${formatCount(selfCount)} \u4eba \u00b7 \u5546\u5b66\u9662 ${formatCount(businessCount)} \u4eba`
        : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";

    ui.metricAttended.textContent = formatCount(attendedCount);
    ui.metricAttendedNote.textContent = total ? "\u5df2\u786e\u8ba4\u5230\u573a" : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricAttendedBadge.textContent = total ? `\u5df2\u7b7e\u5230 ${formatCount(attendedCount)} \u4eba` : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricAttendedInsight.textContent = total === 0
        ? "\u7b49\u5f85\u5bfc\u5165\u6570\u636e"
        : (attendanceRate >= 90
            ? "\u7b7e\u5230\u7387\u4f18\u79c0\uff0c\u7ee7\u7eed\u4fdd\u6301"
            : (attendanceRate >= 60
                ? "\u7b7e\u5230\u8fdb\u5c55\u5e73\u7a33"
                : "\u7b7e\u5230\u6bd4\u4f8b\u504f\u4f4e\uff0c\u5efa\u8bae\u50ac\u529e"));

    ui.metricAbsent.textContent = formatCount(absentCount);
    ui.metricAbsentNote.textContent = total ? "\u9700\u91cd\u70b9\u8ddf\u8fdb\u56de\u8bbf" : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricAbsentBadge.textContent = total ? `\u672a\u7b7e\u5230 ${formatCount(absentCount)} \u4eba` : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricAbsentInsight.textContent = topAbsentCount > 0
        ? `\u7f3a\u52e4\u6700\u591a: ${topAbsentDepartment} (${formatCount(topAbsentCount)} \u4eba)`
        : "\u65e0\u7f3a\u52e4\u9884\u8b66";

    ui.metricRate.textContent = `${attendanceRate}%`;
    ui.metricRateNote.textContent = total ? "\u5b9e\u9645\u5230\u573a / \u5e94\u5230\u4eba\u6570" : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricRateBadge.textContent = total ? `\u53c2\u8bad\u7387 ${attendanceRate}%` : "\u7b49\u5f85\u5bfc\u5165\u6570\u636e";
    ui.metricRateInsight.textContent = total === 0
        ? "\u7b49\u5f85\u5bfc\u5165\u6570\u636e"
        : (attendanceGap > 0
            ? `\u8ddd\u79bb 95% \u8fbe\u6807\u7ebf\u8fd8\u9700 ${formatCount(attendanceGap)} \u4eba\u7b7e\u5230`
            : "\u5df2\u8fbe\u6210 95% \u7b7e\u5230\u76ee\u6807 \ud83c\udf89");

    ui.healthAttendance.textContent = `${attendanceRate}%`;
    ui.healthAlerts.textContent = `${formatCount(absentCount)} \u6761`;
}

function renderUploadList() {
    const files = sortByDateDesc(state.uploadedFiles, "importedAt").slice(0, 4);
    ui.uploadBadge.textContent = `${formatCount(state.uploadedFiles.length)} \u4efd\u8d44\u6599`;

    if (!files.length) {
        ui.uploadList.innerHTML = '<li class="rounded-[1.5rem] bg-slate-50 px-4 py-4 text-sm text-slate-500">\u6682\u65e0\u4e0a\u4f20\u8bb0\u5f55\uff0c\u8fdb\u5165\u6559\u52a1\u540e\u53f0\u540e\u53ef\u5bfc\u5165 Excel \u57f9\u8bad\u6570\u636e\u3002</li>';
        return;
    }

    ui.uploadList.innerHTML = files.map((file) => `
        <li class="rounded-[1.5rem] bg-slate-50 px-4 py-4 text-sm text-slate-600">
            <div class="flex items-start justify-between gap-3">
                <div>
                    <p class="font-medium text-slate-700">${escapeHtml(file.fileName)}</p>
                    <p class="mt-2 text-xs text-slate-500">${formatDateTime(file.importedAt)} · ${formatCount(file.recordCount)} \u6761 · ${humanFileSize(file.fileSize)}</p>
                </div>
                <span class="rounded-full bg-white px-3 py-1 text-xs font-medium text-slate-500">${formatCompactDate(file.importedAt)}</span>
            </div>
        </li>
    `).join("");
}

function renderAdminOverview() {
    ui.adminCurrentUser.textContent = state.session.isLoggedIn ? state.session.currentAdmin : "\u672a\u767b\u5f55";
    ui.adminOverviewSummary.textContent = `${formatCount(state.uploadedFiles.length)} \u4efd\u6587\u4ef6 / ${formatCount(state.records.length)} \u6761\u8bb0\u5f55`;
    ui.adminOverviewTime.textContent = state.uploadedFiles[0]
        ? `\u6700\u65b0\u5bfc\u5165\uff1a${formatDateTime(state.uploadedFiles[0].importedAt)}`
        : "\u5c1a\u672a\u5bfc\u5165\u57f9\u8bad\u6570\u636e";

    const latestFile = state.uploadedFiles[0];
    if (!latestFile) {
        ui.adminImportSummary.textContent = "\u6682\u65e0\u5bfc\u5165\u8bb0\u5f55";
        ui.adminImportBadge.textContent = "\u5f85\u5bfc\u5165";
        ui.adminImportDetails.innerHTML = "<p>\u5bfc\u5165\u540e\u4f1a\u5728\u8fd9\u91cc\u5c55\u793a\u6587\u4ef6\u540d\u3001\u8bb0\u5f55\u6761\u6570\u3001\u5206\u7c7b\u7ed3\u679c\u548c\u63d0\u53d6\u5230\u7684\u65f6\u95f4\u8303\u56f4\u3002</p>";
        return;
    }

    const timeSummary = latestFile.timeSummary?.start && latestFile.timeSummary?.end
        ? `${formatDateTime(latestFile.timeSummary.start)} ~ ${formatDateTime(latestFile.timeSummary.end)}`
        : "\u672a\u63d0\u53d6\u5230\u6709\u6548\u65f6\u95f4";

    ui.adminImportSummary.textContent = `${latestFile.fileName} · ${formatCount(latestFile.recordCount)} \u6761`;
    ui.adminImportBadge.textContent = "\u5df2\u5bfc\u5165";
    ui.adminImportDetails.innerHTML = [
        `<p>\u6587\u4ef6\u5927\u5c0f\uff1a${humanFileSize(latestFile.fileSize)} / \u5de5\u4f5c\u8868\uff1a${formatCount(latestFile.sheetCount || 0)}</p>`,
        `<p>\u5206\u7c7b\u7ed3\u679c\uff1a\u7ebf\u4e0b\u8bfe\u5802 ${formatCount(latestFile.counts?.["\u7ebf\u4e0b\u8bfe\u5802"] || 0)} \u6761\uff0c\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b ${formatCount(latestFile.counts?.["\u81ea\u4e3b\u57f9\u8bad\u8bfe\u7a0b"] || 0)} \u6761\uff0c\u5546\u5b66\u9662 ${formatCount(latestFile.counts?.["\u5546\u5b66\u9662"] || 0)} \u6761\uff0c\u672a\u5206\u7c7b ${formatCount(latestFile.counts?.["\u672a\u5206\u7c7b"] || 0)} \u6761</p>`,
        `<p>\u65f6\u95f4\u8303\u56f4\uff1a${timeSummary}</p>`
    ].join("");
}

function renderAdminList() {
    const isSuperAdmin = normalizeCredential(state.session.currentAdmin, true) === normalizeCredential(DEFAULT_ADMIN_USERNAME, true);
    const superAdminKey = normalizeCredential(DEFAULT_ADMIN_USERNAME, true);
    ui.adminCountLabel.textContent = `${formatCount(state.admins.length)} \u4e2a\u8d26\u53f7`;

    if (isSuperAdmin) {
        ui.adminCreateForm?.classList.remove("hidden");
        ui.adminPermissionNotice?.classList.add("hidden");
    } else {
        ui.adminCreateForm?.classList.add("hidden");
        ui.adminPermissionNotice?.classList.remove("hidden");
    }

    if (!state.admins.length) {
        ui.adminList.innerHTML = '<li class="rounded-[1.25rem] bg-slate-50 px-4 py-4 text-sm text-slate-500">\u6682\u65e0\u7ba1\u7406\u5458\u8d26\u53f7\u3002</li>';
        return;
    }

    ui.adminList.innerHTML = sortByDateDesc(state.admins, "createdAt").map((admin) => `
        <li class="rounded-[1.25rem] bg-slate-50 px-4 py-4 text-sm text-slate-600">
            <div class="flex items-center justify-between gap-3">
                <div>
                    <div class="flex flex-wrap items-center gap-2">
                        <p class="font-medium text-slate-700">${escapeHtml(admin.username)}</p>
                        ${admin.normalizedUsername === superAdminKey ? '<span class="rounded-full bg-amber-100 px-3 py-1 text-xs font-semibold text-amber-700">\u8d85\u7ea7\u7ba1\u7406\u5458</span>' : ""}
                    </div>
                    <p class="mt-1 text-xs text-slate-500">\u521b\u5efa\u4e8e ${formatDateTime(admin.createdAt)} · \u6700\u540e\u767b\u5f55: ${admin.lastLoginAt ? formatDateTime(admin.lastLoginAt) : '\u4ece\u672a\u767b\u5f55'}</p>
                </div>
                <div class="flex items-center gap-2">
                    <span class="rounded-full bg-white px-3 py-1 text-xs font-medium text-primary">\u53ef\u7528</span>
                    ${isSuperAdmin && admin.normalizedUsername !== superAdminKey ? `<button class="delete-admin-btn rounded-xl border border-rose-100 bg-rose-50 px-3 py-1 text-xs font-medium text-rose-600 transition hover:bg-rose-100" type="button" data-admin-id="${escapeHtml(admin.id)}">\u5220\u9664</button>` : ""}
                </div>
            </div>
        </li>
    `).join("");
}

function renderLogs() {
    if (!state.logs.length) {
        ui.logList.innerHTML = '<li class="rounded-[1.25rem] bg-slate-50 px-4 py-4 text-sm text-slate-500">\u6682\u65e0\u65e5\u5fd7\u8bb0\u5f55\u3002</li>';
        return;
    }

    ui.logList.innerHTML = state.logs.slice(0, 12).map((log) => `
        <li class="rounded-[1.25rem] bg-slate-50 px-4 py-4 text-sm text-slate-600">
            <div class="flex flex-col gap-2 md:flex-row md:items-start md:justify-between">
                <div>
                    <p class="font-medium text-slate-700">${escapeHtml(log.action)}</p>
                    <p class="mt-2 leading-6 text-slate-500">${escapeHtml(log.detail)}</p>
                </div>
                <div class="shrink-0 text-xs text-slate-400">${escapeHtml(log.actor || "\u7cfb\u7edf")} · ${formatDateTime(log.createdAt)}</div>
            </div>
        </li>
    `).join("");
}

function renderFileTable() {
    const files = sortByDateDesc(state.uploadedFiles, "importedAt");
    ui.fileManagementCount.textContent = `${formatCount(files.length)} \u4e2a\u6587\u4ef6`;
    ui.fileTableEmpty.textContent = files.length ? "\u53ef\u5728\u8fd9\u91cc\u5220\u9664\u5df2\u5bfc\u5165\u7684\u6587\u4ef6\uff0c\u5220\u9664\u540e\u5173\u8054\u8bb0\u5f55\u4f1a\u540c\u6b65\u79fb\u9664\u3002" : "\u5f53\u524d\u8fd8\u6ca1\u6709\u53ef\u7ba1\u7406\u7684\u4e0a\u4f20\u6587\u4ef6\u3002";

    if (!files.length) {
        ui.fileTableBody.innerHTML = '<tr><td colspan="4" class="px-3 py-4 text-slate-500">\u6682\u65e0\u6587\u4ef6\u8bb0\u5f55\u3002</td></tr>';
        return;
    }

    ui.fileTableBody.innerHTML = files.map((file) => `
        <tr>
            <td class="px-3 py-4">
                <div class="min-w-[180px]">
                    <p class="font-medium text-slate-700">${escapeHtml(file.fileName)}</p>
                    <p class="mt-1 text-xs text-slate-500">${humanFileSize(file.fileSize)}</p>
                </div>
            </td>
            <td class="px-3 py-4 text-slate-600">${formatCount(file.recordCount)}</td>
            <td class="px-3 py-4 text-slate-600">${formatDateTime(file.importedAt)}</td>
            <td class="px-3 py-4">
                <button class="delete-file-btn rounded-2xl border border-rose-200 bg-rose-50 px-4 py-2 text-sm font-medium text-rose-700 transition hover:bg-rose-100" type="button" data-file-id="${escapeHtml(file.id)}">\u5220\u9664</button>
            </td>
        </tr>
    `).join("");
}

function renderTrend(records) {
    if (!ui.trendSvg) {
        return;
    }

    const container = ui.trendSvg.parentElement;
    const isMobile = window.innerWidth < 768;
    const measuredWidth = Math.round(container?.getBoundingClientRect().width || ui.trendSvg.getBoundingClientRect().width || 760);
    const svgWidth = Math.max(isMobile ? 320 : 620, measuredWidth);
    const svgHeight = isMobile ? 250 : 320;
    const padding = {
        top: 24,
        right: isMobile ? 18 : 28,
        bottom: 44,
        left: isMobile ? 40 : 54
    };
    const plotWidth = svgWidth - padding.left - padding.right;
    const plotHeight = svgHeight - padding.top - padding.bottom;

    const months = [];
    const cursor = new Date();
    cursor.setDate(1);
    cursor.setHours(0, 0, 0, 0);
    cursor.setMonth(cursor.getMonth() - 5);

    for (let index = 0; index < 6; index += 1) {
        months.push(new Date(cursor));
        cursor.setMonth(cursor.getMonth() + 1);
    }

    const actualValues = months.map((monthStart, index) => {
        const monthEnd = new Date(monthStart);
        monthEnd.setMonth(monthEnd.getMonth() + 1);
        return records.filter((record) => {
            const timestamp = getReferenceTimestamp(record);
            return timestamp >= monthStart.getTime() && timestamp < monthEnd.getTime();
        }).length;
    });

    const targetValues = actualValues.map((value, index) => Math.max(value + (index > 2 ? 2 : 1), Math.round(value * 1.15)));
    const maxValue = Math.max(6, ...actualValues, ...targetValues);
    const pointGap = months.length > 1 ? plotWidth / (months.length - 1) : plotWidth;
    const labelStep = isMobile ? (svgWidth < 390 ? 3 : 2) : 1;

    const toX = (index) => padding.left + (pointGap * index);
    const toY = (value) => padding.top + plotHeight - ((value / maxValue) * plotHeight);
    const buildLine = (values) => values.map((value, index) => `${index === 0 ? "M" : "L"}${toX(index)} ${toY(value)}`).join("");
    const actualLine = buildLine(actualValues);
    const targetLine = buildLine(targetValues);
    const areaLine = `${actualLine}L${toX(actualValues.length - 1)} ${padding.top + plotHeight}L${toX(0)} ${padding.top + plotHeight}Z`;

    const xAxisLabels = months.map((month, index) => {
        if (index % labelStep !== 0 && index !== months.length - 1) {
            return "";
        }
        return `<text x="${toX(index)}" y="${svgHeight - 12}" text-anchor="middle">${formatMonthLabel(month)}</text>`;
    }).join("");

    const yAxisLabels = Array.from({ length: 5 }, (_, index) => {
        const value = Math.round((maxValue / 4) * index);
        const y = padding.top + plotHeight - ((plotHeight / 4) * index) + 5;
        return `<text x="${padding.left - 12}" y="${y}" text-anchor="end">${value}</text>`;
    }).join("");

    const gridLines = Array.from({ length: 4 }, (_, index) => {
        const y = padding.top + ((plotHeight / 4) * index);
        return `<path d="M${padding.left} ${y}H${svgWidth - padding.right}" stroke="#e2e8f0" stroke-width="1" />`;
    }).join("");

    const points = actualValues.map((value, index) => `
        <circle cx="${toX(index)}" cy="${toY(value)}" r="${isMobile ? 4.5 : 6}" />
    `).join("");

    const peakIndex = actualValues.indexOf(Math.max(...actualValues));
    const peakValue = actualValues[peakIndex] || 0;
    const peakX = toX(peakIndex);
    const peakY = toY(peakValue) - 12;

    ui.trendSvg.setAttribute("viewBox", `0 0 ${svgWidth} ${svgHeight}`);
    ui.trendSvg.innerHTML = `
        <defs>
            <linearGradient id="areaGradient" x1="0%" y1="0%" x2="0%" y2="100%">
                <stop offset="0%" stop-color="#0f766e" stop-opacity="0.28"></stop>
                <stop offset="100%" stop-color="#0f766e" stop-opacity="0.03"></stop>
            </linearGradient>
        </defs>
        <g fill="none" stroke-linecap="round" stroke-linejoin="round">
            <path d="M${padding.left} ${padding.top + plotHeight}H${svgWidth - padding.right}" stroke="#cbd5e1" stroke-width="1.5"></path>
            ${gridLines}
            <path d="${areaLine}" fill="url(#areaGradient)" stroke="none"></path>
            <path d="${actualLine}" stroke="#0f766e" stroke-width="${isMobile ? 3 : 4}"></path>
            <path d="${targetLine}" stroke="#f59e0b" stroke-width="${isMobile ? 2.5 : 3}" stroke-dasharray="10 10"></path>
        </g>
        <g fill="#10212b" font-size="${isMobile ? 11 : 12}">
            ${xAxisLabels}
            ${yAxisLabels}
            <text x="${peakX + 10}" y="${peakY}" fill="#0f766e">${peakValue} \u4eba</text>
        </g>
        <g fill="#0f766e">${points}</g>
    `;
}

function renderPivot() {
    if (!ui.pivotXAxis || !ui.pivotYAxis || !ui.pivotChartType || !ui.pivotChartContainer || !ui.pivotTableContainer) {
        return;
    }

    const xAxisKey = ui.pivotXAxis.value || "month";
    const yAxisKey = ui.pivotYAxis.value || "none";
    const chartType = ui.pivotChartType.value || "bar";
    const pivotData = getPivotData(state.records, xAxisKey, yAxisKey);

    if (chartType === "table") {
        ui.pivotChartContainer.classList.add("hidden");
        ui.pivotTableContainer.classList.remove("hidden");

        if (pivotChartInstance) {
            pivotChartInstance.clear();
        }

        if (!pivotData.categories.length) {
            ui.pivotTableContainer.innerHTML = '<div class="flex min-h-[220px] items-center justify-center text-sm text-slate-500">\u6682\u65e0\u53ef\u900f\u89c6\u7684\u57f9\u8bad\u6570\u636e\u3002</div>';
            return;
        }

        const valueHeaders = pivotData.hasSplit ? pivotData.seriesKeys : [pivotData.metricLabel];
        const totalColumn = pivotData.hasSplit ? '<th class="px-3 py-3 text-right font-medium">\u5408\u8ba1</th>' : "";
        const headerCells = valueHeaders
            .map((name) => `<th class="px-3 py-3 text-right font-medium">${escapeHtml(name)}</th>`)
            .join("");
        const bodyRows = pivotData.rows.map((row) => {
            const valueCells = valueHeaders
                .map((name) => `<td class="px-3 py-3 text-right text-slate-600">${formatCount(row.values[name] || row.total || 0)}</td>`)
                .join("");
            const totalCell = pivotData.hasSplit
                ? `<td class="px-3 py-3 text-right font-medium text-ink">${formatCount(row.total)}</td>`
                : "";

            return `
                <tr class="border-t border-slate-100">
                    <td class="px-3 py-3 font-medium text-ink">${escapeHtml(row.category)}</td>
                    ${valueCells}
                    ${totalCell}
                </tr>
            `;
        }).join("");
        const footerValues = valueHeaders.map((name) => {
            const totalValue = pivotData.hasSplit ? pivotData.columnTotals[name] : pivotData.grandTotal;
            return `<td class="px-3 py-3 text-right font-medium text-ink">${formatCount(totalValue)}</td>`;
        }).join("");
        const footerTotal = pivotData.hasSplit
            ? `<td class="px-3 py-3 text-right font-semibold text-ink">${formatCount(pivotData.grandTotal)}</td>`
            : "";

        ui.pivotTableContainer.innerHTML = `
            <table class="min-w-full divide-y divide-slate-200 text-sm">
                <thead>
                    <tr class="text-left text-slate-500">
                        <th class="px-3 py-3 font-medium">${escapeHtml(pivotData.xAxisLabel)}</th>
                        ${headerCells}
                        ${totalColumn}
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                    ${bodyRows}
                </tbody>
                <tfoot>
                    <tr class="border-t border-slate-200 bg-slate-50">
                        <td class="px-3 py-3 font-semibold text-ink">\u5408\u8ba1</td>
                        ${footerValues}
                        ${footerTotal}
                    </tr>
                </tfoot>
            </table>
        `;
        return;
    }

    ui.pivotTableContainer.classList.add("hidden");
    ui.pivotChartContainer.classList.remove("hidden");

    if (!window.echarts) {
        ui.pivotChartContainer.innerHTML = '<div class="flex h-full items-center justify-center text-sm text-slate-500">\u56fe\u8868\u5e93\u672a\u52a0\u8f7d\uff0c\u65e0\u6cd5\u6e32\u67d3\u81ea\u52a9\u900f\u89c6\u56fe\u3002</div>';
        return;
    }

    if (!pivotChartInstance) {
        pivotChartInstance = window.echarts.init(ui.pivotChartContainer);
    }

    const baseOption = {
        animationDuration: 320,
        color: ["#0f766e", "#14b8a6", "#f59e0b", "#0284c7", "#ef4444", "#7c3aed", "#84cc16"],
        textStyle: {
            fontFamily: '"Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif'
        }
    };

    let option = null;

    if (!pivotData.categories.length) {
        option = {
            ...baseOption,
            tooltip: { show: false },
            xAxis: { show: false },
            yAxis: { show: false },
            series: [],
