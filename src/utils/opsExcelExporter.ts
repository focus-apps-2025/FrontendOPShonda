// opsExcelExporter.ts - MULTI-SHEET VERSION WITH CORRECT ANSWER DISPLAY

import * as XLSX from "xlsx";
import { utils, write } from "xlsx-js-style";
import { injectImagesIntoXlsx, type ImagePlacement } from "../types/imageInjector";

import JSZip from "jszip";

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function col(letter: string): number {
    return utils.decode_col(letter);
}

function encodeCell(c: number, r: number): string {
    return utils.encode_cell({ c, r });
}

function cellStyle(
    bold = false,
    fontSize = 11,
    hAlign: "left" | "center" | "right" = "left",
    vAlign: "center" | "top" | "bottom" = "center",
    wrapText = false,
    top: string | null = null,
    bottom: string | null = null,
    left: string | null = null,
    right: string | null = null,
    fillColor: string | null = null
) {
    const borderObj: any = {};
    if (top) borderObj.top = { style: top };
    if (bottom) borderObj.bottom = { style: bottom };
    if (left) borderObj.left = { style: left };
    if (right) borderObj.right = { style: right };

    return {
        font: { name: "Arial", bold, sz: fontSize },
        alignment: { horizontal: hAlign, vertical: vAlign, wrapText },
        border: borderObj,
        ...(fillColor
            ? { fill: { patternType: "solid", fgColor: { rgb: fillColor } } }
            : {}),
    };
}

function mergeAndSet(
    ws: Record<string, any>,
    range: string,
    value: any,
    style: object
) {
    if (!ws["!merges"]) ws["!merges"] = [];
    ws["!merges"].push(utils.decode_range(range));

    const [startCell, endCell] = range.split(":");
    const startColMatch = startCell.match(/[A-Z]+/);
    const startRowMatch = startCell.match(/\d+/);
    const endColMatch = endCell.match(/[A-Z]+/);
    const endRowMatch = endCell.match(/\d+/);

    if (startColMatch && startRowMatch && endColMatch && endRowMatch) {
        const startColLetter = startColMatch[0];
        const startRow = parseInt(startRowMatch[0]);
        const endColLetter = endColMatch[0];
        const endRow = parseInt(endRowMatch[0]);

        const startColIdx = col(startColLetter);
        const endColIdx = col(endColLetter);

        for (let r = startRow - 1; r <= endRow - 1; r++) {
            for (let c = startColIdx; c <= endColIdx; c++) {
                const cellAddr = encodeCell(c, r);
                if (!ws[cellAddr]) ws[cellAddr] = { v: "", t: "s" };

                if (r === startRow - 1 && c === startColIdx) {
                    ws[cellAddr].v = value;
                    ws[cellAddr].t = typeof value === "number" ? "n" : "s";
                }

                const borderObj: any = {};

                if (r === startRow - 1 && (style as any).border?.top) {
                    borderObj.top = (style as any).border.top;
                }
                if (r === endRow - 1 && (style as any).border?.bottom) {
                    borderObj.bottom = (style as any).border.bottom;
                }
                if (c === startColIdx && (style as any).border?.left) {
                    borderObj.left = (style as any).border.left;
                }
                if (c === endColIdx && (style as any).border?.right) {
                    borderObj.right = (style as any).border.right;
                }

                if (r === startRow - 1 && c === startColIdx) {
                    ws[cellAddr].s = {
                        ...style,
                        border: borderObj
                    };
                } else if (Object.keys(borderObj).length > 0) {
                    ws[cellAddr].s = { border: borderObj };
                }
            }
        }
    }
}

function fillBorderRange(
    ws: Record<string, any>,
    startCol: number,
    endCol: number,
    startRow: number,
    endRow: number,
    borderStyle: string = "thin"
) {
    for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
            const addr = encodeCell(c, r);
            if (!ws[addr]) ws[addr] = { v: "", t: "s" };

            const border: any = {};

            if (r === startRow) border.top = { style: borderStyle };
            if (r === endRow) border.bottom = { style: borderStyle };
            if (c === startCol) border.left = { style: borderStyle };
            if (c === endCol) border.right = { style: borderStyle };

            if (Object.keys(border).length > 0) {
                ws[addr].s = { border };
            }
        }
    }
}

function clearInnerBorders(
    ws: Record<string, any>,
    startCol: number,
    endCol: number,
    startRow: number,
    endRow: number
) {
    for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
            const addr = encodeCell(c, r);
            if (!ws[addr]) ws[addr] = { v: "", t: "s" };
            if (!ws[addr].s) ws[addr].s = {};
            ws[addr].s.border = {};
        }
    }

    for (let c = startCol; c <= endCol; c++) {
        const addr = encodeCell(c, startRow);
        if (!ws[addr].s) ws[addr].s = {};
        ws[addr].s.border = { top: { style: "thin" } };
    }

    for (let c = startCol; c <= endCol; c++) {
        const addr = encodeCell(c, endRow);
        if (!ws[addr].s) ws[addr].s = {};
        ws[addr].s.border = { bottom: { style: "thin" } };
    }

    for (let r = startRow; r <= endRow; r++) {
        const addr = encodeCell(startCol, r);
        if (!ws[addr].s) ws[addr].s = {};
        ws[addr].s.border = { left: { style: "thin" } };
    }

    for (let r = startRow; r <= endRow; r++) {
        const addr = encodeCell(endCol, r);
        if (!ws[addr].s) ws[addr].s = {};
        ws[addr].s.border = { right: { style: "thin" } };
    }
}

async function fetchAssetBase64(path: string): Promise<string> {
    try {
        const res = await fetch(path);
        if (!res.ok) return "";
        const blob = await res.blob();
        return await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                const result = reader.result as string;
                const raw = result.split(",")[1] ?? "";
                resolve(raw);
            };
            reader.readAsDataURL(blob);
        });
    } catch {
        return "";
    }
}

function getResponseTimestamp(response: any): string | undefined {
    return response.timestamp || response.createdAt || response.submittedAt;
}

function getAnswerString(response: any, questionId: string): string {
    const answer = response.answers?.[questionId];
    if (answer === null || answer === undefined || answer === "") return "";

    if (typeof answer === "object" && answer !== null) {
        if (answer.status) return String(answer.status);
        if (answer.chassisNumber) return String(answer.chassisNumber);
        if (answer.remark) return String(answer.remark);
        if (answer.zonesData) {
            const zones = Object.keys(answer.zonesData);
            if (zones.length) return `Zones: ${zones.join(", ")}`;
        }
        try { return JSON.stringify(answer); } catch { return String(answer); }
    }
    return String(answer);
}

// ============================================================
// Helper to get COMBINED answers from multiple responses
// ============================================================
function getCombinedAnswer(responses: any[], question: any): string {
    if (!question || !responses.length) return "";
    const uniqueValues = new Set<string>();
    responses.forEach(response => {
        const answer = getAnswerString(response, question.id);
        if (answer) uniqueValues.add(answer);
    });
    return Array.from(uniqueValues).join(", ");
}

// ============================================================
// Helper to group responses by Format No and Control No
// ============================================================
function groupResponsesByFormat(responses: any[], formatQuestion: any): Map<string, any[]> {
    const groups = new Map<string, any[]>();

    responses.forEach(response => {
        const formatAnswer = formatQuestion ? getAnswerString(response, formatQuestion.id) : "";
        const key = formatAnswer || "NO_FORMAT"; // Use empty string as key if no format

        if (!groups.has(key)) {
            groups.set(key, []);
        }
        groups.get(key)!.push(response);
    });

    return groups;
}

let globalImagePlacements: ImagePlacement[] = [];
function addImageToSheet(
    ws: Record<string, any>,
    base64: string,
    type: "jpeg" | "png" | "gif",
    colStart: number,
    rowStart: number,
    colEnd: number,
    rowEnd: number,
    sheetIndex: number
) {
    if (!ws["!images"]) ws["!images"] = [];
    ws["!images"].push({
        "!pos": { r: rowStart, c: colStart, R: rowEnd, C: colEnd },
        "!datatype": "base64",
        "!type": type,
        "!data": base64,
    });

    // Also add to global placements for final injection
    globalImagePlacements.push({
        base64: base64,
        type: type,
        anchor: "two",
        fromCol: colStart, fromColOff: colStart === 1 ? 70000 : 60000,
        fromRow: rowStart, fromRowOff: rowStart === 1 ? 30000 : 50000,
        toCol: colEnd, toColOff: 0,
        toRow: rowEnd, toRowOff: 0,
        cx: colEnd === 4 ? 2268378 : colEnd === 19 ? 3510396 : colEnd === 54 ? 16310303 : colEnd === 74 ? 8266952 : colEnd === 23 ? 2113972 : colEnd === 26 ? 2116283 : colEnd === 37 ? 4028994 : 500000,
        cy: rowEnd === 10 ? 2502652 : rowEnd === 68 ? 3528002 : rowEnd === 20 ? 3799417 : rowEnd === 20 ? 4184651 : rowEnd === 16 ? 2135908 : rowEnd === 16 ? 2147454 : rowEnd === 20 ? 3847875 : 500000,
        name: `Image_${sheetIndex}_${colStart}_${rowStart}`
    });
}

// ============================================================
// MAIN EXPORT FUNCTION
// ============================================================


async function fetchImageFromUrl(url: string): Promise<string | null> {
    try {
        // Use cors-anywhere proxy
        const proxyUrl = `https://cors-anywhere.herokuapp.com/${url}`;

        console.log("Fetching image via proxy:", proxyUrl);

        const response = await fetch(proxyUrl, {
            headers: {
                'Origin': window.location.origin,
            }
        });

        if (!response.ok) {
            console.error(`Failed to fetch image: ${response.status}`);
            return null;
        }

        const blob = await response.blob();

        // Check if it's an image
        if (!blob.type.startsWith('image/')) {
            console.error("Response is not an image:", blob.type);
            return null;
        }

        return await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                const result = reader.result as string;
                const raw = result.split(",")[1] ?? "";
                resolve(raw);
            };
            reader.onload = () => {
                const result = reader.result as string;
                const raw = result.split(",")[1] ?? "";
                resolve(raw);
            };
            reader.onerror = () => {
                console.error("FileReader error");
                resolve(null);
            };
            reader.readAsDataURL(blob);
        });
    } catch (error) {
        console.error("Failed to fetch image:", error);
        return null;
    }
}

function isImageUrl(url: string): boolean {
    const imageExtensions = /\.(jpg|jpeg|png|gif|webp|bmp)(\?.*)?$/i;
    return imageExtensions.test(url);
}

// ============================================================
// PRINT SETUP HELPER FUNCTION
// ============================================================
// ============================================================
// PRINT SETUP HELPER FUNCTION - CORRECTED VERSION
// ============================================================

function addPrintSettings(ws: Record<string, any>) {

    ws["!printArea"] = "B1:CM77";

    ws["!margins"] = {
        left: 0,
        right: 0,
        top: 0,
        bottom: 0,
        header: 0,
        footer: 0
    };

    ws["!pageSetup"] = {
        orientation: "landscape", // IMPORTANT
        paperSize: "A3",          // IMPORTANT

        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,

        horizontalCentered: true,
        verticalCentered: true,

        scale: 100
    };

    ws["!sheetViews"] = [{
        view: "pageLayout",
        showGridLines: false,
        zoomScale: 100,
        zoomScaleNormal: 100
    }];

    ws["!printOptions"] = {
        headings: false,
        gridLines: false,
        horizontalCentered: true,
        verticalCentered: true
    };

    ws["!pageMargins"] = {
        left: 0,
        right: 0,
        top: 0,
        bottom: 0,
        header: 0,
        footer: 0
    };

    ws["!autofilter"] = undefined;
}
// Helper function to find a question by its field type
function findQuestionByFieldType(questions: any[], fieldType: string): any | null {
    const fieldPatterns: Record<string, string[]> = {
        importance: ['importance', 'item importance', 'star', 'priority'],
        activity: ['step', 'what', 'activity', 'स्टेप', 'operation'],
        method: ['method', 'how', 'procedure', 'way'],
        frequency: ['frequency', 'when', 'how often', 'interval'],
        standard: ['standard', 'spec', 'judgment', 'criteria', 'specification'],
        responsibility: ['responsibility', 'responsible', 'who', 'assigned'],
        equipment: ['equipment', 'measuring', 'gauge', 'tool', 'instrument'],
        abnormalities: ['abnormalities', 'possible', 'abnormality', 'issue', 'problem'],
        reaction: ['reaction', 'plan', 'reaction plan', 'action', 'corrective'],
        part: ['part', 'qty', 'part name', 'quantity', 'component', 'material'],
        ppe: ['ppe', 'required', 'safety', 'protective', 'equipment'],
        record: ['record', 'document', 'form', 'log', 'sheet'],
        remarks: ['remark', 'remarks', 'note', 'comment', 'additional']
    };

    const patterns = fieldPatterns[fieldType] || [fieldType];

    const found = questions.find(q => {
        if (!q || !q.text) return false;
        const text = q.text.toLowerCase();
        const match = patterns.some(pattern => text.includes(pattern.toLowerCase()));
        if (match) {
            console.log(`    → Matched "${fieldType}" with question: "${q.text}"`);
        }
        return match;
    });

    return found || null;
}
async function createSingleSheet(
    groupResponses: any[],
    headerQuestions: any[],
    instructionsQuestions: any[],
    pastProblemsQuestions: any[],
    processStepsQuestions: any[],
    associateSignQuestions: any[],
    sheetIndex: number

): Promise<Record<string, any>> {


    // Fetch images for this sheet
    const [
        logoBImageBase64,
        stopCallWaitImageBase64,
        noSymbolImageBase64,
        noMobileImageBase64,
        ppeGuideImageBase64,
        ppeGlovesImageBase64,
        fiveSImageBase64,
        qrCodeImageBase64,
        shiftTimingImageBase64,
    ] = await Promise.all([
        fetchAssetBase64("/assets/Companylogo.png"),
        fetchAssetBase64("/assets/Safetyposter.png"),
        fetchAssetBase64("/assets/Dontrun.png"),
        fetchAssetBase64("/assets/dontusemobile.png"),
        fetchAssetBase64("/assets/PPEGuide.png"),
        fetchAssetBase64("/assets/PPEGUIDE2.png"),
        fetchAssetBase64("/assets/5S_Guidelines.png"),
        fetchAssetBase64("/assets/Qrcode.png"),
        fetchAssetBase64("/assets/Shift_timing.png"),
    ]);

    const allResponses = groupResponses.filter(r => getResponseTimestamp(r));
    const firstResponse = allResponses[0] || null;


    // ============================================================
    // DEBUG: Log all questions and answers being processed
    // ============================================================
    console.log(`\n========== SHEET ${sheetIndex} DEBUG ==========`);
    console.log(`Total responses in this group: ${groupResponses.length}`);

    // Debug Header Questions
    console.log("\n📋 HEADER QUESTIONS:");
    headerQuestions.forEach((q, idx) => {
        console.log(`  ${idx + 1}. ID: ${q?.id || q?._id}, Text: "${q?.text || q?.label || 'N/A'}"`);
    });

    // Debug Instructions Questions  
    console.log("\n📖 INSTRUCTIONS QUESTIONS:");
    instructionsQuestions.forEach((q, idx) => {
        console.log(`  ${idx + 1}. ID: ${q?.id || q?._id}, Text: "${q?.text || q?.label || 'N/A'}"`);
    });

    // Debug Past Problems Questions
    console.log("\n⚠️ PAST PROBLEMS QUESTIONS:");
    pastProblemsQuestions.forEach((q, idx) => {
        console.log(`  ${idx + 1}. ID: ${q?.id || q?._id}, Text: "${q?.text || q?.label || 'N/A'}"`);
    });

    // Debug Process Steps Questions
    console.log("\n🔧 PROCESS STEPS QUESTIONS:");
    processStepsQuestions.forEach((q, idx) => {
        console.log(`  ${idx + 1}. ID: ${q?.id || q?._id}, Text: "${q?.text || q?.label || 'N/A'}"`);
    });

    // Debug Associate Sign Questions
    console.log("\n✍️ ASSOCIATE SIGN QUESTIONS:");
    associateSignQuestions.forEach((q, idx) => {
        console.log(`  ${idx + 1}. ID: ${q?.id || q?._id}, Text: "${q?.text || q?.label || 'N/A'}"`);
    });

    // Debug Answers from first response

    if (firstResponse) {
        console.log("\n📝 FIRST RESPONSE ANSWERS:");
        const answerKeys = Object.keys(firstResponse.answers || {});
        answerKeys.forEach(key => {
            const answer = firstResponse.answers[key];
            const answerStr = typeof answer === 'object' ? JSON.stringify(answer).substring(0, 100) : String(answer).substring(0, 100);
            console.log(`  ${key}: ${answerStr}${answerStr.length >= 100 ? '...' : ''}`);
        });
    }

    // Debug mapping of process steps questions to expected fields
    console.log("\n🔍 PROCESS STEPS MAPPING:");
    const expectedFields = [
        'importance', 'activity', 'method', 'frequency', 'standard',
        'responsibility', 'equipment', 'abnormalities', 'reaction',
        'part', 'ppe', 'record', 'remarks'
    ];

    expectedFields.forEach(field => {
        const matched = findQuestionByFieldType(processStepsQuestions, field);
        if (matched) {
            console.log(`  ✅ ${field} -> "${matched.text}" (ID: ${matched.id || matched._id})`);
        } else {
            console.log(`  ❌ ${field} -> NO MATCH FOUND`);
        }
    });

    console.log("=============================================\n");

    // Get COMBINED answers for header fields (Section 1)
    const getCombinedHeaderAnswer = (question: any): string => {
        return getCombinedAnswer(allResponses, question);
    };

    const getInstructionsAnswer = (question: any): string => {
        if (!firstResponse || !question) return "";
        return getAnswerString(firstResponse, question.id);
    };

    const getPastProblemsAnswer = (question: any): string => {
        if (!firstResponse || !question) return "";
        return getAnswerString(firstResponse, question.id);
    };

    const getProcessStepAnswer = (response: any, question: any): string => {
        if (!response || !question) return "";
        return getAnswerString(response, question.id);
    };

    const getAssociateAnswer = (question: any): string => {
        if (!firstResponse || !question) return "";
        return getAnswerString(firstResponse, question.id);
    };

    // Create worksheet
    const ws: Record<string, any> = {};
    ws["!ref"] = "B1:CM77";
    ws["!merges"] = [];

    // Row heights
    ws["!rows"] = [];
    const rowHeights: Record<number, number> = {
        0: 27, 1: 39, 2: 39, 3: 39, 4: 39, 5: 39, 6: 39,
        7: 45, 8: 45, 9: 45, 10: 45,
        11: 33.75, 12: 37.5, 13: 37.5, 14: 37.5, 15: 47.25,
        16: 37.5, 17: 37.5, 18: 37.5, 19: 37.5, 20: 37.5,
        21: 15, 22: 54.75, 23: 54.75,
    };
    for (let i = 24; i <= 63; i++) {
        rowHeights[i] = 62.15;
    }
    rowHeights[64] = 15;
    rowHeights[65] = 25;
    rowHeights[66] = 46.5;
    rowHeights[67] = 47.25;
    rowHeights[68] = 29.25;
    rowHeights[69] = 14.25;
    rowHeights[70] = 29.25;
    rowHeights[71] = 29.25;
    rowHeights[72] = 29.25;
    rowHeights[73] = 29.25;
    rowHeights[74] = 29.25;
    rowHeights[75] = 29.25;
    rowHeights[76] = 39;

    for (const [r, h] of Object.entries(rowHeights)) {
        ws["!rows"][Number(r)] = { hpt: h };
    }

    // Column widths
    const colWidthMap: Record<string, number> = {
        A: 15.27, B: 4.54, C: 8.0, D: 5.73, E: 22.73, F: 4.54, G: 8.0,
        H: 8.0, I: 7.73, J: 8.73, K: 8.0, L: 8.0, M: 4.54, N: 8.0,
        O: 9.27, P: 7.54, Q: 6.45, R: 13.18, S: 11.82, T: 13.82, U: 10.27,
        V: 11.54, W: 12.27, X: 4.54, Y: 8.0, Z: 8.0, AA: 8.0, AB: 8.0,
        AC: 8.0, AD: 8.18, AE: 7.18, AF: 7.45, AG: 4.54, AH: 8.0, AI: 6.82,
        AJ: 9.27, AK: 14.82, AL: 8.0, AM: 8.0, AN: 8.0, AO: 8.0, AP: 8.0,
        AQ: 8.0, AR: 8.0, AS: 4.27, AT: 8.0, AU: 8.0, AV: 8.0, AW: 8.0,
        AX: 9.45, AY: 23.0, AZ: 8.0, BA: 8.0, BB: 8.0, BC: 8.0, BD: 8.0,
        BE: 8.0, BF: 8.0, BG: 8.0, BH: 30.82, BI: 8.45, BJ: 22.54, BK: 23.54,
        BL: 7.27, BM: 6.73, BN: 8.45, BO: 8.0, BP: 13.45, BQ: 11.0, BR: 15.18,
        BS: 17.0, BT: 10.82, BU: 5.18, BV: 8.0, BW: 11.82, BX: 5.18, BY: 8.0,
        BZ: 10.18, CA: 5.82, CB: 3.82, CC: 8.0, CD: 6.73, CE: 11.0, CF: 8.0,
        CG: 6.45, CH: 8.0, CI: 5.73, CJ: 9.54, CK: 9.27, CL: 12.0, CM: 5.45,
    };
    ws["!cols"] = [];
    for (let i = 0; i <= 90; i++) {
        const letter = utils.encode_col(i);
        if (colWidthMap[letter]) {
            ws["!cols"][i] = { wch: colWidthMap[letter] };
        } else {
            ws["!cols"][i] = { wch: 8.0 };
        }
    }

    // ============================================================
    // ROW 1: Retention Period
    // ============================================================
    mergeAndSet(ws, "B1:CM1",
        "Retention Period : 20 years after Model is discontinued",
        cellStyle(true, 18, "right", "center", true, "medium", null, "medium", "medium")
    );

    // ============================================================
    // ROWS 2-11: Top header block - COMBINED VALUES for Section 1
    // ============================================================

    clearInnerBorders(ws, col("B"), col("E"), 1, 10);

    // Dept / Section - COMBINED answers
    const q0 = headerQuestions[0];
    mergeAndSet(ws, "F2:I4", q0?.text || "Dept. / Section :",
        cellStyle(true, 20, "left", "center", true, "medium", "thin", null, null));
    mergeAndSet(ws, "J2:L4", getCombinedHeaderAnswer(q0),
        cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", "thin"));

    // Line / Zone - COMBINED answers
    const q1 = headerQuestions[1];
    mergeAndSet(ws, "M2:R4", q1?.text || "Line / Zone :",
        cellStyle(true, 20, "left", "center", false, "medium", "thin", "thin", "thin"));
    mergeAndSet(ws, "S2:W4", getCombinedHeaderAnswer(q1),
        cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", "medium"));

    mergeAndSet(ws, "X2:BA4", "Operation Standard",
        cellStyle(true, 72, "center", "center", false, "medium", "medium", "medium", null));

    // Model - COMBINED answers
    const q2 = headerQuestions[2];
    mergeAndSet(ws, "F5:I7", q2?.text || "Model :",
        cellStyle(true, 20, "left", "center", true, "thin", "thin", "medium", "thin"));
    mergeAndSet(ws, "J5:L7", getCombinedHeaderAnswer(q2),
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));

    // Process / Station - COMBINED answers
    const q3 = headerQuestions[3];
    mergeAndSet(ws, "M5:R7", q3?.text || "Process / Station :",
        cellStyle(true, 20, "left", "center", true, "thin", "medium", "thin", "thin"));
    mergeAndSet(ws, "S5:W7", getCombinedHeaderAnswer(q3),
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "medium"));

    // Trouble section (static)
    mergeAndSet(ws, "AH5:BA5", "Your Work When Trouble Stopped The Production Line",
        cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", null));
    mergeAndSet(ws, "AH6:AI6", "S.No.",
        cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "AJ6:AR6", "Trouble",
        cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "AS6:BA6", "Your task",
        cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", null));

    const troubles: [number, number, string, string][] = [
        [7, 1, "Equipment Trouble / Machine Break Down",
            "Stop The Line\n Inform the Zone Leader\nWrite on card if mentioned in OPS"],
        [8, 2, "A Trouble You Are Responsible For", ""],
        [9, 3, "Empty Marshal Carrier ", ""],
        [10, 4, "Stock Out / Material Shortage ", ""],
        [11, 5, "A Trouble From Different Section", ""],
    ];

    for (const [r, sno, trouble, task] of troubles) {
        mergeAndSet(ws, `AH${r}:AI${r}`, sno,
            cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", "thin"));
        mergeAndSet(ws, `AJ${r}:AR${r}`, trouble,
            cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));
        if (task) {
            mergeAndSet(ws, `AS${r}:BA${r}`, task,
                cellStyle(false, 20, "center", "center", true, "thin", r === 7 ? null : "thin", "thin", null));
        }
    }

    // Rejection Handling
    mergeAndSet(ws, "F8:J11",
        "REJECTION HANDLING :-\n\nClearly Identify Rejected / NG parts. Keep them properly with proper identification at defined Location.",
        cellStyle(false, 18, "left", "center", true, "medium", "medium", "medium", "medium"));

    // Measuring Instruments
    mergeAndSet(ws, "K8:O11", "Measuring Instruments or Gauges",
        cellStyle(true, 18, "center", "center", true, "medium", "medium", null, "thin"));

    const measuringInst = [
        "Always use Calibrated Measuring Instruments / Gauges (Ensure Calibration status before using the same).",
        "Ensure Zero setting before use.",
        "Do Not Use Unidentified Measuring Tool / Gauges.",
        "In case of any abnormality, inform Line leader and Quality Engineer to take action for suspected NG material range."
    ];

    mergeAndSet(ws, "P8:W8", measuringInst[0],
        cellStyle(false, 18, "left", "center", true, "medium", "thin", "thin", "medium"));
    mergeAndSet(ws, "P9:W9", measuringInst[1],
        cellStyle(false, 18, "left", "center", true, "thin", "thin", "thin", "medium"));
    mergeAndSet(ws, "P10:W10", measuringInst[2],
        cellStyle(false, 18, "left", "center", true, "thin", "thin", "thin", "medium"));
    mergeAndSet(ws, "P11:W11", measuringInst[3],
        cellStyle(false, 18, "left", "center", true, "thin", "medium", "thin", "medium"));

    // Prepared/Checked/Approved
    fillBorderRange(ws, col("BB"), col("BE"), 1, 7, "medium");
    fillBorderRange(ws, col("BF"), col("BI"), 1, 7, "medium");
    fillBorderRange(ws, col("BJ"), col("BL"), 1, 7, "medium");

    mergeAndSet(ws, "BB9:BE11", "Prepared",
        cellStyle(true, 28, "center", "center", false, null, "medium", "medium", "medium"));
    mergeAndSet(ws, "BF9:BI11", "Checked",
        cellStyle(true, 28, "center", "center", false, null, "medium", "medium", "medium"));
    mergeAndSet(ws, "BJ9:BL11", "Approved",
        cellStyle(true, 28, "center", "center", false, null, "medium", null, "medium"));

    // Issuance / Revision
    for (let r = 2; r <= 11; r++) {
        const topB = r === 2 ? "medium" : "thin";
        const botB = r === 11 ? null : "thin";
        mergeAndSet(ws, `BM${r}:BN${r}`, r === 11 ? "No." : "",
            cellStyle(r === 11, 20, "center", "center", false, topB, botB, "medium", "thin"));
        mergeAndSet(ws, `BO${r}:BS${r}`, r === 11 ? "DD/MM/YY" : "",
            cellStyle(r === 11, 20, "center", "center", false, topB, r === 11 ? "medium" : botB, "thin", "thin"));
        mergeAndSet(ws, `BT${r}:CD${r}`, r === 11 ? "Issuance / Revision details" : "",
            cellStyle(r === 11, 20, "center", "center", false, topB, botB, "thin", "thin"));
    }

    // Format No, Control No, QR - Make sure ranges don't overlap
    const q4 = headerQuestions[4];
    const q5 = headerQuestions[5];
    const q6 = headerQuestions[6];

    // CE2:CM4 - Format No (rows 2-4)
    mergeAndSet(ws, "CE2:CM4", (q4?.text || "Format No. : ") + getInstructionsAnswer(headerQuestions[4]),
        cellStyle(true, 20, "left", "center", false, "medium", "thin", "thin", "medium"));

    // CE5:CM8 - Control No (rows 5-8) - NOTE: starts at row 5, not overlapping with rows 2-4
    mergeAndSet(ws, "CE5:CM8", (q5?.text || "Control No. : ") + getInstructionsAnswer(headerQuestions[5]),
        cellStyle(true, 20, "left", "center", false, "thin", "thin", "thin", "medium"));

    // CE9:CM11 - QR Code (rows 9-11) - NOTE: starts at row 9, not overlapping
    mergeAndSet(ws, "CE9:CM11", ("QR Code:"),
        cellStyle(true, 20, "left", "center", false, "thin", "medium", "thin", "medium"));

    // ============================================================
    // ROW 12: Section header bar - USE NON-OVERLAPPING RANGES
    // ============================================================

    // First section: B12 to BC12 (General Instructions)
    mergeAndSet(ws, "B12:BC12", "General Instructions",
        cellStyle(true, 20, "center", "center", true, "medium", "thin", "medium", "thin"));

    // Second section: BD12 to BK12 (EMS & Safety Guidelines)
    mergeAndSet(ws, "BD12:BK12", "EMS & Safety Guidelines",
        cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", "thin"));

    // Third section: BL12 to BW12 (5S Guidelines)
    mergeAndSet(ws, "BL12:BW12", "5S Guidelines",
        cellStyle(true, 20, "center", "center", true, "medium", "thin", "thin", "thin"));

    // Fourth section: BX12 to CM12 (Process Instructions)
    mergeAndSet(ws, "BX12:CM12", "Process Instructions",
        cellStyle(true, 20, "center", "center", true, "medium", "thin", "medium", "thin"));
    // ============================================================
    // ROW 13: Sub-headers
    // ============================================================
    mergeAndSet(ws, "B13:L13", "FIFO System",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "medium", "thin"));
    mergeAndSet(ws, "M13:U13", "Non Lubrication Rule",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "V13:AD13", "Always wear PPEs / Proper uniform",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "AE13:AL13", "Wear PPEs as per your station's requirements",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "AM13:BC13", "Shift Timings",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "BD13:BG13", "Environmental Issues",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "BH13:BK13", "Safety Issues",
        cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "BX13:CM13", "1. Do Exercise at Shift Start.",
        cellStyle(false, 20, "left", "center", true, null, "thin", "medium", "thin"));

    // ============================================================
    // ROWS 14-21: General Instructions content
    // ============================================================
    const instQ4 = instructionsQuestions[4];
    const instQ5 = instructionsQuestions[5];
    const instQ6 = instructionsQuestions[6];
    const instQ7 = instructionsQuestions[7];

    mergeAndSet(ws, "B14:L21",
        getInstructionsAnswer(instQ4) ||
        "1. Bin/trolley must be changed only after complete usage of all material in it.\n" +
        "2. Empty bin/trolley should be replaced with new one\n" +
        "3. Don't top up partially filled bin\n" +
        "4. Follow FIFO on line during Process.\n" +
        "5. Do not use next bin / Trolley material until running not consumed.",
        cellStyle(false, 20, "left", "center", true, "thin", "thin", "medium", "thin"));

    mergeAndSet(ws, "M14:U15", getInstructionsAnswer(instQ5) || "Do not use any lubrication if not specified in OPS / Process Sheet.",
        cellStyle(false, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "M16:R16", getInstructionsAnswer(instQ6) || "Do not use mobile on the shopfloor",
        cellStyle(false, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "S16:U16", getInstructionsAnswer(instQ7) || "Do not run on the shopfloor",
        cellStyle(false, 20, "center", "center", true, "thin", "thin", "thin", "thin"));

    for (let r = 17; r <= 21; r++) {
        mergeAndSet(ws, `M${r}:R${r}`, "",
            cellStyle(false, 20, "left", "center", false, "thin", r === 21 ? "medium" : "thin", "thin", "thin"));
        mergeAndSet(ws, `S${r}:U${r}`, "",
            cellStyle(false, 20, "left", "center", false, "thin", r === 21 ? "medium" : "thin", "thin", "thin"));
    }

    // PPE area borders
    clearInnerBorders(ws, col("AM"), col("BC"), 13, 20);
    clearInnerBorders(ws, col("V"), col("AD"), 13, 20);
    clearInnerBorders(ws, col("AE"), col("AL"), 13, 20);

    // EMS & Safety
    const instQ8 = instructionsQuestions[8];
    const instQ9 = instructionsQuestions[9];

    mergeAndSet(ws, "BD14:BG21",
        getInstructionsAnswer(instQ8) ||
        "1. Do waste segregation.\n2. Switch off idle lights & machines\n" +
        "3. Ensure 3R Principal in daily activities\n4. If there was any leakage, communicate to Sub Leader",
        cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, "BH14:BK21",
        getInstructionsAnswer(instQ9) ||
        "1. Follow POS sheet in case of any Chemical\n2. Follow MSDS/SDS in case of any emergency regarding chemical\n3. Follow your PPE's",
        cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));

    // 5S box
    fillBorderRange(ws, col("BL"), col("BW"), 12, 20, "thin");

    // Process Instructions
    const processInstructions = [
        "1. Do Exercise at Shift Start.",
        "2. Do Not Use Fallen Electrical/Functional Parts.",
        "3. Ensure Model / Variant Change.",
        "4. Report in case of part / hardware fallen inside vehicle.",
        "5. TQ Wrench Arrow Mark should be in correct direction.",
        "6. Put Fallen Hardware in Red Bin for Zone In-Charge judgement.",
        "7. Take approval from SH / HOD before changing process sequence.",
        "8. Zone In-Charge is overall responsible to ensure work is as per OPS.",
        "9. Contaminant parts should be covered properly.",
    ];

    for (let i = 0; i < processInstructions.length; i++) {
        const r = 14 + i;  // Start from row 14
        if (r <= 21) {
            mergeAndSet(ws, `BX${r}:CM${r}`, processInstructions[i],
                cellStyle(false, 20, "left", "center", true,
                    "thin", r === 21 ? "medium" : "thin", "medium", "thin"));
        }
    }
    // ============================================================
    // ROW 22: Separator
    // ============================================================
    mergeAndSet(ws, "B22:CM22", "", cellStyle(false, 11, "left", "center", false, null, null, "medium", "medium"));

    //============================================================
    // ILLUSTRATIONS SECTION (Columns B-U, Rows 23-64)
    // ============================================================

    // First, get the illustrations question
    const illustrationsQuestion = instructionsQuestions[0];
    const illustrationsQuestionText = illustrationsQuestion?.text || "Upload Illustrations / Process Images Here";

    // Header for Illustrations (rows 23-24)
    mergeAndSet(ws, "B23:U24", illustrationsQuestionText,
        cellStyle(true, 22, "center", "middle", true, "medium", "thin", "medium", "medium"));

    // Add answer rows for illustrations - ONE MERGED CELL per response (spanning ALL 8 rows)
    const MAX_ROWS = 5;
    const displayResponses = allResponses.slice(0, MAX_ROWS);

    for (let respIdx = 0; respIdx < displayResponses.length; respIdx++) {
        const response = displayResponses[respIdx];
        const startRow = 25 + respIdx * 8;      // Starting row (25, 33, 41, 49, 57)
        const endRow = startRow + 7;             // Ending row (32, 40, 48, 56, 64) - ALL 8 rows
        const isLast = respIdx === displayResponses.length - 1;
        const bottomBorder = isLast ? "medium" : "thin";

        let imageAnswer = "";
        if (illustrationsQuestion && response) {
            imageAnswer = getAnswerString(response, illustrationsQuestion.id);
        }

        // Merge ALL 8 rows into ONE cell for this response's illustration
        mergeAndSet(ws, `B${startRow}:U${endRow}`, imageAnswer,
            cellStyle(false, 20, "center", "middle", true,
                respIdx === 0 ? "thin" : "thin",
                bottomBorder,
                "medium",
                "medium"));
    }

    // Fill remaining empty rows (if less than 5 responses)
    for (let respIdx = displayResponses.length; respIdx < MAX_ROWS; respIdx++) {
        const startRow = 25 + respIdx * 8;
        const endRow = startRow + 7;
        const isLast = respIdx === MAX_ROWS - 1;
        const bottomBorder = isLast ? "medium" : "thin";

        mergeAndSet(ws, `B${startRow}:U${endRow}`, "",
            cellStyle(false, 20, "center", "middle", true,
                "thin",
                bottomBorder,
                "medium",
                "medium"));
    }

    // ============================================================
    // PROCESS STEPS TABLE (Columns V-CM, Rows 23-64)
    // ============================================================

    // Column headers for Process Steps Table (rows 23-24)
    // Define each column with its corresponding question index
    const colHeadersList: { range: string, defaultLabel: string, questionIndex: number | null }[] = [
        { range: "V23:V24", defaultLabel: "SN", questionIndex: null },  // Static - no question
        { range: "W23:Z24", defaultLabel: "Item Importance", questionIndex: 0 },  // q_step1_importance
        { range: "AA23:AJ24", defaultLabel: "स्टेप\n (What / Activity)", questionIndex: 1 },  // q_step1_activity
        { range: "AK23:AR24", defaultLabel: "Method \n(How)", questionIndex: 2 },  // q_step1_method
        { range: "AS23:AX24", defaultLabel: "Frequency / When", questionIndex: 3 },  // q_step1_frequency
        { range: "AY23:BG24", defaultLabel: "Standard \n(Spec. / Judgment Criteria)", questionIndex: 4 },  // q_step1_standard
        { range: "BH23:BH24", defaultLabel: "Responsibility", questionIndex: 5 },  // q_step1_responsibility
        { range: "BI23:BJ24", defaultLabel: "Equipment / Measuring Eq.", questionIndex: 6 },  // q_step1_equipment
        { range: "BK23:BO24", defaultLabel: "Possible Abnormalities", questionIndex: 7 },  // q_step1_abnormalities
        { range: "BP23:BS24", defaultLabel: "Reaction Plan", questionIndex: 8 },  // q_step1_reactionPlan
        { range: "BT23:BY24", defaultLabel: "Part Name & QTY", questionIndex: 9 },  // q_step1_partNameQty
        { range: "BZ23:CD24", defaultLabel: "PPEs required", questionIndex: 10 },  // q_step1_ppe
        { range: "CE23:CI24", defaultLabel: "Record / Document", questionIndex: null },  // No matching question - keep static
        { range: "CJ23:CM24", defaultLabel: "Remarks", questionIndex: 11 },  // q_step1_remarks
    ];

    // Set column headers (rows 23-24)
    for (let i = 0; i < colHeadersList.length; i++) {
        const colDef = colHeadersList[i];
        const isLast = colDef.range.startsWith("CJ");

        let label = colDef.defaultLabel;
        // If there's a question at this index, use its text
        if (colDef.questionIndex !== null && processStepsQuestions[colDef.questionIndex]) {
            label = processStepsQuestions[colDef.questionIndex].text;
        }

        mergeAndSet(ws, colDef.range, label,
            cellStyle(true, 22, "center", "center", true,
                "medium", "thin", "thin", isLast ? "medium" : "thin"));
    }

    // Fill data rows for each response
    for (let respIdx = 0; respIdx < displayResponses.length; respIdx++) {
        const response = displayResponses[respIdx];
        const er = 25 + respIdx * 8;
        const hr = er + 4;
        const isLast = respIdx === displayResponses.length - 1;
        const outerBot = isLast ? "medium" : "thin";

        // SN Column (Static - just row number)
        mergeAndSet(ws, `V${er}:V${er + 3}`, String(respIdx + 1),
            cellStyle(true, 22, "center", "center", false, "thin", "thin", "thin", "thin"));
        mergeAndSet(ws, `V${hr}:V${hr + 3}`, "",
            cellStyle(false, 22, "center", "center", false, "thin", outerBot, "thin", "thin"));

        // Process each column (skip index 0 which is SN)
        for (let colIdx = 1; colIdx < colHeadersList.length; colIdx++) {
            const colDef = colHeadersList[colIdx];

            let answer = "";

            // Get answer from the corresponding question if it exists
            if (colDef.questionIndex !== null && processStepsQuestions[colDef.questionIndex]) {
                const question = processStepsQuestions[colDef.questionIndex];
                answer = getProcessStepAnswer(response, question);
            }

            // Handle image URLs if present
            if (answer && (answer.startsWith("http://") || answer.startsWith("https://"))) {
                const isImageExt = /\.(jpg|jpeg|png|gif|webp|bmp)(\?.*)?$/i.test(answer);
                if (isImageExt) {
                    answer = `</td>前台<td><img src="${answer}" width="100" height="100" /></td></tr></tr>`;
                }
            }

            // Extract column letters from range
            const [startRange] = colDef.range.split(":");
            const startColLetter = startRange.match(/[A-Z]+/)?.[0] || "";
            const endColLetter = colDef.range.split(":")[1].match(/[A-Z]+/)?.[0] || startColLetter;
            const isLastCol = colDef.range.startsWith("CJ");

            // Merge the 4 rows for this column's answer
            mergeAndSet(ws, `${startColLetter}${er}:${endColLetter}${er + 3}`,
                answer,
                cellStyle(false, 22, "left", "center", true,
                    "thin", "thin", "thin", isLastCol ? "medium" : "thin"));
        }
    }

    // Fill empty process step rows (if less than MAX_ROWS responses)
    for (let respIdx = displayResponses.length; respIdx < MAX_ROWS; respIdx++) {
        const er = 25 + respIdx * 8;
        const hr = er + 4;
        const isLast = respIdx === MAX_ROWS - 1;
        const outerBot = isLast ? "medium" : "thin";

        // SN for empty row
        mergeAndSet(ws, `V${er}:V${er + 3}`, String(respIdx + 1),
            cellStyle(true, 22, "center", "center", false, "thin", "thin", "thin", "thin"));
        mergeAndSet(ws, `V${hr}:V${hr + 3}`, "",
            cellStyle(false, 22, "center", "center", false, "thin", outerBot, "thin", "thin"));

        // Empty cells for all other columns
        for (let colIdx = 1; colIdx < colHeadersList.length; colIdx++) {
            const colDef = colHeadersList[colIdx];
            const [startRange] = colDef.range.split(":");
            const startColLetter = startRange.match(/[A-Z]+/)?.[0] || "";
            const endColLetter = colDef.range.split(":")[1].match(/[A-Z]+/)?.[0] || startColLetter;
            const isLastCol = colDef.range.startsWith("CJ");

            mergeAndSet(ws, `${startColLetter}${er}:${endColLetter}${er + 3}`, "",
                cellStyle(false, 22, "left", "center", true,
                    "thin", "thin", "thin", isLastCol ? "medium" : "thin"));
            mergeAndSet(ws, `${startColLetter}${hr}:${endColLetter}${hr + 3}`, "",
                cellStyle(false, 22, "center", "center", true,
                    "thin", outerBot, "thin", isLastCol ? "medium" : "thin"));
        }
    }

    // ============================================================
    // ENFORCE COMPLETE BORDERS FOR PROCESS STEPS TABLE
    // ============================================================
    const startColV = col("V");
    const endColCM = col("CM");

    // Top border (row 23)
    for (let colIdx = startColV; colIdx <= endColCM; colIdx++) {
        const cellAddr = encodeCell(colIdx, 22);
        if (!ws[cellAddr]) ws[cellAddr] = { v: "", t: "s" };
        if (!ws[cellAddr].s) ws[cellAddr].s = {};
        ws[cellAddr].s.border = { ...ws[cellAddr].s.border, top: { style: "medium" } };
    }

    // Bottom border (row 64)
    for (let colIdx = startColV; colIdx <= endColCM; colIdx++) {
        const cellAddr = encodeCell(colIdx, 63);
        if (!ws[cellAddr]) ws[cellAddr] = { v: "", t: "s" };
        if (!ws[cellAddr].s) ws[cellAddr].s = {};
        ws[cellAddr].s.border = { ...ws[cellAddr].s.border, bottom: { style: "medium" } };
    }

    // Left border (column V)
    for (let row = 23; row <= 64; row++) {
        const cellAddr = encodeCell(startColV, row - 1);
        if (!ws[cellAddr]) ws[cellAddr] = { v: "", t: "s" };
        if (!ws[cellAddr].s) ws[cellAddr].s = {};
        ws[cellAddr].s.border = { ...ws[cellAddr].s.border, left: { style: "medium" } };
    }

    // Right border (column CM)
    for (let row = 23; row <= 64; row++) {
        const cellAddr = encodeCell(endColCM, row - 1);
        if (!ws[cellAddr]) ws[cellAddr] = { v: "", t: "s" };
        if (!ws[cellAddr].s) ws[cellAddr].s = {};
        ws[cellAddr].s.border = { ...ws[cellAddr].s.border, right: { style: "medium" } };
    }
    // ============================================================
    // ROW 65: Separator
    // ============================================================
    mergeAndSet(ws, "B65:CM65", "", cellStyle(false, 11, "left", "center", false, null, null, "medium", "medium"));

    // ============================================================
    // ROWS 66-69: Abnormality handling
    // ============================================================

    // Get the abnormality route question (first question in pastProblemsQuestions)
    const abnormalityRouteQuestion = pastProblemsQuestions[0];
    const abnormalityRouteAnswer = abnormalityRouteQuestion ? getPastProblemsAnswer(abnormalityRouteQuestion) : "";

    // Get the past problem details question (second question in pastProblemsQuestions)
    const pastProblemDetailsQuestion = pastProblemsQuestions[1];
    const pastProblemDetailsAnswer = pastProblemDetailsQuestion ? getPastProblemsAnswer(pastProblemDetailsQuestion) : "";

    // LEFT SIDE (B66:T69) - Show Abnormality handling route with Question + Answer
    if (abnormalityRouteQuestion) {
        // Show question text followed by answer
        const displayText = `${abnormalityRouteQuestion.text} :\n${abnormalityRouteAnswer || "Not provided"}`;
        mergeAndSet(ws, "B66:T69", displayText,
            cellStyle(true, 26, "left", "center", true, "medium", "thin", "medium", "thin"));
    } else {
        // Fallback static content
        mergeAndSet(ws, "B66:T69",
            "Abnormality handling route : \nIn case of any abnormality inform the Zone In-Charge\n" +
            "Flow of Communication :-\nOperator  ► Team Member ► Section Mgr ► As required",
            cellStyle(true, 26, "left", "center", true, "medium", "thin", "medium", "thin"));
    }

    // RIGHT SIDE (U66:CM69) - Show Past Problem Details with Question + Answer
    if (pastProblemDetailsQuestion) {
        // Show question text followed by answer
        const displayText = `${pastProblemDetailsQuestion.text} :\n${pastProblemDetailsAnswer || "No past problems reported"}`;
        mergeAndSet(ws, "U66:CM69", displayText,
            cellStyle(true, 26, "left", "top", true, "medium", "thin", "thin", "medium"));
    } else if (pastProblemsQuestions.length > 1) {
        // Fallback using second question if available
        const fallbackQuestion = pastProblemsQuestions[1];
        const fallbackAnswer = getPastProblemsAnswer(fallbackQuestion);
        const displayText = `${fallbackQuestion?.text || "Past Problem Details"} :\n${fallbackAnswer || "No past problems reported"}`;
        mergeAndSet(ws, "U66:CM69", displayText,
            cellStyle(true, 26, "left", "top", true, "medium", "thin", "thin", "medium"));
    } else {
        // Ultimate fallback
        mergeAndSet(ws, "U66:CM69", "Past Problem Details :\nNo past problems reported",
            cellStyle(true, 26, "left", "top", true, "medium", "thin", "thin", "medium"));
    }

    // Fill the separator rows (67-69) - these are just borders, keep them empty
    for (let r = 67; r <= 69; r++) {
        mergeAndSet(ws, `B${r}:T${r}`, "",
            cellStyle(false, 11, "left", "center", false, "thin", r === 69 ? "medium" : "thin", "medium", "thin"));
        mergeAndSet(ws, `U${r}:CM${r}`, "",
            cellStyle(false, 11, "left", "center", false, "thin", r === 69 ? "medium" : "thin", "thin", "medium"));
    }

    // ============================================================
    // ROW 70: Separator
    // ============================================================
    mergeAndSet(ws, "B70:CM70", "", cellStyle(false, 11, "left", "center", false, null, null, "medium", "medium"));

    // ============================================================
    // ROWS 71-76: Associate / Sign rows
    // ============================================================


    mergeAndSet(ws, "B71:G73", "Associate Name \n& Emp. Code",
        cellStyle(true, 26, "left", "center", true, "medium", "thin", "medium", "thin"));
    mergeAndSet(ws, "B74:G76", "Sign & Date",
        cellStyle(true, 26, "left", "center", true, "thin", "thin", "medium", "thin"));

    const assocMerges71 = [
        "H71:K73", "L71:Q73", "R71:U73", "V71:Z73", "AA71:AF73",
        "AG71:AK73", "AL71:AO73", "AP71:AS73", "AT71:AZ73", "BA71:BD73",
        "BE71:BE73", "BF71:BG73", "BH71:BH73", "BI71:BJ73", "BK71:BL73",
        "BM71:BQ73", "BR71:BT73", "BU71:BY73", "BZ71:CD73", "CE71:CG73",
        "CH71:CJ73", "CK71:CM73",
    ];
    const assocMerges74 = [
        "H74:K76", "L74:Q76", "R74:U76", "V74:Z76", "AA74:AF76",
        "AG74:AK76", "AL74:AO76", "AP74:AS76", "AT74:AZ76", "BA74:BD76",
        "BE74:BE76", "BF74:BG76", "BH74:BH76", "BI74:BJ76", "BK74:BL76",
        "BM74:BQ76", "BR74:BT76", "BU74:BY76", "BZ74:CD76", "CE74:CG76",
        "CH74:CJ76", "CK74:CM76",
    ];



    for (const range of assocMerges71) {
        const isLast = range.startsWith("CK");
        mergeAndSet(ws, range, "", cellStyle(false, 11, "left", "center", false,
            "medium", "thin", "thin", isLast ? "medium" : "thin"));
    }
    for (const range of assocMerges74) {
        const isLast = range.startsWith("CK");
        mergeAndSet(ws, range, "", cellStyle(false, 11, "left", "center", false,
            "thin", "thin", "thin", isLast ? "medium" : "thin"));
    }

    // ============================================================
    // ROW 77: Page Number
    // ============================================================
    mergeAndSet(ws, "B77:CD77", "",
        cellStyle(false, 11, "left", "center", false, "medium", "medium", "medium", null));
    mergeAndSet(ws, "CE77:CM77", `Page Number : 01 / ${Math.ceil(allResponses.length / 20) || 1}`,
        cellStyle(true, 26, "center", "center", false, "medium", "medium", null, "medium"));

    console.log("=== IMAGE DEBUG ===");
    console.log("logoBImageBase64:", logoBImageBase64 ? `${logoBImageBase64.length} chars` : "MISSING");
    console.log("ppeGuideImageBase64:", ppeGuideImageBase64 ? `${ppeGuideImageBase64.length} chars` : "MISSING");
    console.log("ppeGlovesImageBase64:", ppeGlovesImageBase64 ? `${ppeGlovesImageBase64.length} chars` : "MISSING");
    console.log("fiveSImageBase64:", fiveSImageBase64 ? `${fiveSImageBase64.length} chars` : "MISSING");
    // ============================================================
    // ADD IMAGES TO THIS SHEET using the global collector
    // ============================================================
    if (logoBImageBase64) {
        addImageToSheet(ws, logoBImageBase64, "png", col("B"), 1, col("E"), 10, sheetIndex);
    }
    if (stopCallWaitImageBase64) {
        // X5:AG11 → X=23, AG=32, Row5=4, Row11=10
        addImageToSheet(ws, stopCallWaitImageBase64, "png", col("X"), 4, col("AG"), 10, sheetIndex);
    }
    if (noSymbolImageBase64) {
        addImageToSheet(ws, noSymbolImageBase64, "png", col("M"), 16, col("O"), 19, sheetIndex);
    }

    // If you want it to span from S17 to U20
    if (noMobileImageBase64) {
        addImageToSheet(ws, noMobileImageBase64, "png", col("S"), 16, col("U"), 19, sheetIndex);
    }
    // PPE Guide - Span from V14 to AD21 (columns V to AD, rows 14 to 21)
    if (ppeGuideImageBase64) {
        addImageToSheet(ws, ppeGuideImageBase64, "png", col("V"), 13, col("AD"), 20, sheetIndex);
    }

    // PPE Gloves - Span from AE14 to AL21 (columns AE to AL, rows 14 to 21)
    if (ppeGlovesImageBase64) {
        addImageToSheet(ws, ppeGlovesImageBase64, "png", col("AE"), 13, col("AL"), 20, sheetIndex);
    }

    if (shiftTimingImageBase64) {
        addImageToSheet(ws, shiftTimingImageBase64, "png", col("AM"), 13, col("BC"), 20, sheetIndex);
    }
    if (fiveSImageBase64) {
        addImageToSheet(ws, fiveSImageBase64, "png", col("BL"), 12, col("BW"), 20, sheetIndex);
    }
    if (qrCodeImageBase64) {
        addImageToSheet(ws, qrCodeImageBase64, "png", col("CE"), 8, col("CM"), 9, sheetIndex);
    }

    addPrintSettings(ws);
    return ws;
}

function hasSameFormatOnly(
    responses: any[],
    formatQuestion: any
): boolean {
    if (!formatQuestion || responses.length === 0) return true;

    let firstFormatValue: string | null = null;

    for (const response of responses) {
        const formatAnswer = getAnswerString(response, formatQuestion.id);

        if (firstFormatValue === null) {
            firstFormatValue = formatAnswer;
        } else if (formatAnswer !== firstFormatValue) {
            return false;
        }
    }

    return true;
}



export async function exportResponsesToOPSExcel(
    form: any,
    responses: any[],
    sectionMapping: {
        headerSectionId: string;
        generalInstructionsSectionId: string;
        pastProblemsSectionId: string;
        processStepsSectionId: string;
        associateSignSectionId: string;
    },
    images?: OpsExcelImages,
    config?: any,
    formTitle?: string
): Promise<void> {

    // Find questions in each section
    const findSectionQuestions = (sectionId: string) => {
        const section = form.sections?.find((s: any) => s.id === sectionId);
        if (!section) return [];
        return (section.questions || []).filter(
            (q: any) => !q.parentId && !q.showWhen?.questionId
        );
    };

    const headerQuestions = findSectionQuestions(sectionMapping.headerSectionId);
    const instructionsQuestions = findSectionQuestions(sectionMapping.generalInstructionsSectionId);
    const pastProblemsQuestions = findSectionQuestions(sectionMapping.pastProblemsSectionId);
    const processStepsQuestions = findSectionQuestions(sectionMapping.processStepsSectionId);
    const associateSignQuestions = findSectionQuestions(sectionMapping.associateSignSectionId);
    const pastProblemQuestion = pastProblemsQuestions[0]; // First question in past problems section
    const pastProblemQuestionText = pastProblemQuestion?.text || "Past Problem Details";

    // Questions for grouping - FROM instructionsQuestions (Section 2)
    const formatQuestion = instructionsQuestions[0];
    const controlQuestion = instructionsQuestions[1];

    // Filter responses that have timestamps
    const validResponses = responses.filter(r => getResponseTimestamp(r));

    if (validResponses.length === 0) {
        console.warn("No valid responses to export");
        showToast?.("No valid responses to export", "error");
        return;
    }

    // Check if all responses have the same Format No and Control No
    const sameFormatAndControl = hasSameFormatOnly(validResponses, formatQuestion);

    console.log(`Same Format/Control: ${sameFormatAndControl}, Total responses: ${validResponses.length}`);

    if (sameFormatAndControl) {
        // EXPORT AS SINGLE EXCEL FILE
        console.log("Exporting as single Excel file...");

        // Create a single sheet with all responses
        const ws = await createSingleSheet(
            validResponses,
            headerQuestions,
            instructionsQuestions,
            processStepsQuestions,
            pastProblemsQuestions,
            associateSignQuestions,
            1
        );

        // Create workbook with single sheet
        const workbook = utils.book_new();
        utils.book_append_sheet(workbook, ws as any, "Format");

        // Write to buffer
        const xlsxBuffer = write(workbook, { type: "array", bookType: "xlsx" });

        // Build image placements
        const imagePlacements: ImagePlacement[] = (ws["!images"] || []).map((img: any, idx: number) => ({
            base64: img["!data"],
            type: img["!type"],
            anchor: "two",
            fromCol: img["!pos"].c, fromColOff: 60000,
            fromRow: img["!pos"].r, fromRowOff: 50000,
            toCol: img["!pos"].C, toColOff: 0,
            toRow: img["!pos"].R, toRowOff: 0,
            cx: 2000000,
            cy: 2000000,
            name: `Image_${idx}`
        })).filter(Boolean) as ImagePlacement[];

        // Inject images
        const patchedBlob = await injectImagesIntoXlsx(xlsxBuffer, imagePlacements);

        // Download single file
        const url = URL.createObjectURL(patchedBlob);
        const a = document.createElement("a");
        const fileName = formTitle || form?.title || "OPS_Export";

        // Generate filename with Format No and Control No if available
        let fileSuffix = "";
        if (validResponses.length > 0 && formatQuestion && controlQuestion) {
            const firstResponse = validResponses[0];
            const formatValue = getAnswerString(firstResponse, formatQuestion.id);
            const controlValue = getAnswerString(firstResponse, controlQuestion.id);
            if (formatValue || controlValue) {
                fileSuffix = `_${formatValue}_${controlValue}`;
            }
        }

        a.download = `${fileName}${fileSuffix}_${new Date().toLocaleDateString('en-CA')}.xlsx`;
        a.href = url;
        a.click();
        URL.revokeObjectURL(url);

        console.log("Single Excel file exported successfully");
        showToast?.("Excel file exported successfully!", "success");

    } else {
        // EXPORT AS ZIP FILE WITH MULTIPLE EXCEL FILES
        console.log("Exporting as ZIP file with multiple Excel files...");

        // Group responses by Format No and Control No
        const responseGroups = groupResponsesByFormat(validResponses, formatQuestion);

        console.log(`Found ${responseGroups.size} different Format No/Control No combinations`);

        // Create a ZIP file to hold all Excel files
        const zip = new JSZip();
        let fileCounter = 0;

        for (const [key, groupResponses] of responseGroups.entries()) {
            fileCounter++;
            const [formatValue, controlValue] = key.split("|");

            let sheetName = `${formatValue}_${controlValue}`.replace(/[^a-zA-Z0-9_-]/g, "_");
            if (!sheetName || sheetName.length === 0 || sheetName === "_") {
                sheetName = `Group_${fileCounter}`;
            }

            console.log(`Creating file ${fileCounter}: "${sheetName}.xlsx" with ${groupResponses.length} responses`);

            // Create a single sheet for this group
            const ws = await createSingleSheet(
                groupResponses,
                headerQuestions,
                instructionsQuestions,
                pastProblemsQuestions,
                processStepsQuestions,
                associateSignQuestions,
                fileCounter
            );

            // Create workbook with single sheet
            const workbook = utils.book_new();
            utils.book_append_sheet(workbook, ws as any, "Format");

            // Write to buffer
            const xlsxBuffer = write(workbook, { type: "array", bookType: "xlsx" });

            // Build image placements for this sheet
            const imagePlacements: ImagePlacement[] = (ws["!images"] || []).map((img: any, idx: number) => ({
                base64: img["!data"],
                type: img["!type"],
                anchor: "two",
                fromCol: img["!pos"].c, fromColOff: 60000,
                fromRow: img["!pos"].r, fromRowOff: 50000,
                toCol: img["!pos"].C, toColOff: 0,
                toRow: img["!pos"].R, toRowOff: 0,
                cx: 2000000,
                cy: 2000000,
                name: `Image_${idx}`
            })).filter(Boolean) as ImagePlacement[];

            // Inject images
            const patchedBlob = await injectImagesIntoXlsx(xlsxBuffer, imagePlacements);

            // Add to ZIP
            zip.file(`${sheetName}.xlsx`, patchedBlob);
        }

        // Generate and download the ZIP file
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement("a");
        const fileName = formTitle || form?.title || "OPS_Export";
        a.download = `${fileName}_${new Date().toLocaleDateString('en-CA')}.zip`;
        a.href = url;
        a.click();
        URL.revokeObjectURL(url);

        console.log(`Successfully created ZIP file with ${responseGroups.size} Excel files`);
        showToast?.(`Exported ${responseGroups.size} files as ZIP!`, "success");
    }
}

// Helper function to show toast (pass from component)
let showToast: ((message: string, type: "success" | "error" | "info") => void) | null = null;

export function setToastFunction(toastFn: typeof showToast) {
    showToast = toastFn;
}