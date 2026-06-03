/**
 * opsExcelParser.ts - DEBUG VERSION to see what's happening
 */

import * as XLSX from "xlsx";
import type { Section, FollowUpQuestion } from "../types/forms";

function colIndex(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}

function getCellValue(ws: XLSX.WorkSheet, col: string, row: number): any {
  const cellAddr = col + row;
  const cell = ws[cellAddr];
  return cell ? cell.v : undefined;
}

function str(val: any): string {
  if (val === undefined || val === null) return "";
  return String(val).trim();
}

export function isOPSFormat(workbook: XLSX.WorkBook): boolean {
  const sheetNames = workbook.SheetNames;
  const hasFormatSheet = sheetNames.some((name) => name.toLowerCase() === "format");
  return hasFormatSheet;
}

export function parseOPSExcel(buffer: ArrayBuffer): any {
  const workbook = XLSX.read(buffer, { type: "array" });
  const formatSheetName = workbook.SheetNames.find((n) => n.toLowerCase() === "format") || workbook.SheetNames[0];
  const ws = workbook.Sheets[formatSheetName];

  console.log("[DEBUG] Sheet name:", formatSheetName);
  console.log("[DEBUG] Available sheets:", workbook.SheetNames);


  // Convert sheet to JSON to see what's in it
  const sheetJson = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  console.log("[DEBUG] Sheet has", sheetJson.length, "rows");

  // Log first 30 rows to see structure
  for (let i = 0; i < Math.min(30, sheetJson.length); i++) {
    const row = sheetJson[i];
    if (row && row.length > 0) {
      const hasContent = row.some((cell: any) => cell && cell.toString().trim());
      if (hasContent) {
        console.log(`[DEBUG] Row ${i + 1}:`, row.slice(0, 20));
      }
    }
  }

  // ============================================
  // DYNAMIC QUESTION TEXTS - Read from Excel cells
  // ============================================

  // Helper to get cell value by column and row
  const getCell = (col: string, row: number): string => {
    const cell = ws[`${col}${row}`];
    return cell?.v?.toString().trim() || "";
  };

  // Helper to get merged cell value
  const getMergedCell = (range: string): string => {
    const [start, end] = range.split(":");
    const cell = ws[start];
    return cell?.v?.toString().trim() || "";
  };

  // 1. Basic Info Headers
  const basicInfoQuestions = {
    deptSection: getMergedCell("F2:I4") || "Dept. / Section",
    lineZone: getMergedCell("M2:R4") || "Line / Zone",
    model: getMergedCell("F5:I7") || "Model",
    processStation: getMergedCell("M5:R7") || "Process / Station",
  };
  console.log("[DEBUG] Basic Info Headers:", basicInfoQuestions);

  // 2. Basic Info VALUES (what user entered)
  const basicInfoValues = {
    deptSection: getCell("J", 2) || getCell("J", 3) || getCell("J", 4) || "",
    lineZone: getCell("S", 2) || getCell("S", 3) || getCell("S", 4) || "",
    model: getCell("J", 5) || getCell("J", 6) || getCell("J", 7) || "",
    processStation: getCell("S", 5) || getCell("S", 6) || getCell("S", 7) || "",
  };
  console.log("[DEBUG] Basic Info Values:", basicInfoValues);

  // 3. Document Control Headers and Values
  const docControlHeaders = {
    formatNo: getMergedCell("CE2:CM4") || "Format No.",
    controlNo: getMergedCell("CE5:CM8") || "Control No.",
  };
  console.log("[DEBUG] Doc Control Headers:", docControlHeaders);

  const docControlValues = {
    formatNo: getCell("CE", 2) || getCell("CE", 3) || getCell("CE", 4) || "",
    controlNo: getCell("CE", 5) || getCell("CE", 6) || getCell("CE", 7) || getCell("CE", 8) || "",
  };
  console.log("[DEBUG] Doc Control Values:", docControlValues);

  // 4. PROCESS STEPS - READ DIRECTLY FROM CELLS
  // Step data rows: 25-64 (Excel rows)
  // Column mappings:
  // V: SN, W-Z: Importance, AA-AJ: Activity, AK-AR: Method, AS-AX: Frequency
  // AY-BG: Standard, BH: Responsibility, BI-BJ: Equipment, BK-BO: Abnormalities
  // BP-BS: Reaction Plan, BT-BW: Part Name, BZ-CD: PPE, CE-CI: Record Doc, CJ-CM: Remarks

  // First, get the column headers from row 23-24
  const stepHeaders: Record<string, string> = {
    sn: getCell("V", 23) || "SN",
    importance: getCell("W", 23) || "Item Importance",
    activity: getCell("AA", 23) || "स्टेप (What / Activity)",
    method: getCell("AK", 23) || "Method (How)",
    frequency: getCell("AS", 23) || "Frequency / When",
    standard: getCell("AY", 23) || "Standard",
    responsibility: getCell("BH", 23) || "Responsibility",
    equipment: getCell("BI", 23) || "Equipment",
    abnormalities: getCell("BK", 23) || "Possible Abnormalities",
    reactionPlan: getCell("BP", 23) || "Reaction Plan",
    partNameQty: getCell("BT", 23) || "Part Name & QTY",
    ppe: getCell("BZ", 23) || "PPEs required",
    recordDocument: getCell("CE", 23) || "Record / Document",
    remarks: getCell("CJ", 23) || "Remarks",
  };
  console.log("[DEBUG] Step Headers from Excel:", stepHeaders);

  const processSteps: any[] = [];

  // Loop through rows 25 to 64 (Excel rows)
  for (let row = 25; row <= 64; row++) {
    // Get SN first to see if this row has data
    const sn = getCell("V", row);
    if (!sn) continue; // Skip empty rows

    console.log(`[DEBUG] Reading step at row ${row}, SN: ${sn}`);

    const stepData: any = {
      sn: parseInt(sn) || row - 24,
    };

    // Helper to get value from a range of columns
    const getRangeValue = (startCol: string, endCol: string, rowNum: number): string => {
      let value = "";
      const startCode = startCol.charCodeAt(0);
      const endCode = endCol.charCodeAt(0);
      for (let code = startCode; code <= endCode; code++) {
        const col = String.fromCharCode(code);
        const cellValue = getCell(col, rowNum);
        if (cellValue) {
          value = cellValue;
          break;
        }
        // Also check the Hindi row below (row+4)
        const hindiValue = getCell(col, rowNum + 4);
        if (hindiValue) {
          value = hindiValue;
          break;
        }
      }
      return value;
    };

    // Extract each field
    stepData.importance = getRangeValue("W", "Z", row);
    stepData.activity = getRangeValue("AA", "AJ", row);
    stepData.method = getRangeValue("AK", "AR", row);
    stepData.frequency = getRangeValue("AS", "AX", row);
    stepData.standard = getRangeValue("AY", "BG", row);
    stepData.responsibility = getCell("BH", row);
    stepData.equipment = getRangeValue("BI", "BJ", row);
    stepData.abnormalities = getRangeValue("BK", "BO", row);
    stepData.reactionPlan = getRangeValue("BP", "BS", row);
    stepData.partNameQty = getRangeValue("BT", "BW", row);
    stepData.ppe = getRangeValue("BZ", "CD", row);
    stepData.recordDocument = getRangeValue("CE", "CI", row);
    stepData.remarks = getRangeValue("CJ", "CM", row);

    // Check if any field has data (besides sn)
    const hasData = Object.values(stepData).some(v => v && v.toString().trim());

    if (hasData) {
      console.log(`[DEBUG] Step ${sn} has data:`, {
        importance: stepData.importance,
        activity: stepData.activity,
        method: stepData.method,
        standard: stepData.standard,
      });
      processSteps.push(stepData);
    }
  }

  console.log(`[DEBUG] Total process steps found: ${processSteps.length}`);

  // 5. Abnormality Section
  const abnormalityHeaders = {
    route: getMergedCell("B66:T69") || "Abnormality Handling Route",
    pastProblemDetails: getMergedCell("U66:CM69") || "Past Problem Details",
  };
  console.log("[DEBUG] Abnormality Headers:", abnormalityHeaders);

  const abnormalityValues = {
    route: getCell("B", 66) || getCell("B", 67) || getCell("B", 68) || getCell("B", 69) || "",
    pastProblemDetails: getCell("U", 66) || getCell("U", 67) || getCell("U", 68) || getCell("U", 69) || "",
  };
  console.log("[DEBUG] Abnormality Values:", abnormalityValues);

  // // 6. Illustrations - Check if there's any image
  // const illustrationsImages: string[] = [];
  // if ((ws as any)["!images"] && (ws as any)["!images"].length > 0) {
  //   console.log(`[DEBUG] Found ${(ws as any)["!images"].length} images in the sheet`);
  //   illustrationsImages.push("image_present");
  // }

  const illustrationsQuestionText = getCellValue(ws, "B", 23) || "Illustrations / Images";

  console.log("[DEBUG] Illustrations question text from B23:", illustrationsQuestionText);

  // Also get any images in this range


  return {
    basicInfo: {
      deptSection: basicInfoValues.deptSection,
      lineZone: basicInfoValues.lineZone,
      model: basicInfoValues.model,
      processStation: basicInfoValues.processStation,
    },
    basicInfoHeaders: basicInfoQuestions,
    docControl: {
      formatNo: docControlValues.formatNo,
      controlNo: docControlValues.controlNo,
    },
    docControlHeaders: docControlHeaders,
    processSteps,
    stepHeaders,
    abnormalitySection: {
      route: abnormalityValues.route,
      pastProblemDetails: abnormalityValues.pastProblemDetails,
    },
    abnormalityHeaders,
    illustrations: { questionText: illustrationsQuestionText },
  };
}

export function convertOPSToFormQuestions(parsedData: any): {
  sections: Section[];
} {
  const sections: Section[] = [];

  // ============================================
  // SECTION 1: Basic Info & Document Control
  // ============================================
  const section1Questions: FollowUpQuestion[] = [];

  // Basic Info Questions
  const basicInfoFields = [
    { key: 'deptSection', header: parsedData.basicInfoHeaders?.deptSection, value: parsedData.basicInfo.deptSection },
    { key: 'lineZone', header: parsedData.basicInfoHeaders?.lineZone, value: parsedData.basicInfo.lineZone },
    { key: 'model', header: parsedData.basicInfoHeaders?.model, value: parsedData.basicInfo.model },
    { key: 'processStation', header: parsedData.basicInfoHeaders?.processStation, value: parsedData.basicInfo.processStation },
  ];

  for (const field of basicInfoFields) {
    if (field.header) {
      section1Questions.push({
        id: `q_basic_${field.key}`,
        text: field.header.replace(/[:\s]+$/, ''),
        type: "text",
        required: field.key !== 'processStation',
        description: `Enter ${field.header.toLowerCase()}`,
        suggestion: field.value || "",
        followUpQuestions: [],
        sectionId: "sec_basic_doc_control",
      } as any);
    }
  }

  // Document Control Questions
  if (parsedData.docControlHeaders?.formatNo) {
    section1Questions.push({
      id: "q_format_no",
      text: parsedData.docControlHeaders.formatNo.replace(/[:\s]+$/, ''),
      type: "text",
      required: true,
      description: `Enter format number`,
      suggestion: parsedData.docControl?.formatNo || "",
      followUpQuestions: [],
      sectionId: "sec_basic_doc_control",
    } as any);
  }

  if (parsedData.docControlHeaders?.controlNo) {
    section1Questions.push({
      id: "q_control_no",
      text: parsedData.docControlHeaders.controlNo.replace(/[:\s]+$/, ''),
      type: "text",
      required: true,
      description: `Enter control number`,
      suggestion: parsedData.docControl?.controlNo || "",
      followUpQuestions: [],
      sectionId: "sec_basic_doc_control",
    } as any);
  }

  if (section1Questions.length > 0) {
    sections.push({
      id: "sec_basic_doc_control",
      title: "Basic Information & Document Control",
      description: "Department, line, model, station, and document control details",
      questions: section1Questions,
    });
  }

  // ============================================
  // SECTION 2: Illustrations
  // ============================================
  // SECTION 2: Illustrations - Question is TEXT, Answer is IMAGE UPLOAD
  const section2Questions: FollowUpQuestion[] = [];

  // Get the question text from Excel (plain text from cell B23)
  const illustrationsQuestionText = parsedData.illustrations?.questionText || "Illustrations / Images";

  console.log("[Convert] Illustrations question text:", illustrationsQuestionText);

  section2Questions.push({
    id: "q_illustrations_images",
    text: illustrationsQuestionText,  // ← The question text from Excel
    type: "file",                      // ← Answer type is file upload (image)
    required: false,
    description: `Upload ${illustrationsQuestionText.toLowerCase()}`,
    allowedFileTypes: ["image"],       // ← Only allow images
    followUpQuestions: [],
    sectionId: "sec_illustrations",
  } as any);

  // ✅ ADD THIS MISSING CODE - Push section 2 to sections array
  if (section2Questions.length > 0) {
    sections.push({
      id: "sec_illustrations",
      title: "Illustrations & Process Details",
      description: "Upload illustrations, diagrams, and process images",
      questions: section2Questions,
    });
  }


  // ============================================
  // SECTION 3: Process Steps (Step 1 only)
  // ============================================
  const section3Questions: FollowUpQuestion[] = [];

  const stepHeaders = parsedData.stepHeaders || {};
  const stepNum = 1; // Only Step 1

  const stepHeaderKeys = [
    { key: 'importance', label: 'ada podaa', type: 'search-select' as const, options: ["Critical", "Major", "Minor", "N/A"] },
    { key: 'activity', label: 'Poda deeiii', type: 'paragraph' as const },
    { key: 'method', label: 'deii dei', type: 'paragraph' as const },
    { key: 'frequency', label: 'Frequency / When', type: 'text' as const },
    { key: 'standard', label: 'Standard (Spec. / Judgment Criteria)', type: 'paragraph' as const },
    { key: 'responsibility', label: 'Responsibility', type: 'text' as const },
    { key: 'equipment', label: 'Equipment / Measuring Eq.', type: 'text' as const },
    { key: 'abnormalities', label: 'Possible Abnormalities', type: 'paragraph' as const },
    { key: 'reactionPlan', label: 'Reaction Plan', type: 'paragraph' as const },
    { key: 'partNameQty', label: 'Part Name & QTY', type: 'text' as const },
    { key: 'ppe', label: 'PPEs required', type: 'checkbox' as const, options: ["Helmet", "Safety Goggles", "Ear Plugs", "Mask", "Apron", "Cotton Gloves", "Rubber Gloves", "Safety Shoes"] },
    { key: 'remarks', label: 'Remarks', type: 'paragraph' as const },
  ];

  // Get the step data for Step 1
  const stepData = parsedData.processSteps?.find((s: any) => s.sn === stepNum) || {};

  for (const field of stepHeaderKeys) {
    // Use custom header if available from Excel, otherwise use the label
    const headerText = stepHeaders[field.key] || field.label;
    const value = stepData[field.key] || "";

    const question: any = {
      id: `q_step${stepNum}_${field.key}`,
      text: `Step ${stepNum} - ${headerText.replace(/[\r\n]+/g, ' ').trim()}`,
      type: field.type,
      required: field.key === 'activity' || field.key === 'standard',
      description: `Enter ${headerText.toLowerCase()} for Step ${stepNum}`,
      suggestion: value,
      followUpQuestions: [],
      sectionId: "sec_process_steps",
    };

    if (field.options) {
      question.options = field.options;
    }

    section3Questions.push(question);
    console.log(`[Convert] Created: Step ${stepNum} - ${headerText}`);
  }

  if (section3Questions.length > 0) {
    sections.push({
      id: "sec_process_steps",
      title: "Process Steps - Step 1",
      description: "Detailed description and checklist for process step 1",
      questions: section3Questions,
    });
  }

  // ============================================
  // SECTION 4: Abnormality Handling
  // ============================================
  const section4Questions: FollowUpQuestion[] = [];

  if (parsedData.abnormalityHeaders?.route) {
    section4Questions.push({
      id: "q_abnormality_route",
      text: parsedData.abnormalityHeaders.route.replace(/[\r\n]+/g, ' ').trim(),
      type: "paragraph",
      required: true,
      description: `Enter abnormality handling route`,
      suggestion: parsedData.abnormalitySection?.route || "",
      followUpQuestions: [],
      sectionId: "sec_abnormality",
    } as any);
  }

  if (parsedData.abnormalityHeaders?.pastProblemDetails) {
    section4Questions.push({
      id: "q_past_problem_details",
      text: parsedData.abnormalityHeaders.pastProblemDetails,
      type: "paragraph",
      required: false,
      description: `Enter past problem details`,
      suggestion: parsedData.abnormalitySection?.pastProblemDetails || "",
      followUpQuestions: [],
      sectionId: "sec_abnormality",
    } as any);
  }

  if (section4Questions.length > 0) {
    sections.push({
      id: "sec_abnormality",
      title: "Abnormality Handling & Troubleshooting",
      description: "Response routing for failures and history logs",
      questions: section4Questions,
    });
  }

  console.log(`[Convert] Total sections created: ${sections.length}`);
  console.log(`[Convert] Total questions: ${sections.reduce((acc, s) => acc + s.questions.length, 0)}`);


  return { sections };
}