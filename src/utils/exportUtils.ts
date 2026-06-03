import type { Question, Response } from "../types";
import type { FollowUpQuestion, Section, OpsTemplateConfig } from "../types/forms";
import * as XLSX from "xlsx";
import XLSX_STYLE from "xlsx-js-style"; // Import xlsx-style for styling
import { utils, } from "xlsx";
import { isOPSFormat, parseOPSExcel, convertOPSToFormQuestions } from "./opsExcelParser";
import { injectImagesIntoXlsx, type ImagePlacement } from "../types/imageInjector";
import { write } from "xlsx-js-style";

// Create a combined utils object
const { utils: styleUtils, writeFile } = XLSX_STYLE;
const { utils: baseUtils, read } = XLSX;

const VALID_FILE_TYPES = ["image", "pdf", "excel", "stp", "pvz", "doc", "docx", "xls", "xlsx"];

// Add this interface at the top of your file
interface FormRowData {
  [key: string]: string | undefined;
}

interface FollowUpNode {
  path: string; // "1.1.1.1.1"
  question: string;
  type: string;
  required: boolean;
  options?: string[];
  parentPath: string; // "1.1.1.1"
  triggerValue: string; // What triggers this
}
function generateId() {
  return typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : `id-${Math.random().toString(36).slice(2, 11)}`;
}

function parseBoolean(value: unknown) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return ["true", "yes", "1", "y"].includes(normalized);
  }
  return false;
}

function parseNumber(value: unknown) {
  if (value === "" || value === null || value === undefined) {
    return undefined;
  }
  if (typeof value === "number") {
    return Number.isNaN(value) ? undefined : value;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return undefined;
    }
    const normalized = trimmed.endsWith("%")
      ? trimmed.slice(0, -1).trim()
      : trimmed;
    const parsed = Number(normalized);
    return Number.isNaN(parsed) ? undefined : parsed;
  }
  const parsed = Number(value);
  return Number.isNaN(parsed) ? undefined : parsed;
}

function parseList(value: unknown) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item).trim()).filter(Boolean);
  }
  if (typeof value === "string") {
    return value
      .split("|")
      .map((item) => item.trim())
      .filter(Boolean);
  }
  return [];
}

function collectQuestions(section: Section) {
  const rows: Array<Record<string, any>> = [];
  const visit = (question: FollowUpQuestion) => {
    const followUpConfig = (question as any).followUpConfig
      ? JSON.stringify((question as any).followUpConfig)
      : "";
    rows.push({
      SectionId: section.id,
      QuestionId: question.id,
      SubParam1: question.subParam1 || "",
      SubParam2: question.subParam2 || "",
      QuestionText: question.text,
      QuestionType: question.type,
      Required: question.required,
      Options: (question.options || []).join("|"),
      Description: question.description || "",
      Suggestion: question.suggestion || "",
      ParentQuestionId: question.showWhen?.questionId || "",
      TriggerValue: question.showWhen?.value ?? "",
      GridRows: question.gridOptions?.rows?.join("|") || "",
      GridColumns: question.gridOptions?.columns?.join("|") || "",
      Min: question.min ?? "",
      Max: question.max ?? "",
      Step: question.step ?? "",
      ImageUrl: question.imageUrl || "",
      FollowUpConfig: followUpConfig,
      GoToSection: (question as any).followUpConfig
        ? Object.values((question as any).followUpConfig as Record<string, any>)
          .map((config) => config?.goToSection || "")
          .filter(Boolean)
          .join("|")
        : "",
      LinkedSectionIds: (question as any).followUpConfig
        ? Object.values((question as any).followUpConfig as Record<string, any>)
          .map((config) => config?.linkedSectionId || "")
          .filter(Boolean)
          .join("|")
        : "",
      LinkedFormIds: (question as any).followUpConfig
        ? Object.values((question as any).followUpConfig as Record<string, any>)
          .map((config) => config?.linkedFormId || "")
          .filter(Boolean)
          .join("|")
        : "",
      "Correct Answer": (question as any).correctAnswer || "",
      "Correct Answers": (question as any).correctAnswers
        ? (question as any).correctAnswers.join("|")
        : "",
      SectionBranching: (question as any).branchingRules
        ? (question as any).branchingRules
          .map((rule: any) => `${rule.optionLabel}:${rule.targetSectionId}`)
          .join("|")
        : "",
      "Ranking Logic": question.trackResponseRank ? "TRUE" : "FALSE",
      "Track Question": question.trackResponseQuestion ? "TRUE" : "FALSE",
    });
    if (question.followUpQuestions && question.followUpQuestions.length > 0) {
      question.followUpQuestions.forEach((child) => visit(child));
    }
  };
  section.questions.forEach((question) => visit(question));
  return rows;
}

export function exportResponsesToExcel(
  responses: Response[],
  question: Question
) {
  const data = responses.map((response) => {
    const rawTimestamp =
      (response as any).timestamp || (response as any).fallbackTimestamp;
    const row: Record<string, string> = {
      Timestamp: rawTimestamp ? new Date(rawTimestamp).toLocaleString() : "",
    };

    const allQuestions =
      question.sections.length > 0
        ? question.sections.flatMap((section) => section.questions)
        : question.followUpQuestions;

    allQuestions.forEach((questionItem) => {
      const answer = response.answers[questionItem.id];
      let formattedAnswer = "";

      if (Array.isArray(answer)) {
        formattedAnswer = answer.join(", ");
      } else if (typeof answer === "object" && answer !== null) {
        // Handle Product NPS Buckets (Hierarchy)
        if (answer.level1 || answer.level2 || answer.level3) {
          formattedAnswer = [
            answer.level1,
            answer.level2,
            answer.level3,
            answer.level4,
            answer.level5,
            answer.level6,
          ]
            .filter(Boolean)
            .join(" > ");
        } else {
          try {
            formattedAnswer = JSON.stringify(answer);
          } catch (e) {
            formattedAnswer = String(answer);
          }
        }
      } else {
        formattedAnswer = answer ?? "";
      }

      row[questionItem.text] = String(formattedAnswer);
    });

    return row;
  });

  const worksheet = utils.json_to_sheet(data);
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, worksheet, "Responses");
  writeFile(workbook, `responses-${question.id}.xlsx`);
}

export function exportFormStructureToExcel(form: Question) {
  const formSheet = utils.json_to_sheet([
    {
      FormId: (form as any).id || (form as any)._id || "",
      Title: form.title,
      Description: form.description,
      IsVisible:
        typeof (form as any).isVisible === "boolean"
          ? (form as any).isVisible
          : true,
      TenantId: (form as any).tenantId || "",
      LogoUrl: form.logoUrl || "",
      ImageUrl: form.imageUrl || "",
      ParentFormId: form.parentFormId || "",
      ParentFormTitle: form.parentFormTitle || "",
      LocationEnabled: (form as any).locationEnabled ? "Yes" : "No",
    },
  ]);

  const sections = form.sections || [];
  const sectionSheet = utils.json_to_sheet(
    sections.map((section, index) => ({
      SectionOrder: index + 1,
      SectionId: section.id,
      SectionTitle: section.title,
      SectionDescription: section.description || "",
      LinkedToOption: section.linkedToOption || "",
      LinkedToQuestionId: section.linkedToQuestionId || "",
    }))
  );

  const questionRows = sections.flatMap((section) => collectQuestions(section));
  const questionsSheet = utils.json_to_sheet(questionRows, {
    header: [
      "SectionId",
      "QuestionId",
      "SubParam1",
      "SubParam2",
      "QuestionText",
      "QuestionType",
      "Required",
      "Options",
      "Description",
      "Suggestion",
      "ParentQuestionId",
      "TriggerValue",
      "GridRows",
      "GridColumns",
      "Min",
      "Max",
      "Step",
      "ImageUrl",
      "FollowUpConfig",
      "GoToSection",
      "LinkedSectionIds",
      "LinkedFormIds",
      "Correct Answer",
      "Correct Answers",
      "SectionBranching",
      "Ranking Logic",
      "Track Question",
    ],
  });

  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, formSheet, "Form");
  utils.book_append_sheet(workbook, sectionSheet, "Sections");
  utils.book_append_sheet(workbook, questionsSheet, "Questions");
  writeFile(
    workbook,
    `${(form.title || "form")
      .replace(/[^a-z0-9]+/gi, "-")
      .toLowerCase()}-structure.xlsx`
  );
}

function getSectionForOption(
  optionsStr: string,
  linkedSection: string,
  optionIndex: number
): string {
  if (!optionsStr || !linkedSection) return "";

  const options = optionsStr.split(",").map((opt) => opt.trim());
  if (optionIndex >= options.length) return "";

  const optionText = options[optionIndex];

  const links = linkedSection.split(",").map((link) => link.trim());
  for (const link of links) {
    const [sectionNum, ...optionParts] = link.split(":");
    const linkedOption = optionParts.join(":").trim();
    if (linkedOption === optionText) {
      return sectionNum;
    }
  }

  return "";
}

export function createSampleFormData() {
  const sampleData = [
    {
      "Form Title": "Bike Service & Maintenance Form",
      "Form Description":
        "Comprehensive bike service assessment with nested diagnostics",
      "Section Number": "1",
      "Section Title": "Basic Bike Information",
      "Section Description": "Basic details about the bike",
      "Section Merging": "",
      Question: "What is your bike make and model?",
      "Question Description": "Manufacturer and specific model",
      "Question Type": "shortText",
      Required: "TRUE",
      Options: "",
      SubParam1: "Bike Details",
      SubParam2: "Identification",
    },
    {
      Question: "What is the bike's registration number?",
      "Question Description": "Official registration/plate number",
      "Question Type": "shortText",
      Required: "TRUE",
      Options: "",
      SubParam1: "Registration",
      SubParam2: "Legal Info",
    },
    {
      Question: "What is the current odometer reading?",
      "Question Description": "Total kilometers/miles ridden",
      "Question Type": "number",
      Required: "TRUE",
      Options: "",
      SubParam1: "Usage Data",
      SubParam2: "Mileage",
    },
    {
      "Section Number": "2",
      "Section Title": "Service Requirements Assessment",
      "Section Description": "Evaluate what service the bike needs",
      "Section Merging": "",

      // ========== MAIN QUESTION 1: ENGINE ISSUES (WITH NESTED FOLLOW-UPS) ==========
      Question: "Are you experiencing any engine issues?",
      "Question Description": "Problems related to engine performance",
      "Question Type": "multipleChoice",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Engine Health",
      SubParam2: "Performance",

      // FU1: FOR YES (WITH NESTING)
      "FU1: Option": "Yes",
      "FU1: Question Type": "dropdown",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Engine Problem Type",
      "FU1: SubParam2": "Diagnosis",
      "FU1: Question Text": "What type of engine issue are you experiencing?",
      "FU1: Options":
        "Starting Problem,Overheating,Knocking Sound,Oil Leak,Loss of Power",

      // FU1.1: Nested under "Starting Problem"
      "FU1.1: Option": "Starting Problem",
      "FU1.1: Question Type": "multipleChoice",
      "FU1.1: Required": "TRUE",
      "FU1.1: SubParam1": "Start Issue Details",
      "FU1.1: SubParam2": "Electrical",
      "FU1.1: Question Text": "What happens when you try to start?",
      "FU1.1: Options":
        "No Sound at All,Clicking Sound,Cranks But Won't Start,Starts Then Dies",

      // FU1.1.1: Nested under "Cranks But Won't Start"
      "FU1.1.1: Option": "Cranks But Won't Start",
      "FU1.1.1: Question Type": "multipleChoice",
      "FU1.1.1: Required": "TRUE",
      "FU1.1.1: SubParam1": "Fuel System",
      "FU1.1.1: SubParam2": "Ignition",
      "FU1.1.1: Question Text": "When did this problem start?",
      "FU1.1.1: Options":
        "After Fuel Fill,After Rain,Gradually Worsened,Suddenly",

      // FU1.2: Nested under "Oil Leak"
      "FU1.2: Option": "Oil Leak",
      "FU1.2: Question Type": "dropdown",
      "FU1.2: Required": "TRUE",
      "FU1.2: SubParam1": "Leak Location",
      "FU1.2: SubParam2": "Mechanical",
      "FU1.2: Question Text": "Where is the oil leaking from?",
      "FU1.2: Options": "Engine Bottom,Under Seat,Near Chain,From Filter",

      // FU2: FOR NO (SIMPLE - NO NESTING)
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Engine Status",
      "FU2: SubParam2": "Positive",
      "FU2: Question Text": "When was your last engine service?",

      // FU3: FOR N/A (SIMPLE - NO NESTING)
      "FU3: Option": "N/A",
      "FU3: Question Type": "shortText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Explanation",
      "FU3: Question Text": "Why is this not applicable?",

      // FU4: ADDITIONAL ENGINE QUESTION
      "FU4: Option": "Yes", // Same trigger as FU1
      "FU4: Question Type": "yesNoNA",
      "FU4: Required": "FALSE",
      "FU4: SubParam1": "Warning Lights",
      "FU4: SubParam2": "Dashboard",
      "FU4: Question Text": "Are any warning lights on?",

      // FU5: FINAL ENGINE QUESTION
      "FU5: Option": "Yes", // Same trigger as FU1
      "FU5: Question Type": "shortText",
      "FU5: Required": "FALSE",
      "FU5: SubParam1": "Additional Info",
      "FU5: SubParam2": "Details",
      "FU5: Question Text": "Any other engine symptoms?",
    },
    {

      // ========== MAIN QUESTION 2: BRAKE SYSTEM (WITH NESTED FOLLOW-UPS) ==========
      Question: "Are there any brake system problems?",
      "Question Description": "Issues with braking performance",
      "Question Type": "multipleChoice",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Brake Safety",
      SubParam2: "Critical",

      // FU1: FOR YES (WITH NESTING)
      "FU1: Option": "Yes",
      "FU1: Question Type": "dropdown",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Brake Problem Type",
      "FU1: SubParam2": "Safety Issue",
      "FU1: Question Text": "What brake problem are you experiencing?",
      "FU1: Options":
        "Soft Brake Lever,Grinding Noise,Brake Drag,Poor Stopping,Spongy Feel",

      // FU1.1: Nested under "Grinding Noise"
      "FU1.1: Option": "Grinding Noise",
      "FU1.1: Question Type": "dropdown",
      "FU1.1: Required": "TRUE",
      "FU1.1: SubParam1": "Noise Details",
      "FU1.1: SubParam2": "Wear Indicators",
      "FU1.1: Question Text": "When do you hear the grinding noise?",
      "FU1.1: Options":
        "Always When Braking,Only During Hard Braking,When Releasing Brakes,With Specific Speed",

      // FU1.1.1: Nested under "Always When Braking"
      "FU1.1.1: Option": "Always When Braking",
      "FU1.1.1: Question Type": "dropdown",
      "FU1.1.1: Required": "TRUE",
      "FU1.1.1: SubParam1": "Pad Condition",
      "FU1.1.1: SubParam2": "Maintenance",
      "FU1.1.1: Question Text": "When were brakes last serviced?",
      "FU1.1.1: Options":
        "Within Month,1-3 Months,3-6 Months,Over 6 Months,Never",

      // FU1.2: Nested under "Poor Stopping"
      "FU1.2: Option": "Poor Stopping",
      "FU1.2: Question Type": "dropdown",
      "FU1.2: Required": "TRUE",
      "FU1.2: SubParam1": "Stopping Distance",
      "FU1.2: SubParam2": "Performance",
      "FU1.2: Question Text": "How much has stopping distance increased?",
      "FU1.2: Options":
        "Slightly Noticeable,Significantly Increased,Very Dangerous,Unpredictable",

      // FU2: FOR NO (SIMPLE - NO NESTING)
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Brake Status",
      "FU2: SubParam2": "Good Condition",
      "FU2: Question Text": "When were brakes last checked?",

      // FU3: FOR N/A (SIMPLE - NO NESTING)
      "FU3: Option": "N/A",
      "FU3: Question Type": "shortText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Explanation",
      "FU3: Question Text": "Why are brakes not applicable?",

      // FU4: ADDITIONAL BRAKE QUESTION
      "FU4: Option": "Yes", // Same trigger as FU1
      "FU4: Question Type": "yesNoNA",
      "FU4: Required": "FALSE",
      "FU4: SubParam1": "Brake Fluid",
      "FU4: SubParam2": "Maintenance",
      "FU4: Question Text": "Has brake fluid been changed recently?",

      // FU5: FINAL BRAKE QUESTION
      "FU5: Option": "Yes", // Same trigger as FU1
      "FU5: Question Type": "shortText",
      "FU5: Required": "FALSE",
      "FU5: SubParam1": "Additional Info",
      "FU5: SubParam2": "Details",
      "FU5: Question Text": "Any vibration during braking?",
    },
    {

      // ========== MAIN QUESTION 3: TIRE CONDITION (SIMPLE FOLLOW-UPS - NO NESTING) ==========
      Question: "Are there any tire issues?",
      "Question Description": "Problems with tires and wheels",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Tire Safety",
      SubParam2: "Wheels",

      // FU1: FOR YES (SIMPLE - NO NESTING)
      "FU1: Option": "Yes",
      "FU1: Question Type": "multipleChoice",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Tire Problem Type",
      "FU1: SubParam2": "Condition",
      "FU1: Question Text": "What tire issue are you facing?",
      "FU1: Options":
        "Puncture,Wear Uneven,Wear Excessive,Bulging,Pressure Loss",

      // FU2: FOR NO (SIMPLE - NO NESTING)
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Tire Status",
      "FU2: SubParam2": "Good Condition",
      "FU2: Question Text": "When were tires last replaced?",

      // FU3: FOR N/A (SIMPLE - NO NESTING)
      "FU3: Option": "N/A",
      "FU3: Question Type": "shortText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Explanation",
      "FU3: Question Text": "Why are tires not applicable?",

      // FU4: ADDITIONAL TIRE QUESTION
      "FU4: Option": "Yes", // Same trigger as FU1
      "FU4: Question Type": "multipleChoice",
      "FU4: Required": "FALSE",
      "FU4: SubParam1": "Tire Age",
      "FU4: SubParam2": "Maintenance",
      "FU4: Question Text": "How old are your tires?",
      "FU4: Options": "Less than 1 year,1-2 years,2-3 years,Over 3 years",

      // FU5: FINAL TIRE QUESTION
      "FU5: Option": "Yes", // Same trigger as FU1
      "FU5: Question Type": "shortText",
      "FU5: Required": "FALSE",
      "FU5: SubParam1": "Additional Info",
      "FU5: SubParam2": "Details",
      "FU5: Question Text": "Any recent impacts on tires?",
    },
    {

      // ========== MAIN QUESTION 4: ELECTRICAL SYSTEM (SIMPLE FOLLOW-UPS - NO NESTING) ==========
      Question: "Are there any electrical problems?",
      "Question Description": "Issues with lights, battery, electronics",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Electrical System",
      SubParam2: "Electronics",

      // FU1: FOR YES (SIMPLE - NO NESTING)
      "FU1: Option": "Yes",
      "FU1: Question Type": "multipleChoice",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Electrical Problem",
      "FU1: SubParam2": "Diagnosis",
      "FU1: Question Text": "What electrical issue exists?",
      "FU1: Options":
        "Battery Drain,Light Failure,Indicator Problem,Horn Not Working,Display Issues",

      // FU2: FOR NO (SIMPLE - NO NESTING)
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Electrical Status",
      "FU2: SubParam2": "Good Condition",
      "FU2: Question Text": "When was battery last replaced?",

      // FU3: FOR N/A (SIMPLE - NO NESTING)
      "FU3: Option": "N/A",
      "FU3: Question Type": "shortText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Explanation",
      "FU3: Question Text": "Why is electrical system not applicable?",

      // FU4: ADDITIONAL ELECTRICAL QUESTION
      "FU4: Option": "Yes", // Same trigger as FU1
      "FU4: Question Type": "yesNoNA",
      "FU4: Required": "FALSE",
      "FU4: SubParam1": "Charging System",
      "FU4: SubParam2": "Battery",
      "FU4: Question Text": "Is the charging system working properly?",

      // FU5: FINAL ELECTRICAL QUESTION
      "FU5: Option": "Yes", // Same trigger as FU1
      "FU5: Question Type": "shortText",
      "FU5: Required": "FALSE",
      "FU5: SubParam1": "Additional Info",
      "FU5: SubParam2": "Details",
      "FU5: Question Text": "Any recent electrical modifications?",
    },
    {

      // ========== MAIN QUESTION 5: SUSPENSION & HANDLING (SIMPLE FOLLOW-UPS - NO NESTING) ==========
      Question: "Are there any suspension or handling issues?",
      "Question Description": "Problems with ride comfort and control",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Suspension",
      SubParam2: "Handling",

      // FU1: FOR YES (SIMPLE - NO NESTING)
      "FU1: Option": "Yes",
      "FU1: Question Type": "multipleChoice",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Suspension Problem",
      "FU1: SubParam2": "Ride Quality",
      "FU1: Question Text": "What handling issue do you notice?",
      "FU1: Options": "Too Soft,Too Hard,Uneven,Bottoming Out,Noise Over Bumps",

      // FU2: FOR NO (SIMPLE - NO NESTING)
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Suspension Status",
      "FU2: SubParam2": "Good Condition",
      "FU2: Question Text": "When was suspension last serviced?",

      // FU3: FOR N/A (SIMPLE - NO NESTING)
      "FU3: Option": "N/A",
      "FU3: Question Type": "shortText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Explanation",
      "FU3: Question Text": "Why is suspension not applicable?",

      // FU4: ADDITIONAL SUSPENSION QUESTION
      "FU4: Option": "Yes", // Same trigger as FU1
      "FU4: Question Type": "multipleChoice",
      "FU4: Required": "FALSE",
      "FU4: SubParam1": "Ride Quality",
      "FU4: SubParam2": "Comfort",
      "FU4: Question Text": "How would you rate ride comfort?",
      "FU4: Options":
        "Very Comfortable,Comfortable,Average,Uncomfortable,Very Uncomfortable",

      // FU5: FINAL SUSPENSION QUESTION
      "FU5: Option": "Yes", // Same trigger as FU1
      "FU5: Question Type": "shortText",
      "FU5: Required": "FALSE",
      "FU5: SubParam1": "Additional Info",
      "FU5: SubParam2": "Details",
      "FU5: Question Text": "Any handling issues during cornering?",
    },
    {
      "Section Number": "3",
      "Section Title": "Service History & Preferences",
      "Section Description": "Previous service records and preferences",
      "Section Merging": "",
      Question: "When was your last full service?",
      "Question Description": "Complete professional service",
      "Question Type": "dropdown",
      Required: "TRUE",
      Options: "Within 3 months,3-6 months,6-12 months,Over 1 year,Never",
      SubParam1: "Service History",
      SubParam2: "Maintenance",
    },
    {
      Question: "What type of service do you prefer?",
      "Question Description": "Service package preference",
      "Question Type": "multipleChoice",
      Required: "TRUE",
      Options:
        "Basic Service,Standard Service,Comprehensive Service,Premium Service,Custom",
      SubParam1: "Service Preference",
      SubParam2: "Package",
    },
    {
      Question: "Do you need a pickup/drop service?",
      "Question Description": "Transportation assistance",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Transport",
      SubParam2: "Logistics",
    },
  ];

  return sampleData;
}

export async function loadSampleFormData(): Promise<
  Partial<Question> & { sections: Section[] }
> {
  const sampleData = createSampleFormData();

  // Type assertion to FormRowData[]
  return parseNewTemplateFormat(sampleData as FormRowData[], []);
}

export function downloadNestedFormImportTemplate() {
  // Color definitions
  const COLORS = {
    MAIN: { fgColor: { rgb: "000000" } }, // Black for main headers
    FU_DARK: { fgColor: { rgb: "0000FF" } }, // Dark Blue
    FU_MEDIUM: { fgColor: { rgb: "6666FF" } }, // Medium Blue
    FU_LIGHT: { fgColor: { rgb: "CCCCFF" } }, // Light Blue
    FU_DARK_GREEN: { fgColor: { rgb: "008000" } }, // Dark Green
    FU_MEDIUM_GREEN: { fgColor: { rgb: "66B266" } }, // Medium Green
    FU_LIGHT_GREEN: { fgColor: { rgb: "CCE5CC" } }, // Light Green
  };

  // Helper function to get color for other FU groups
  function getColorForFULevel(fuNumber, level) {
    const baseColors = {
      "1": { dark: "0000FF", medium: "6666FF", light: "CCCCFF" },
      "2": { dark: "008000", medium: "66B266", light: "CCE5CC" },
      "3": { dark: "800080", medium: "B266B2", light: "E5CCE5" }, // Purple
      "4": { dark: "FF6600", medium: "FF9966", light: "FFCC99" }, // Orange
      "5": { dark: "FF0000", medium: "FF6666", light: "FFCCCC" }, // Red
    };

    const colorHex =
      baseColors[fuNumber]?.[
      level === 1 ? "dark" : level === 2 ? "medium" : "light"
      ] || "000000";

    return { fgColor: { rgb: colorHex } };
  }

  // Helper function to get contrasting text color
  function getContrastTextColor(hexColor) {
    // Remove # if present
    hexColor = hexColor.replace("#", "");

    // Convert hex to RGB
    const r = parseInt(hexColor.substr(0, 2), 16);
    const g = parseInt(hexColor.substr(2, 2), 16);
    const b = parseInt(hexColor.substr(4, 2), 16);

    // Calculate luminance
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;

    // Return black or white based on luminance
    return luminance > 0.5 ? "000000" : "FFFFFF";
  }

  const mainHeaders = [
    "Form Title",
    "Form Description",
    "Section Number",
    "Section Title",
    "Section Description",
    "Section Weightage",
    "After Section Action",
    "Subsection Of", // Replaces "Section Merging"
    "Question",
    "Question Description",
    "Question Type",
    "Required",
    "Options",
    "Option 1 Route",
    "Option 2 Route",
    "Option 3 Route",
    "Option 4 Route",
    "Option 5 Route",
    "Section Routing", // Still here for backward compatibility or bulk entry
    "Suggestion",
    "SubParam1",
    "SubParam2",
    "Allowed File Types",
    "Correct Answer",
    "Correct Answers",
    "Ranking Logic",
    "Track Question",
    "Image/File URL",
  ];

  const followUpHeaders = [];

  // Generate headers for nested follow-ups (up to 3 levels deep)
  for (let level1 = 1; level1 <= 5; level1++) {
    // Level 1 follow-ups (FU1:, FU2:, etc.)
    followUpHeaders.push(
      `FU${level1}: Option`,
      `FU${level1}: Question Type`,
      `FU${level1}: Required`,
      `FU${level1}: SubParam1`,
      `FU${level1}: SubParam2`,
      `FU${level1}: Question Text`,
      `FU${level1}: Options`,
      `FU${level1}: Correct Answer`,
      `FU${level1}: Ranking Logic`,
      `FU${level1}: Track Question`,
      `FU${level1}: Image/File URL`,
      `FU${level1}: Description`
    );

    // Level 2 follow-ups (FU1.1:, FU1.2:, etc.)
    for (let level2 = 1; level2 <= 3; level2++) {
      followUpHeaders.push(
        `FU${level1}.${level2}: Option`,
        `FU${level1}.${level2}: Question Type`,
        `FU${level1}.${level2}: Required`,
        `FU${level1}.${level2}: SubParam1`,
        `FU${level1}.${level2}: SubParam2`,
        `FU${level1}.${level2}: Question Text`,
        `FU${level1}.${level2}: Options`,
        `FU${level1}.${level2}: Correct Answer`,
        `FU${level1}.${level2}: Ranking Logic`,
        `FU${level1}.${level2}: Track Question`,
        `FU${level1}.${level2}: Image/File URL`,
        `FU${level1}.${level2}: Description`
      );

      // Level 3 follow-ups (FU1.1.1:, FU1.1.2:, etc.)
      for (let level3 = 1; level3 <= 2; level3++) {
        followUpHeaders.push(
          `FU${level1}.${level2}.${level3}: Option`,
          `FU${level1}.${level2}.${level3}: Question Type`,
          `FU${level1}.${level2}.${level3}: Required`,
          `FU${level1}.${level2}.${level3}: SubParam1`,
          `FU${level1}.${level2}.${level3}: SubParam2`,
          `FU${level1}.${level2}.${level3}: Question Text`,
          `FU${level1}.${level2}.${level3}: Options`,
          `FU${level1}.${level2}.${level3}: Correct Answer`,
          `FU${level1}.${level2}.${level3}: Ranking Logic`,
          `FU${level1}.${level2}.${level3}: Track Question`,
          `FU${level1}.${level2}.${level3}: Image/File URL`,
          `FU${level1}.${level2}.${level3}: Description`
        );
      }
    }
  }
  const headers = [...mainHeaders, ...followUpHeaders];

  const descriptions = [
    "Name of the form",
    "Overview/purpose of the form",
    "Which section (1, 2, 3...)",
    "Title of the section",
    "Description of what this section covers",
    "Percentage weight (0-100, must total 100% if used)",
    "Action after section: number (go to section #), 'end' (Submit Form), or blank (Continue to next)",
    "Subsection Of: Enter the parent section number (e.g., '1' to make this section a subsection of section 1)",
    "The question text to ask",
    "Additional details about the question",
    "Type: text, paragraph, radio, checkbox, search-select, yesNoNA, file, chassis-with-zone, chassis-without-zone, zone-in, zone-out",
    "TRUE/FALSE - is this question required?",
    "For choice questions: Option 1, Option 2, Option 3 (comma-separated)",
    "Jump to Section for Option 1: number (e.g. 2), 'end', or '0' (none)",
    "Jump to Section for Option 2: number, 'end', or '0'",
    "Jump to Section for Option 3: number, 'end', or '0'",
    "Jump to Section for Option 4: number, 'end', or '0'",
    "Jump to Section for Option 5: number, 'end', or '0'",
    "Section Routing (Advanced): comma-separated section numbers (e.g., 2,3,4 means opt1→sec2, opt2→sec3, opt3→sec4; use 0 to skip)",
    "Suggestions or recommendations for this question",
    "Additional parameter 1 for custom question configuration",
    "Additional parameter 2 for custom question configuration",
    "For file: allowed file types (image,pdf,excel) - comma-separated",
    "For quiz: correct answer value",
    "For quiz: multiple correct answers separated by |",
    "Enable ranking/tracking for this question (TRUE/FALSE)",
    "Enable response tracking for this question (TRUE/FALSE)",
    "URL for question image or file (Images, STP, or PVZ files)",
  ];

  for (let level1 = 1; level1 <= 5; level1++) {
    descriptions.push(
      `Follow-up #${level1}: Which option triggers this follow-up (must match main options)`,
      `Follow-up #${level1}: Question type`,
      `Follow-up #${level1}: Required (TRUE/FALSE)`,
      `Follow-up #${level1}: SubParam1`,
      `Follow-up #${level1}: SubParam2`,
      `Follow-up #${level1}: The follow-up question text`,
      `Follow-up #${level1}: Options (comma-separated)`,
      `Follow-up #${level1}: Correct answer (if quiz)`,
      `Follow-up #${level1}: Ranking Logic (TRUE/FALSE)`,
      `Follow-up #${level1}: Track Question (TRUE/FALSE)`,
      `Follow-up #${level1}: Image/File URL`,
      `Follow-up #${level1}: Description`
    );

    for (let level2 = 1; level2 <= 3; level2++) {
      descriptions.push(
        `Follow-up #${level1}.${level2}: Which option triggers this follow-up`,
        `Follow-up #${level1}.${level2}: Question type`,
        `Follow-up #${level1}.${level2}: Required (TRUE/FALSE)`,
        `Follow-up #${level1}.${level2}: SubParam1`,
        `Follow-up #${level1}.${level2}: SubParam2`,
        `Follow-up #${level1}.${level2}: The follow-up question text`,
        `Follow-up #${level1}.${level2}: Options`,
        `Follow-up #${level1}.${level2}: Correct answer`,
        `Follow-up #${level1}.${level2}: Ranking Logic (TRUE/FALSE)`,
        `Follow-up #${level1}.${level2}: Track Question (TRUE/FALSE)`,
        `Follow-up #${level1}.${level2}: Image/File URL`,
        `Follow-up #${level1}.${level2}: Description`
      );

      for (let level3 = 1; level3 <= 2; level3++) {
        descriptions.push(
          `Follow-up #${level1}.${level2}.${level3}: Which option triggers this follow-up`,
          `Follow-up #${level1}.${level2}.${level3}: Question type`,
          `Follow-up #${level1}.${level2}.${level3}: Required (TRUE/FALSE)`,
          `Follow-up #${level1}.${level2}.${level3}: SubParam1`,
          `Follow-up #${level1}.${level2}.${level3}: SubParam2`,
          `Follow-up #${level1}.${level2}.${level3}: The follow-up question text`,
          `Follow-up #${level1}.${level2}.${level3}: Options`,
          `Follow-up #${level1}.${level2}.${level3}: Correct answer`,
          `Follow-up #${level1}.${level2}.${level3}: Ranking Logic (TRUE/FALSE)`,
          `Follow-up #${level1}.${level2}.${level3}: Track Question (TRUE/FALSE)`,
          `Follow-up #${level1}.${level2}.${level3}: Image/File URL`,
          `Follow-up #${level1}.${level2}.${level3}: Description`
        );
      }
    }
  }

  const headerRow = headers.reduce((obj, header) => {
    obj[header] = header;
    return obj;
  }, {} as Record<string, string>);

  const descriptionRow = headers.reduce((obj, header, idx) => {
    obj[header] = descriptions[idx];
    return obj;
  }, {} as Record<string, string>);

  const separatorRow = headers.reduce((obj, header) => {
    obj[header] = "";
    return obj;
  }, {} as Record<string, string>);

  const templateData: Record<string, any>[] = [
    headerRow,
    descriptionRow,
    separatorRow,
  ];

  const formTitle = "Follow-up Testing Form - Nested Support";
  const formDesc =
    "Test form with 3 sections demonstrating basic and nested follow-ups";

  const rows = [
    {
      "Form Title": formTitle,
      "Form Description": formDesc,
      "Section Number": "1",
      "Section Title": "Section 1: Basic Screening",
      "Section Description":
        "Initial qualification questions - no follow-ups required",
      "Section Merging": "",
      Question: "Are you 18 years or older?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Age Verification",
      SubParam2: "Eligibility Check",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "Image/File URL": "",
    },
    {
      "Section Merging": "",
      Question: "Do you have valid identification documents?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Document Verification",
      SubParam2: "ID Requirements",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "Image/File URL": "",
    },
    {
      "Section Merging": "",
      Question: "Have you previously used our service before?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Service History",
      SubParam2: "Previous Experience",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "Image/File URL": "",
    },
    {
      "Section Merging": "",
      Question: "Are you available for a follow-up appointment if needed?",
      "Question Type": "yesNoNA",
      Required: "FALSE",
      Options: "Yes,No,N/A",
      SubParam1: "Availability",
      SubParam2: "Scheduling",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "Image/File URL": "",
    },
    {
      "Section Number": "2",
      "Section Title": "Section 2: Service Experience & Nested Follow-ups",
      "Section Description":
        "Questions about service experience with multi-level follow-ups",
      "Section Merging": "",
      Question: "Are you satisfied with our service quality?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Satisfaction Rating",
      SubParam2: "Quality Assessment",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "Image/File URL": "",
    },
    {
      "Section Merging": "",
      Question: "Did you complete your desired goal with our help?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Goal Achievement",
      SubParam2: "Success Metrics",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "Image/File URL": "",
    },
    {
      "Section Merging": "",
      Question: "Would you recommend us to others?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "NPS Question",
      SubParam2: "Referral Intent",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "shortText",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Positive Feedback",
      "FU1: SubParam2": "Highlight Success",
      "FU1: Question Text": "Which aspect of our service would you highlight?",
      "FU1: Options": "",
      "FU1: Correct Answer": "",
      "FU1: Ranking Logic": "FALSE",
      "FU1: Track Question": "FALSE",
      "FU1: Image/File URL": "",
      "FU2: Option": "No",
      "FU2: Question Type": "longText",
      "FU2: Required": "TRUE",
      "FU2: SubParam1": "Improvement Areas",
      "FU2: SubParam2": "Critical Feedback",
      "FU2: Question Text": "What specific improvements would you suggest?",
      "FU2: Options": "",
      "FU2: Correct Answer": "",
      "FU2: Ranking Logic": "FALSE",
      "FU2: Track Question": "FALSE",
      "FU2: Image/File URL": "",
      "FU3: Option": "N/A",
      "FU3: Question Type": "longText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Optional Explanation",
      "FU3: Question Text": "Why is this not applicable to your situation?",
      "FU3: Options": "",
      "FU3: Correct Answer": "",
      "FU3: Ranking Logic": "FALSE",
      "FU3: Track Question": "FALSE",
      "FU3: Image/File URL": "",
    },
    {
      "Section Merging": "",
      Question: "Will you use our service again in the future?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Future Intent",
      SubParam2: "Retention",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "yesNoNA",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Future Usage",
      "FU1: SubParam2": "Timeline Question",
      "FU1: Question Text": "How soon do you plan to use our service again?",
      "FU1: Options": "Yes,No,N/A",
      "FU1: Correct Answer": "",
      "FU2: Option": "No",
      "FU2: Question Type": "dropdown",
      "FU2: Required": "TRUE",
      "FU2: SubParam1": "Retention Factors",
      "FU2: SubParam2": "Improvement Areas",
      "FU2: Question Text":
        "What would change your mind about using our service?",
      "FU2: Options":
        "Better pricing,Improved features,Different support,Other",
      "FU2: Correct Answer": "",
      "FU3: Option": "N/A",
      "FU3: Question Type": "longText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Optional Context",
      "FU3: Question Text": "Please explain why this is not applicable",
      "FU3: Options": "",
      "FU3: Correct Answer": "",
    },
    {
      "Section Merging": "",
      Question: "Is your issue completely resolved?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Resolution Status",
      SubParam2: "Issue Closure",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "yesNoNA",
      "FU1: Required": "FALSE",
      "FU1: SubParam1": "Resolution Check",
      "FU1: SubParam2": "Follow-up Concerns",
      "FU1: Question Text": "Are there any remaining concerns?",
      "FU1: Options": "Yes,No,N/A",
      "FU1: Correct Answer": "",
      "FU2: Option": "No",
      "FU2: Question Type": "longText",
      "FU2: Required": "TRUE",
      "FU2: SubParam1": "Unresolved Issues",
      "FU2: SubParam2": "Detailed Feedback",
      "FU2: Question Text": "What part of your issue remains unresolved?",
      "FU2: Options": "",
      "FU2: Correct Answer": "",
      "FU3: Option": "N/A",
      "FU3: Question Type": "shortText",
      "FU3: Required": "FALSE",
      "FU3: SubParam1": "Not Applicable",
      "FU3: SubParam2": "Optional Context",
      "FU3: Question Text": "Please elaborate on why this is N/A",
      "FU3: Options": "",
      "FU3: Correct Answer": "",
    },
    {
      "Section Number": "3",
      "Section Title": "Section 3: Follow-up Support & Feedback",
      "Section Description":
        "Final section with yes/no/n/a questions and follow-ups",
      "Section Merging": "",
      Question: "Do you need additional support or resources?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Support Needs",
      SubParam2: "Resource Requirements",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "dropdown",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Support Types",
      "FU1: SubParam2": "Additional Resources",
      "FU1: Question Text": "What type of support do you need?",
      "FU1: Options": "Technical,Training,Consulting,Other",
      "FU1: Correct Answer": "",
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Satisfaction",
      "FU2: SubParam2": "Positive Feedback",
      "FU2: Question Text": "What made you feel supported?",
      "FU2: Options": "",
      "FU2: Correct Answer": "",
    },
    {
      "Section Merging": "",
      Question: "Can we contact you with service updates?",
      "Question Type": "yesNoNA",
      Required: "FALSE",
      Options: "Yes,No,N/A",
      SubParam1: "Communication Consent",
      SubParam2: "Update Preferences",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "multipleChoice",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Contact Preferences",
      "FU1: SubParam2": "Communication Method",
      "FU1: Question Text": "Preferred contact method:",
      "FU1: Options": "Email,Phone,SMS,Postal Mail",
      "FU1: Correct Answer": "",
    },
    {
      "Section Merging": "",
      Question: "Will you provide feedback on your experience?",
      "Question Type": "yesNoNA",
      Required: "FALSE",
      Options: "Yes,No,N/A",
      SubParam1: "Feedback Consent",
      SubParam2: "Review Participation",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "longText",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Detailed Feedback",
      "FU1: SubParam2": "Comprehensive Review",
      "FU1: Question Text": "Please share your detailed feedback:",
      "FU1: Options": "",
      "FU1: Correct Answer": "",
      "FU2: Option": "No",
      "FU2: Question Type": "shortText",
      "FU2: Required": "FALSE",
      "FU2: SubParam1": "Decline Reason",
      "FU2: SubParam2": "Optional Context",
      "FU2: Question Text": "Would you share why?",
      "FU2: Options": "",
      "FU2: Correct Answer": "",
    },
    {
      "Section Merging": "",
      Question: "Do you consent to data usage for service improvement?",
      "Question Type": "yesNoNA",
      Required: "TRUE",
      Options: "Yes,No,N/A",
      SubParam1: "Data Privacy",
      SubParam2: "Analytics Consent",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "shortText",
      "FU1: Required": "FALSE",
      "FU1: SubParam1": "Improvement Focus",
      "FU1: SubParam2": "Specific Areas",
      "FU1: Question Text": "What specific area should we improve?",
      "FU1: Options": "",
      "FU1: Correct Answer": "",
    },
    {
      "Section Merging": "",
      Question: "Is there anything else you'd like to share?",
      "Question Type": "yesNoNA",
      Required: "FALSE",
      Options: "Yes,No,N/A",
      SubParam1: "Eligibility",
      SubParam2: "Document Verification",
      "Allowed File Types": "",
      "Correct Answer": "",
      "Correct Answers": "",
      "FU1: Option": "Yes",
      "FU1: Question Type": "longText",
      "FU1: Required": "TRUE",
      "FU1: SubParam1": "Quality Assessment",
      "FU1: SubParam2": "Feedback Collection",
      "FU1: Question Text": "Please provide additional comments:",
      "FU1: Options": "",
      "FU1: Correct Answer": "",
    },
    {
      "Section Merging": "",
      Question: "Please upload any supporting documents (optional)",
      "Question Type": "file",
      Required: "FALSE",
      Options: "",
      SubParam1: "Service History",
      SubParam2: "Document Verification",
      "Allowed File Types": "pdf,image",
      "Correct Answer": "",
      "Correct Answers": "",
    },
  ];

  // Track section counts for merging
  const sectionCounts: Record<string, number> = {};
  const sectionFirstRow: Record<string, number> = {};

  rows.forEach((row) => {
    const sectionNum = row["Section Number"]?.toString().trim();
    if (sectionNum) {
      sectionCounts[sectionNum] = (sectionCounts[sectionNum] || 0) + 1;
      if (!sectionFirstRow[sectionNum]) {
        sectionFirstRow[sectionNum] = templateData.length + 3; // +3 for header, description, separator rows
      }
    }
  });

  rows.forEach((row: any, rowIndex) => {
    // Use `any` type
    const sectionNum = row["Section Number"]?.toString().trim();

    // Build branching string from individual columns or combined branching column
    let branchingStr = "";
    const branchingValues: (string | number)[] = [];

    // Check if we have the old format (5 separate columns)
    for (let i = 1; i <= 5; i++) {
      const branchKey = `Branching: Option ${i} Section`;
      if (row[branchKey]) {
        branchingValues.push(row[branchKey]);
      } else {
        branchingValues.push(0);
      }
    }

    // Only include if there's actual branching data
    if (branchingValues.some((v) => v !== 0 && v !== "")) {
      branchingStr = branchingValues.join(",");
    }

    const fullRow: Record<string, any> = {
      "Form Title": row["Form Title"] || "",
      "Form Description": row["Form Description"] || "",
      "Section Number": row["Section Number"] || "",
      "Section Title": row["Section Title"] || "",
      "Section Description": row["Section Description"] || "",
      "Section Weightage": row["Section Weightage"] || "",
      "After Section Action": row["After Section Action"] || "",
      "Subsection Of": row["Subsection Of"] || row["Section Merging"] || "",
      Question: row.Question || "",
      "Question Description": row["Question Description"] || "",
      "Question Type": row["Question Type"] || "",
      Required: row.Required || "FALSE",
      Options: row.Options || "",
      "Option 1 Route": branchingValues[0] !== 0 ? branchingValues[0] : "",
      "Option 2 Route": branchingValues[1] !== 0 ? branchingValues[1] : "",
      "Option 3 Route": branchingValues[2] !== 0 ? branchingValues[2] : "",
      "Option 4 Route": branchingValues[3] !== 0 ? branchingValues[3] : "",
      "Option 5 Route": branchingValues[4] !== 0 ? branchingValues[4] : "",
      "Section Routing": branchingStr,
      Suggestion: row["Suggestion"] || "",
      SubParam1: row.SubParam1 || "",
      SubParam2: row.SubParam2 || "",
      "Allowed File Types": row["Allowed File Types"] || "",
      "Correct Answer": row["Correct Answer"] || "",
      "Correct Answers": row["Correct Answers"] || "",
      "Ranking Logic": row["Ranking Logic"] || "FALSE",
      "Track Question": row["Track Question"] || "FALSE",
      "Image/File URL": row["Image/File URL"] || "",
    };

    // Add merge instructions for section columns (3-8) when same section has multiple questions
    if (sectionNum && sectionCounts[sectionNum] > 1) {
      const currentRowNum = templateData.length + 3; // +3 for header rows
      const firstRow = sectionFirstRow[sectionNum];
      // Merge columns 3-8 (C-H: Section Number, Title, Description, Weightage, Action, Subsection Of)
      fullRow["Section Merging"] = `C${firstRow}:H${firstRow + sectionCounts[sectionNum] - 1
        }`;
    }

    for (let level1 = 1; level1 <= 5; level1++) {
      const fu1Prefix = `FU${level1}`;
      fullRow[`${fu1Prefix}: Option`] = row[`${fu1Prefix}: Option`] || "";
      fullRow[`${fu1Prefix}: Question Type`] = row[`${fu1Prefix}: Question Type`] || "";
      fullRow[`${fu1Prefix}: Required`] = row[`${fu1Prefix}: Required`] || "";
      fullRow[`${fu1Prefix}: SubParam1`] = row[`${fu1Prefix}: SubParam1`] || "";
      fullRow[`${fu1Prefix}: SubParam2`] = row[`${fu1Prefix}: SubParam2`] || "";
      fullRow[`${fu1Prefix}: Question Text`] = row[`${fu1Prefix}: Question Text`] || "";
      fullRow[`${fu1Prefix}: Description`] = row[`${fu1Prefix}: Description`] || "";
      fullRow[`${fu1Prefix}: Options`] = row[`${fu1Prefix}: Options`] || "";
      fullRow[`${fu1Prefix}: Correct Answer`] = row[`${fu1Prefix}: Correct Answer`] || "";
      fullRow[`${fu1Prefix}: Ranking Logic`] = row[`${fu1Prefix}: Ranking Logic`] || "FALSE";
      fullRow[`${fu1Prefix}: Track Question`] = row[`${fu1Prefix}: Track Question`] || "FALSE";
      fullRow[`${fu1Prefix}: Image/File URL`] = row[`${fu1Prefix}: Image/File URL`] || "";

      for (let level2 = 1; level2 <= 3; level2++) {
        const fu2Prefix = `FU${level1}.${level2}`;
        fullRow[`${fu2Prefix}: Option`] = row[`${fu2Prefix}: Option`] || "";
        fullRow[`${fu2Prefix}: Question Type`] = row[`${fu2Prefix}: Question Type`] || "";
        fullRow[`${fu2Prefix}: Required`] = row[`${fu2Prefix}: Required`] || "";
        fullRow[`${fu2Prefix}: SubParam1`] = row[`${fu2Prefix}: SubParam1`] || "";
        fullRow[`${fu2Prefix}: SubParam2`] = row[`${fu2Prefix}: SubParam2`] || "";
        fullRow[`${fu2Prefix}: Question Text`] = row[`${fu2Prefix}: Question Text`] || "";
        fullRow[`${fu2Prefix}: Description`] = row[`${fu2Prefix}: Description`] || "";
        fullRow[`${fu2Prefix}: Options`] = row[`${fu2Prefix}: Options`] || "";
        fullRow[`${fu2Prefix}: Correct Answer`] = row[`${fu2Prefix}: Correct Answer`] || "";
        fullRow[`${fu2Prefix}: Ranking Logic`] = row[`${fu2Prefix}: Ranking Logic`] || "FALSE";
        fullRow[`${fu2Prefix}: Track Question`] = row[`${fu2Prefix}: Track Question`] || "FALSE";
        fullRow[`${fu2Prefix}: Image/File URL`] = row[`${fu2Prefix}: Image/File URL`] || "";

        for (let level3 = 1; level3 <= 2; level3++) {
          const fu3Prefix = `FU${level1}.${level2}.${level3}`;
          fullRow[`${fu3Prefix}: Option`] = row[`${fu3Prefix}: Option`] || "";
          fullRow[`${fu3Prefix}: Question Type`] = row[`${fu3Prefix}: Question Type`] || "";
          fullRow[`${fu3Prefix}: Required`] = row[`${fu3Prefix}: Required`] || "";
          fullRow[`${fu3Prefix}: SubParam1`] = row[`${fu3Prefix}: SubParam1`] || "";
          fullRow[`${fu3Prefix}: SubParam2`] = row[`${fu3Prefix}: SubParam2`] || "";
          fullRow[`${fu3Prefix}: Question Text`] = row[`${fu3Prefix}: Question Text`] || "";
          fullRow[`${fu3Prefix}: Description`] = row[`${fu3Prefix}: Description`] || "";
          fullRow[`${fu3Prefix}: Options`] = row[`${fu3Prefix}: Options`] || "";
          fullRow[`${fu3Prefix}: Correct Answer`] = row[`${fu3Prefix}: Correct Answer`] || "";
          fullRow[`${fu3Prefix}: Ranking Logic`] = row[`${fu3Prefix}: Ranking Logic`] || "FALSE";
          fullRow[`${fu3Prefix}: Track Question`] = row[`${fu3Prefix}: Track Question`] || "FALSE";
          fullRow[`${fu3Prefix}: Image/File URL`] = row[`${fu3Prefix}: Image/File URL`] || "";
        }
      }
    }

    templateData.push(fullRow);
  });

  const headerArray = [...mainHeaders, ...followUpHeaders];

  const worksheet = utils.json_to_sheet(templateData, {
    header: headerArray,
  });

  worksheet["!cols"] = headerArray.map(() => ({ wch: 25 }));

  // Apply styling to header row (row 1)
  const HEADER_ROW = 1; // Excel is 1-indexed, row 1 is header

  headerArray.forEach((header, colIndex) => {
    const cellAddress = utils.encode_cell({ r: HEADER_ROW - 1, c: colIndex });

    if (!worksheet[cellAddress]) {
      worksheet[cellAddress] = {};
    }

    // Default to main header color
    let color = COLORS.MAIN;

    // Check for follow-up headers and apply appropriate colors
    if (header.includes("FU")) {
      const fuMatch = header.match(/^FU(\d+)(?:\.(\d+)(?:\.(\d+))?)?/);

      if (fuMatch) {
        const [, level1, level2, level3] = fuMatch;

        if (level3) {
          // Level 3 headers (e.g., FU1.1.1)
          color =
            level1 === "1"
              ? COLORS.FU_LIGHT
              : level1 === "2"
                ? COLORS.FU_LIGHT_GREEN
                : getColorForFULevel(level1, 3);
        } else if (level2) {
          // Level 2 headers (e.g., FU1.1)
          color =
            level1 === "1"
              ? COLORS.FU_MEDIUM
              : level1 === "2"
                ? COLORS.FU_MEDIUM_GREEN
                : getColorForFULevel(level1, 2);
        } else {
          // Level 1 headers (e.g., FU1)
          color =
            level1 === "1"
              ? COLORS.FU_DARK
              : level1 === "2"
                ? COLORS.FU_DARK_GREEN
                : getColorForFULevel(level1, 1);
        }
      }
    }
    worksheet[cellAddress].s = {
      fill: {
        patternType: "solid",
        ...color,
      },
      font: {
        bold: true,
        color: { rgb: getContrastTextColor(color.fgColor.rgb) },
      },
    };
  });

  templateData.forEach((row, idx) => {
    const mergeInstructions = row["Section Merging"];
    if (mergeInstructions && typeof mergeInstructions === "string") {
      // Parse merge instructions like "C5:F10"
      try {
        worksheet["!merges"].push(utils.decode_range(mergeInstructions));
      } catch (e) {
        console.warn(`Failed to parse merge instruction: ${mergeInstructions}`);
      }
    }
  });

  // Add data validation for SubParam columns
  const dataValidations: any[] = [];

  // Main SubParam columns
  const mainSubParam1Idx = mainHeaders.indexOf("SubParam1");
  const mainSubParam2Idx = mainHeaders.indexOf("SubParam2");

  if (mainSubParam1Idx !== -1) {
    const col = utils.encode_col(mainSubParam1Idx);
    dataValidations.push({
      sqref: `${col}5:${col}1000`,
      type: "list",
      formula1: "=Parameters!$A$4:$A$1000",
    });
  }

  if (mainSubParam2Idx !== -1) {
    const col = utils.encode_col(mainSubParam2Idx);
    dataValidations.push({
      sqref: `${col}5:${col}1000`,
      type: "list",
      formula1: "=Parameters!$A$4:$A$1000",
    });
  }

  // Follow-up SubParam columns (Level 1 only for validation to keep it simple but covering most cases)
  for (let i = 1; i <= 5; i++) {
    const fuSubParam1Idx = headerArray.indexOf(`FU${i}: SubParam1`);
    const fuSubParam2Idx = headerArray.indexOf(`FU${i}: SubParam2`);

    if (fuSubParam1Idx !== -1) {
      const col = utils.encode_col(fuSubParam1Idx);
      dataValidations.push({
        sqref: `${col}5:${col}1000`,
        type: "list",
        formula1: "=Parameters!$A$4:$A$1000",
      });
    }

    if (fuSubParam2Idx !== -1) {
      const col = utils.encode_col(fuSubParam2Idx);
      dataValidations.push({
        sqref: `${col}5:${col}1000`,
        type: "list",
        formula1: "=Parameters!$A$4:$A$1000",
      });
    }
  }

  worksheet["!datavalidation"] = dataValidations;

  // Create Parameters sheet
  const parametersHeaders = ["Parameter Name", "Type"];
  const parametersDescriptions = [
    "Name of the parameter",
    "Type: Main or Followup",
  ];

  const parametersHeaderRow = {
    "Parameter Name": "Parameter Name",
    Type: "Type",
  };

  const parametersDescriptionRow = {
    "Parameter Name": parametersDescriptions[0],
    Type: parametersDescriptions[1],
  };

  const parametersSeparatorRow = {
    "Parameter Name": "",
    Type: "",
  };

  const parametersSampleData = [
    { "Parameter Name": "Eligibility", Type: "Main" },
    { "Parameter Name": "Document Verification", Type: "Main" },
    { "Parameter Name": "Service History", Type: "Main" },
    { "Parameter Name": "Quality Assessment", Type: "Followup" },
    { "Parameter Name": "Feedback Collection", Type: "Followup" },
  ];

  const parametersData = [
    parametersHeaderRow,
    parametersDescriptionRow,
    parametersSeparatorRow,
    ...parametersSampleData,
  ];

  const parametersWorksheet = utils.json_to_sheet(parametersData, {
    header: parametersHeaders,
  });

  parametersWorksheet["!cols"] = parametersHeaders.map(() => ({ wch: 25 }));

  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, parametersWorksheet, "Parameters");
  utils.book_append_sheet(workbook, worksheet, "Form Template");
  writeFile(workbook, "form-import-template-nested-followups.xlsx");
}



/**
 * Updated downloadFormImportTemplate with image embedding.
 *
 * HOW TO USE
 * ----------
 * 1. Import the base64 strings from `opsFormImages.ts` (generated separately):
 *
 *      import {
 *        logoBImageBase64,
 *        stopCallWaitImageBase64,
 *        noSymbolImageBase64,
 *        ppeGuideImageBase64,
 *        fiveSImageBase64,
 *        qrCodeImageBase64,
 *      } from "./opsFormImages";
 *
 * 2. Pass them as the `images` argument:
 *
 *      downloadFormImportTemplate({
 *        logoBImageBase64,
 *        stopCallWaitImageBase64,
 *        noSymbolImageBase64,
 *        ppeGuideImageBase64,
 *        fiveSImageBase64,
 *        qrCodeImageBase64,
 *      });
 *
 * IMAGE PLACEMENT MAP
 * -------------------
 *  logoBImageBase64        → B2:E11   (top-left company logo box)
 *  stopCallWaitImageBase64 → B66:T69  (Abnormality handling section)
 *  noSymbolImageBase64     → V14:AD14 (No-symbol icon, top of PPE area)
 *  ppeGuideImageBase64     → V14:AD21 (Full PPE guide image)
 *  fiveSImageBase64        → BL13:BW21 (5S Guidelines box)
 *  qrCodeImageBase64       → CE9:CM11  (QR Code area)
 *
 * IMPORTANT: xlsx-js-style / SheetJS Pro image support
 * ----------------------------------------------------
 * Images are added via  ws["!images"]  array.
 * Each entry is:
 *   {
 *     "!pos": { r: rowStart, c: colStart, R: rowEnd, C: colEnd },  // 0-indexed
 *     "!datatype": "base64",
 *     "!type": "jpeg",          // or "png"
 *     "!data": "<base64string>",
 *   }
 * This is supported in xlsx-js-style ≥ 1.2.0.
 * If you are on the free SheetJS (xlsx) community edition, images are NOT
 * supported – you need SheetJS Pro or xlsx-js-style.
 */


// ─── Image option types ──────────────────────────────────────────────────────

export interface OpsFormImages {
  /** Red B company logo  →  B2:E11 */
  logoBImageBase64?: string;
  /** STOP / CALL / WAIT poster  →  B66:T69 */
  stopCallWaitImageBase64?: string;
  /** No-entry circle (black bg)  →  overlaid on PPE area top-left */
  noSymbolImageBase64?: string;
  /** PPE uniform + gloves guide  →  V14:AD21 */
  ppeGuideImageBase64?: string;
  ppeGlovesImageBase64?: string;
  /** 5S Sort / Set / Shine poster  →  BL13:BW21 */
  fiveSImageBase64?: string;
  /** QR Code placeholder  →  CE9:CM11 */
  qrCodeImageBase64?: string;
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function col(letter: string): number {
  return utils.decode_col(letter);
}

function encodeCell(c: number, r: number): string {
  return utils.encode_cell({ c, r });
}

function bs(style: string | null) {
  return style ? { style } : {};
}

function borderStyle(
  top: string | null = null,
  bottom: string | null = null,
  left: string | null = null,
  right: string | null = null
) {
  return {
    top: bs(top),
    bottom: bs(bottom),
    left: bs(left),
    right: bs(right),
  };
}

function cellStyle(
  bold = false,
  fontSize = 11,
  hAlign = "left",
  vAlign = "center",
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

function setCell(
  ws: Record<string, any>,
  cellAddr: string,
  value: any,
  style: object
) {
  if (!ws[cellAddr]) ws[cellAddr] = {};
  ws[cellAddr].v = value;
  ws[cellAddr].t = typeof value === "number" ? "n" : "s";
  ws[cellAddr].s = style;
}

function mergeAndSet(
  ws: Record<string, any>,
  range: string,
  value: any,
  style: object
) {
  if (!ws["!merges"]) ws["!merges"] = [];
  ws["!merges"].push(utils.decode_range(range));

  const topLeft = range.split(":")[0];

  // Parse the range to understand the full dimensions
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

    // Apply border to ALL cells in the merged range (not just top-left)
    // This ensures borders show correctly in Excel
    for (let r = startRow - 1; r <= endRow - 1; r++) {
      for (let c = startColIdx; c <= endColIdx; c++) {
        const cellAddr = encodeCell(c, r);
        if (!ws[cellAddr]) ws[cellAddr] = { v: "", t: "s" };

        // Only set the value on the top-left cell
        if (r === startRow - 1 && c === startColIdx) {
          ws[cellAddr].v = value;
          ws[cellAddr].t = typeof value === "number" ? "n" : "s";
        }

        // Apply borders based on position in merged range
        const borderObj: any = {};

        // Top border - only for top row of merged range
        if (r === startRow - 1 && (style as any).border?.top) {
          borderObj.top = (style as any).border.top;
        }

        // Bottom border - only for bottom row of merged range
        if (r === endRow - 1 && (style as any).border?.bottom) {
          borderObj.bottom = (style as any).border.bottom;
        }

        // Left border - only for left column of merged range
        if (c === startColIdx && (style as any).border?.left) {
          borderObj.left = (style as any).border.left;
        }

        // Right border - only for right column of merged range
        if (c === endColIdx && (style as any).border?.right) {
          borderObj.right = (style as any).border.right;
        }

        // Apply other style properties (fill, font, alignment) only to top-left
        if (r === startRow - 1 && c === startColIdx) {
          ws[cellAddr].s = {
            ...style,
            border: borderObj
          };
        } else {
          // For other cells in merged range, only apply borders
          ws[cellAddr].s = { border: borderObj };
        }
      }
    }
  } else {
    setCell(ws, topLeft, value, style);
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

      // Top border
      if (r === startRow) border.top = { style: borderStyle };
      // Bottom border
      if (r === endRow) border.bottom = { style: borderStyle };
      // Left border
      if (c === startCol) border.left = { style: borderStyle };
      // Right border
      if (c === endCol) border.right = { style: borderStyle };

      ws[addr].s = { border };
    }
  }
}

// ─── Image helper ────────────────────────────────────────────────────────────

/**
 * Add an image to the worksheet using xlsx-js-style's !images API.
 *
 * @param ws        worksheet object
 * @param base64    raw base64 string (no data-URL prefix)
 * @param type      image mime sub-type: "jpeg" | "png" | "gif"
 * @param colStart  0-indexed start column  (e.g. col("B") = 1)
 * @param rowStart  0-indexed start row     (e.g. row 2 in Excel = index 1)
 * @param colEnd    0-indexed end column (inclusive)
 * @param rowEnd    0-indexed end row    (inclusive)
 */
function addImage(
  ws: Record<string, any>,
  base64: string,
  type: "jpeg" | "png" | "gif",
  colStart: number,
  rowStart: number,
  colEnd: number,
  rowEnd: number
) {
  if (!ws["!images"]) ws["!images"] = [];
  ws["!images"].push({
    "!pos": { r: rowStart, c: colStart, R: rowEnd, C: colEnd },
    "!datatype": "base64",
    "!type": type,
    "!data": base64,
  });
}

// ─── Fetch image from public folder as raw base64 ────────────────────────────
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
        // ✅ Sanity check — should NOT start with "data:"
        if (raw.startsWith("data:")) {
          console.error("[fetchAssetBase64] Strip failed for", path);
        }
        resolve(raw);
      };
      reader.readAsDataURL(blob);
    });
  } catch {
    return "";
  }
}

const noBorder = { left: {}, right: {}, top: {}, bottom: {} };
const noBorderStyle = { border: noBorder };
// New helper: only draw the 4 outer edges of a range, inner cells get NO borders
function clearInnerBorders(
  ws: Record<string, any>,
  startCol: number,
  endCol: number,
  startRow: number,
  endRow: number
) {
  // First, set ALL cells in range to have NO borders
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const addr = encodeCell(c, r);
      if (!ws[addr]) ws[addr] = { v: "", t: "s" };
      ws[addr].s = { ...(ws[addr].s || {}), border: {} };
    }
  }

  // Then add ONLY outer borders
  // Top row
  for (let c = startCol; c <= endCol; c++) {
    const addr = encodeCell(c, startRow);
    ws[addr].s.border.top = { style: "thin" };
  }

  // Bottom row
  for (let c = startCol; c <= endCol; c++) {
    const addr = encodeCell(c, endRow);
    ws[addr].s.border.bottom = { style: "thin" };
  }

  // Left column
  for (let r = startRow; r <= endRow; r++) {
    const addr = encodeCell(startCol, r);
    ws[addr].s.border.left = { style: "thin" };
  }

  // Right column
  for (let r = startRow; r <= endRow; r++) {
    const addr = encodeCell(endCol, r);
    ws[addr].s.border.right = { style: "thin" };
  }
}


// ─── Main export ─────────────────────────────────────────────────────────────
export async function downloadFormImportTemplate(images: OpsFormImages = {}, parsedData?: any, config: OpsTemplateConfig = {}) {
  console.log("[Template] Starting to fetch images...");

  const [
    logoBImageBase64,
    stopCallWaitImageBase64,
    noSymbolImageBase64,
    noMobileImageBase64,
    ppeGuideImageBase64,
    ppeGlovesImageBase64,  // NEW
    fiveSImageBase64,
    qrCodeImageBase64,
    shiftTimingImageBase64,
  ] = await Promise.all([
    images.logoBImageBase64 ?? fetchAssetBase64("/assets/Companylogo.png"),
    images.stopCallWaitImageBase64 ?? fetchAssetBase64("/assets/Safetyposter.png"),
    images.noSymbolImageBase64 ?? fetchAssetBase64("/assets/Dontrun.png"),
    fetchAssetBase64("/assets/dontusemobile.png"),
    images.ppeGuideImageBase64 ?? fetchAssetBase64("/assets/PPEGuide.png"),
    images.ppeGlovesImageBase64 ?? fetchAssetBase64("/assets/PPEGUIDE2.png"),  // NEW - add this line
    images.fiveSImageBase64 ?? fetchAssetBase64("/assets/5S_Guidelines.png"),
    images.qrCodeImageBase64 ?? fetchAssetBase64("/assets/Qrcode.png"),
    fetchAssetBase64("/assets/Shift_timing.png"),
  ]);

  console.log("[Template] Images fetched:", {
    logo: logoBImageBase64 ? `${logoBImageBase64.length} chars` : "MISSING",
    safety: stopCallWaitImageBase64 ? `${stopCallWaitImageBase64.length} chars` : "MISSING",
    dontRun: noSymbolImageBase64 ? `${noSymbolImageBase64.length} chars` : "MISSING",
    noMobile: noMobileImageBase64 ? `${noMobileImageBase64.length} chars` : "MISSING",
    ppe: ppeGuideImageBase64 ? `${ppeGuideImageBase64.length} chars` : "MISSING",
    fiveS: fiveSImageBase64 ? `${fiveSImageBase64.length} chars` : "MISSING",
    qr: qrCodeImageBase64 ? `${qrCodeImageBase64.length} chars` : "MISSING",
    shiftTiming: shiftTimingImageBase64 ? `${shiftTimingImageBase64.length} chars` : "MISSING",
  });

  // ─── Build OPS worksheet ───────────────────────────────────────────────────

  const ws: Record<string, any> = {};
  ws["!ref"] = "B1:CM77";
  ws["!merges"] = [];

  // Row heights (0-indexed internally)
  ws["!rows"] = [];
  const rowHeights: Record<number, number> = {
    0: 27, 1: 39, 2: 39, 3: 39, 4: 39, 5: 39, 6: 39,
    7: 45, 8: 45, 9: 45, 10: 45,
    11: 33.75, 12: 37.5, 13: 37.5, 14: 37.5, 15: 47.25,
    16: 37.5, 17: 37.5, 18: 37.5, 19: 37.5, 20: 37.5,
    21: 15, 22: 54.75, 23: 54.75,
    ...Object.fromEntries(Array.from({ length: 40 }, (_, i) => [24 + i, 62.15])),
    64: 15,
    65: 25, 66: 46.5, 67: 47.25, 68: 29.25, 69: 14.25,
    70: 29.25, 71: 29.25, 72: 29.25,
    73: 29.25, 74: 29.25, 75: 29.25,
    76: 39,
  };
  for (const [r, h] of Object.entries(rowHeights)) {
    ws["!rows"][Number(r)] = { hpt: h };
  }

  // Column widths (0-indexed)
  const colWidthMap: Record<string, number> = {
    A: 15.27, B: 4.54, D: 5.73, E: 22.73, F: 4.54, G: 8.0,
    I: 7.73, J: 8.73, M: 4.54, O: 9.27, P: 7.54, Q: 6.45,
    R: 13.18, S: 11.82, T: 13.82, U: 10.27, V: 11.54, W: 12.27,
    X: 4.54, AD: 8.18, AE: 7.18, AF: 7.45, AG: 4.54, AH: 8.0,
    AI: 6.82, AJ: 9.27, AK: 14.82, AS: 4.27, AX: 9.45, AY: 23.0,
    BH: 30.82, BI: 8.45, BJ: 22.54, BK: 23.54, BL: 7.27,
    BM: 6.73, BN: 8.45, BP: 13.45, BQ: 11.0, BR: 15.18,
    BS: 17.0, BT: 10.82, BU: 5.18, BW: 11.82, BX: 5.18,
    BZ: 10.18, CA: 5.82, CB: 3.82, CD: 6.73, CE: 11.0,
    CG: 6.45, CI: 5.73, CJ: 9.54, CK: 9.27, CL: 12.0,
    CM: 5.45,
  };
  ws["!cols"] = [];
  for (const [letter, width] of Object.entries(colWidthMap)) {
    ws["!cols"][col(letter)] = { wch: width };
  }

  // ── ROW 1: Retention Period ────────────────────────────────────────────────
  mergeAndSet(ws, "B1:CM1",
    "Retention Period : 20 years after Model is discontinued",
    cellStyle(true, 18, "right", "center", true, "medium", null, "medium", "medium")
  );

  // ── ROWS 2-11: Top header block ────────────────────────────────────────────

  // Left image box B2:E11 — keep border cells, image will overlay
  clearInnerBorders(ws, col("B"), col("E"), 1, 10);

  // Dept / Section
  mergeAndSet(ws, "F2:I4", config.basicInfoLabels?.deptSection || "Dept. / Section :",
    cellStyle(true, 20, "left", "center", true, "medium", "thin", "medium", null));
  mergeAndSet(ws, "J2:L4", "",
    cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", "thin"));
  mergeAndSet(ws, "M2:R4", config.basicInfoLabels?.lineZone || "Line / Zone :",
    cellStyle(true, 20, "left", "center", false, "medium", "thin", "thin", "thin"));
  mergeAndSet(ws, "S2:W4", "",
    cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", "medium"));

  mergeAndSet(ws, "X2:BA4", "Operation Standard ",
    cellStyle(true, 72, "center", "center", false, "medium", "medium", "medium", null));

  mergeAndSet(ws, "F5:I7", config.basicInfoLabels?.model || "Model :",
    cellStyle(true, 20, "left", "center", true, "thin", "thin", "medium", "thin"));
  mergeAndSet(ws, "J5:L7", "",
    cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "M5:R7", config.basicInfoLabels?.processStation || "Process / Station :",
    cellStyle(true, 20, "left", "center", true, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "S5:W7", "",
    cellStyle(true, 20, "center", "center", true, "thin", "thin", "thin", "medium"));
  mergeAndSet(ws, "AH5:BA5", "Your Work When Trouble Stopped The Production Line",
    cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", null));

  mergeAndSet(ws, "AH6:AI6", "S.No.",
    cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "AJ6:AR6", "Trouble",
    cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "AS6:BA6", "Your task",
    cellStyle(true, 20, "center", "center", false, "thin", "thin", "thin", null));

  const troubles: [number, number, string, string][] = config.troubleTasks && config.troubleTasks.length > 0
    ? config.troubleTasks.map((t, i) => [7 + i, t.sno, t.trouble, t.task] as [number, number, string, string])
    : [
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

  mergeAndSet(ws, "F8:J11",
    config.rejectionHandling || "REJECTION HANDLING :-\n\nClearly Identify Rejected / NG parts. Keep them properly with proper identification at defined Location.",
    cellStyle(false, 18, "left", "center", true, "medium", "medium", "medium", "medium"));
  mergeAndSet(ws, "K8:O11", "Measuring Instruments or Gauges ",
    cellStyle(true, 18, "center", "center", true, "medium", "medium", null, "thin"));
  const defaultMeasuringInst = [
    "Always use Calibrated Measuring Instruments / Gauges (Ensure Calibration status before using the same).",
    "Ensure Zero setting before use.",
    "Do Not Use Unidentified Measuring Tool / Gauges.",
    "In case of any abnormality, inform Line leader and Quality Engineer to take action for suspected NG material range."
  ];
  const measuringInst = config.measuringInstruments || defaultMeasuringInst;
  mergeAndSet(ws, "P8:W8", measuringInst[0] || "",
    cellStyle(false, 18, "left", "center", true, "medium", "thin", "thin", "medium"));
  mergeAndSet(ws, "P9:W9", measuringInst[1] || "",
    cellStyle(false, 18, "left", "center", true, "thin", "thin", "thin", "medium"));
  mergeAndSet(ws, "P10:W10", measuringInst[2] || "",
    cellStyle(false, 18, "left", "center", true, "thin", "thin", "thin", "medium"));
  mergeAndSet(ws, "P11:W11", measuringInst[3] || "",
    cellStyle(false, 18, "left", "center", true, "thin", "medium", "thin", "medium"));

  fillBorderRange(ws, col("BB"), col("BE"), 1, 7, "medium");
  fillBorderRange(ws, col("BF"), col("BI"), 1, 7, "medium");
  fillBorderRange(ws, col("BJ"), col("BL"), 1, 7, "medium");
  mergeAndSet(ws, "BB9:BE11", "Prepared",
    cellStyle(true, 28, "center", "center", false, null, "medium", "medium", "medium"));
  mergeAndSet(ws, "BF9:BI11", "Checked",
    cellStyle(true, 28, "center", "center", false, null, "medium", "medium", "medium"));
  mergeAndSet(ws, "BJ9:BL11", "Approved",
    cellStyle(true, 28, "center", "center", false, null, "medium", null, "medium"));

  for (let r = 2; r <= 11; r++) {
    const topB = r === 2 ? "medium" : "thin";
    const botB = r === 11 ? null : "thin";
    mergeAndSet(ws, `BM${r}:BN${r}`, r === 11 ? "No." : "",
      cellStyle(r === 11, 20, "center", "center", false, topB, botB, "medium", "thin"));
    mergeAndSet(ws, `BO${r}:BS${r}`, r === 11 ? " DD /MM/ YY" : "",
      cellStyle(r === 11, 20, "center", "center", false, topB, r === 11 ? "medium" : botB, "thin", "thin"));
    mergeAndSet(ws, `BT${r}:CD${r}`, r === 11 ? "Issuance / Revision details" : "",
      cellStyle(r === 11, 20, "center", "center", false, topB, botB, "thin", "thin"));
  }

  mergeAndSet(ws, "CE2:CM4", (config.basicInfoLabels?.formatNo || "Format No.  : ") + "q_format_no",
    cellStyle(true, 20, "left", "center", false, "medium", "thin", "thin", "medium"));
  mergeAndSet(ws, "CE5:CM8", (config.basicInfoLabels?.controlNo || "Control No. : ") + "q_control_no",
    cellStyle(true, 20, "left", "center", false, "thin", "thin", "thin", "medium"));
  // QR Code cell
  mergeAndSet(ws, "CE9:CM11", "q_qr_code:",
    cellStyle(true, 20, "left", "center", false, "thin", null, "thin", "medium"));

  // ── ROW 12: Section header bar ─────────────────────────────────────────────
  mergeAndSet(ws, "B12:BC12", "General Instructions",
    cellStyle(true, 20, "center", "center", true, "medium", "thin", "medium", "thin"));
  mergeAndSet(ws, "BD12:BK12", "EMS & Safety Guidelines",
    cellStyle(true, 20, "center", "center", false, "medium", "thin", "thin", "thin"));
  mergeAndSet(ws, "BL12:BW12", "5S Guidelines",
    cellStyle(true, 20, "center", "center", true, "medium", "thin", "thin", "thin"));
  mergeAndSet(ws, "BX12:CM12", "Process Instructions ",
    cellStyle(true, 20, "center", "center", true, "medium", "thin", "medium", "thin"));

  // ── ROW 13: Sub-headers ────────────────────────────────────────────────────
  mergeAndSet(ws, "B13:L13", "FIFO System",
    cellStyle(true, 20, "center", "center", true, "thin", "thin", "medium", "thin"));
  mergeAndSet(ws, "M13:U13", "Non Lubrication Rule: ",
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

  // ── ROWS 14-21: General Instructions content ───────────────────────────────
  mergeAndSet(ws, "B14:L21",
    "1. Bin/trolley must be changed only after complete usage of all material in it. \n" +
    "2. Empty bin/trolley should be replaced with new one \n" +
    "3. Don't top up partially filled bin\n" +
    "4.Follow FIFO on line during Process .\n" +
    "5.Do not use next bin / Trolley material until running not consumed.",
    cellStyle(false, 20, "left", "center", true, "thin", "thin", "medium", "thin"));
  mergeAndSet(ws, "M14:U15",
    "Do not use any lubrication if not specified in OPS / Process Sheet.",
    cellStyle(false, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "M16:R16", "Do not use mobile on the shopfloor",
    cellStyle(false, 20, "center", "center", true, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "S16:U16", "Do not run on the shopfloor",
    cellStyle(false, 20, "center", "center", true, "thin", "thin", "thin", "thin"));

  for (let r = 17; r <= 21; r++) {
    mergeAndSet(ws, `M${r}:R${r}`, "",
      cellStyle(false, 20, "left", "center", false, "thin", r === 21 ? "medium" : "thin", "thin", "thin"));
    mergeAndSet(ws, `S${r}:U${r}`, "",
      cellStyle(false, 20, "left", "center", false, "thin", r === 21 ? "medium" : "thin", "thin", "thin"));
  }

  // PPE area borders — images will overlay these ranges
  clearInnerBorders(ws, col("AM"), col("BC"), 13, 20);
  // 3. Always wear PPEs / Proper uniform  V14:AD21  (col 21–29, row 13–20)
  clearInnerBorders(ws, col("V"), col("AD"), 13, 20);

  // 4. Wear PPEs as per your station's requirements  AE14:AL21  (col 30–37, row 13–20)
  clearInnerBorders(ws, col("AE"), col("AL"), 13, 20);

  mergeAndSet(ws, "BD14:BG21",
    "1. Do waste segregation.\n2. Switch off idle lights & machines\n" +
    "3. Ensure 3R Principal in daily activities\n4. If there was any leakage, communicate to Sub Leader",
    cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "BH14:BK21",
    "1. Follow POS sheet in case of any Chemical\n2. Follow MSDS/SDS in case of any emergency regarding chemical\n3. Follow your PPE's",
    cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));

  // 5S box border — image will overlay
  fillBorderRange(ws, col("BL"), col("BW"), 12, 20, "thin");

  const processInstructions = config.processInstructions && config.processInstructions.length > 0 ? config.processInstructions : [
    "2. Do Not Use Fallen Electrical/Functional Parts.",
    "3. Ensure Model / Variant Change.",
    "4. Report in case of part / hardware fallen inside vehicle.",
    "5. TQ Wrench Arrow Mark should be in correct direction.",
    "6. Put Fallen Hardware in Red Bin for Zone In-Charge judgement.",
    "7. Take approval from SH / HOD before changing process sequence.",
    "8. Zone In-Charge is overall responsible to ensure work is as per OPS.",
    "9. Contaminant parts should be covered properly.",
  ];
  processInstructions.forEach((text, i) => {
    const r = 14 + i;
    mergeAndSet(ws, `BX${r}:CM${r}`, text,
      cellStyle(false, 20, "left", "center", true,
        "thin", r === 21 ? "medium" : "thin", "medium", "thin"));
  });

  // ── ROW 22: Separator ──────────────────────────────────────────────────────
  mergeAndSet(ws, "B22:CM22", "", cellStyle(false, 11, "left", "center", false, null, null, "medium", "medium"));


  mergeAndSet(
    ws,
    "B23:U64",
    config.illustrationsQuestion || "Upload Illustrations / Process Images Here",
    cellStyle(
      true,
      20,
      "center",
      "center",
      true,
      "medium",
      "medium",
      "medium",
      "thin"
    )
  );

  const firstStep = parsedData?.processSteps?.[0] || {};
  const th = config.tableHeaders || {};
  const colHeaders: [string, string][] = [
    ["V23:V24", String(firstStep.sn || th.sn || "SN")],
    ["W23:Z24", firstStep.importance || th.itemImportance || "Item Importance"],
    ["AA23:AJ24", firstStep.activity || th.stepWhat || "स्टेप\n (What / Activity)"],
    ["AK23:AR24", firstStep.method || th.methodHow || "Method \n(How)"],
    ["AS23:AX24", firstStep.frequency || th.frequencyWhen || "Frequency / When"],
    ["AY23:BG24", firstStep.standard || th.standardCriteria || "Standard \n(Spec. / Judgment Criteria)"],
    ["BH23:BH24", firstStep.responsibility || th.responsibility || "Responsibility"],
    ["BI23:BJ24", firstStep.equipment || th.equipmentMeasuring || "Equipment / Measuring Eq."],
    ["BK23:BO24", firstStep.abnormalities || th.possibleAbnormalities || "Possible Abnormalities"],
    ["BP23:BS24", firstStep.reactionPlan || th.reactionPlan || "Reaction Plan"],
    ["BT23:BY24", firstStep.partNameQty || th.partName || "Part Name & QTY"],
    ["BZ23:CD24", firstStep.ppe || th.ppeRequired || "PPEs required"],
    ["CE23:CI24", firstStep.recordDocument || th.recordDocument || "Record / Document"],
    ["CJ23:CM24", firstStep.remarks || th.remarks || "Remarks"],
  ];
  for (const [range, label] of colHeaders) {
    const isLast = range.startsWith("CJ");
    mergeAndSet(ws, range, label,
      cellStyle(true, 22, "center", "center", true,
        "medium", "thin", "thin", isLast ? "medium" : "thin"));
  }

  // ── ROWS 25-64: 5 step blocks ──────────────────────────────────────────────
  interface StepData {
    sn: string;
    star: string;
    enStep: string; enMethod: string; enFreq: string; enStd: string;
    enResp: string; enEquip: string; enAbn: string; enReact: string;
    enPart: string; enQty: string; enPpe: string; enRec: string; enRem: string;
    hiStep: string; hiMethod: string; hiFreq: string; hiStd: string;
    hiResp: string; hiEquip: string; hiAbn: string; hiReact: string;
    hiPpe: string; hiRec: string; hiRem: string;
  }

  const steps: StepData[] = [1, 2, 3, 4, 5].map((n) => ({
    sn: String(n),              // Keep "1", "2", "3", "4", "5"
    star: "",                   // Empty (was "☆")
    enStep: "",                 // Empty (was pre-filled activity)
    enMethod: "",               // Empty (was pre-filled method)
    enFreq: "",                 // Empty (was pre-filled frequency)
    enStd: "",                  // Empty (was pre-filled standard)
    enResp: "",                 // Empty (was "OPERATOR")
    enEquip: "",                // Empty (was pre-filled)
    enAbn: "",                  // Empty (was pre-filled)
    enReact: "",                // Empty (was pre-filled)
    enPart: "",                 // Empty
    enQty: "",                  // Empty
    enPpe: "",                  // Empty (was pre-filled)
    enRec: "",                  // Empty (was "SPECIFICATION CHECK SHEET")
    enRem: "",                  // Empty
    hiStep: "",                 // Empty
    hiMethod: "",               // Empty
    hiFreq: "",                 // Empty
    hiStd: "",                  // Empty
    hiResp: "",                 // Empty
    hiEquip: "",                // Empty
    hiAbn: "",                  // Empty
    hiReact: "",                // Empty
    hiPpe: "",                  // Empty
    hiRec: "",                  // Empty
    hiRem: "",                  // Empty
  }));

  steps.forEach((step, idx) => {
    const er = 25 + idx * 8;
    const hr = er + 4;
    const isLast = idx === 4;
    const outerBot = isLast ? "medium" : "thin";

    const dataStyle = (hAlign: string, isHindi: boolean, leftBorder = "thin", rightBorder = "thin") =>
      cellStyle(false, 22, hAlign, "center", true,
        "thin", isHindi ? outerBot : "thin", leftBorder, rightBorder);

    mergeAndSet(ws, `V${er}:V${er + 3}`, step.sn,
      cellStyle(true, 22, "center", "center", false, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, `V${hr}:V${hr + 3}`, "",
      cellStyle(false, 22, "center", "center", false, "thin", outerBot, "thin", "thin"));

    mergeAndSet(ws, `W${er}:Z${er + 3}`, step.star,
      cellStyle(true, 72, "center", "center", true, "thin", "thin", "thin", "thin"));
    mergeAndSet(ws, `W${hr}:Z${hr + 3}`, "",
      cellStyle(false, 22, "center", "center", false, "thin", outerBot, "thin", "thin"));

    mergeAndSet(ws, `AA${er}:AJ${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `AA${hr}:AJ${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `AK${er}:AR${er + 3}`, "", dataStyle("left", false));
    mergeAndSet(ws, `AK${hr}:AR${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `AS${er}:AX${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `AS${hr}:AX${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `AY${er}:BG${er + 3}`, "", dataStyle("left", false));
    mergeAndSet(ws, `AY${hr}:BG${hr + 3}`, "", dataStyle("left", true));
    mergeAndSet(ws, `BH${er}:BH${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `BH${hr}:BH${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `BI${er}:BJ${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `BI${hr}:BJ${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `BK${er}:BO${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `BK${hr}:BO${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `BP${er}:BS${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `BP${hr}:BS${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `BT${er}:BW${er + 7}`, "", dataStyle("center", true));
    mergeAndSet(ws, `BX${er}:BY${er + 7}`, "", dataStyle("center", true));
    mergeAndSet(ws, `BZ${er}:CD${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `BZ${hr}:CD${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `CE${er}:CI${er + 3}`, "", dataStyle("center", false));
    mergeAndSet(ws, `CE${hr}:CI${hr + 3}`, "", dataStyle("center", true));
    mergeAndSet(ws, `CJ${er}:CM${er + 3}`, "", dataStyle("left", false, "thin", "medium"));
    mergeAndSet(ws, `CJ${hr}:CM${hr + 3}`, "", dataStyle("left", true, "thin", "medium"));
  });

  // ── ROW 65: Separator ─────────────────────────────────────────────────────
  mergeAndSet(ws, "B65:CM65", "", cellStyle(false, 11, "left", "center", false, null, null, "medium", "medium"));

  // ── ROWS 66-69: Abnormality / Past Problems ───────────────────────────────
  mergeAndSet(ws, "B66:T69",
    config.abnormalityHandlingRoute || "Abnormality handling route : \nIn case of any abnormality inform the Zone In-Charge\n" +
    "Flow of Communication :-\nOperator  ► Team Member ► Section Mgr ► As required",
    cellStyle(true, 26, "left", "center", true, "medium", "thin", "medium", "thin"));
  mergeAndSet(ws, "U66:CM69", config.abnormalityDetailsLabel || "Past Problem Details",
    cellStyle(true, 26, "center", "top", false, "medium", "thin", "thin", "medium"));

  for (let r = 67; r <= 69; r++) {
    mergeAndSet(ws, `B${r}:T${r}`, "",
      cellStyle(false, 11, "left", "center", false, "thin", r === 69 ? "medium" : "thin", "medium", "thin"));
    mergeAndSet(ws, `U${r}:CM${r}`, "",
      cellStyle(false, 11, "left", "center", false, "thin", r === 69 ? "medium" : "thin", "thin", "medium"));
  }

  // ── ROW 70: Separator ─────────────────────────────────────────────────────
  mergeAndSet(ws, "B70:CM70", "", cellStyle(false, 11, "left", "center", false, null, null, "medium", "medium"));

  // ── ROWS 71-76: Associate / Sign rows ─────────────────────────────────────
  mergeAndSet(ws, "B71:G73", config.associateSignArea?.title1 || "Associate Name \n& Emp. Code",
    cellStyle(true, 26, "left", "center", true, "medium", "thin", "medium", "thin"));
  mergeAndSet(ws, "B74:G76", config.associateSignArea?.title2 || "Sign & Date",
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

  // ── ROW 77: Page Number ───────────────────────────────────────────────────
  mergeAndSet(ws, "B77:CD77", "",
    cellStyle(false, 11, "left", "center", false, "medium", "medium", "medium", null));
  mergeAndSet(ws, "CE77:CM77", "Page Number : XX / XX",
    cellStyle(true, 26, "center", "center", false, "medium", "medium", null, "medium"));

  // ─────────────────────────────────────────────────────────────────────────
  // ── IMAGE EMBEDDING ───────────────────────────────────────────────────────
  // All row/col indices below are 0-based (Excel row 2 = index 1, col B = 1)
  // ─────────────────────────────────────────────────────────────────────────

  // Image 2 — Company logo (logo.png)
  // Target cell range: B2:E11  →  cols 1–4, rows 1–10
  if (logoBImageBase64) {
    addImage(ws, logoBImageBase64, "png", col("B"), 1, col("E"), 10);
  }

  // Image 3 — Safety poster / STOP·CALL·WAIT (Safetyposter.png)
  // Target cell range: B66:T69  →  cols 1–19, rows 65–68
  if (stopCallWaitImageBase64) {
    addImage(ws, stopCallWaitImageBase64, "png",
      col("B"), 65,
      col("T"), 68
    );
  }

  // Image 4 — "Don't Run" icon  (Dontrun.png)
  // Placed at left of PPE section: V14:X17
  if (noSymbolImageBase64) {
    addImage(ws, noSymbolImageBase64, "png",
      col("V"), 13,
      col("X"), 16
    );
  }

  // Image 4b — "Don't Use Mobile" icon  (dontusemobile.png)
  // Placed next to no-run icon: Y14:AA17
  if (noMobileImageBase64) {
    addImage(ws, noMobileImageBase64, "png",
      col("Y"), 13,
      col("AA"), 16
    );
  }

  // Image 5 — PPE guide (PPEGUIDE2.png)
  // Target cell range: AE14:AL21  →  cols 30–37, rows 13–20
  if (ppeGuideImageBase64) {
    addImage(ws, ppeGuideImageBase64, "png",
      col("AE"), 13,
      col("AL"), 20
    );
  }

  // Image 5b — Shift Timing table (Shift_timing.png)
  // Target cell range: AM14:BC21
  if (shiftTimingImageBase64) {
    addImage(ws, shiftTimingImageBase64, "png",
      col("AM"), 13,
      col("BC"), 20
    );
  }

  // Image 6 — 5S Sort / Set in Order / Shine poster (5S_Guidelines.png)
  // Target cell range: BL13:BW21  →  cols 63–74, rows 12–20
  if (fiveSImageBase64) {
    addImage(ws, fiveSImageBase64, "png",
      col("BL"), 12,
      col("BW"), 20
    );
  }

  // Image 7 — QR Code placeholder (Qrcode.png)
  // Target cell range: CE9:CM11  →  cols 82–90, rows 8–10
  if (qrCodeImageBase64) {
    addImage(ws, qrCodeImageBase64, "png",
      col("CE"), 8,
      col("CM"), 10
    );
  }

  // ─── Build workbook & save ─────────────────────────────────────────────────
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, ws as any, "Format");

  // Injected Named Ranges for TVS-OPS Format Template mapping
  if (!workbook.Workbook) workbook.Workbook = {};
  if (!workbook.Workbook.Names) workbook.Workbook.Names = [];

  // Basic Info (4)
  workbook.Workbook.Names.push({ Name: "q_dept_section", Ref: "Format!$F$2:$I$4", Sheet: 0 });
  workbook.Workbook.Names.push({ Name: "q_line_zone", Ref: "Format!$M$2:$R$4", Sheet: 0 });
  workbook.Workbook.Names.push({ Name: "q_model", Ref: "Format!$F$5:$I$7", Sheet: 0 });
  workbook.Workbook.Names.push({ Name: "q_process_station", Ref: "Format!$M$5:$R$7", Sheet: 0 });

  // Doc Control (3)
  workbook.Workbook.Names.push({ Name: "q_format_no", Ref: "Format!$CE$2:$CM$4", Sheet: 0 });
  workbook.Workbook.Names.push({ Name: "q_control_no", Ref: "Format!$CE$5:$CM$8", Sheet: 0 });

  // Illustrations (2)
  workbook.Workbook.Names.push({ Name: "q_illustrations_images", Ref: "Format!$B$23:$U$64", Sheet: 0 });

  console.log("[Images] ws['!images']:", ws["!images"]?.length, "images");
  ws["!images"]?.forEach((img: any, i: number) => {
    console.log(`  Image ${i}: type=${img["!type"]}, data length=${img["!data"]?.length}, pos=`, img["!pos"]);
  });
  // Write workbook to buffer (NOT directly to file)
  const xlsxBuffer: ArrayBuffer = write(workbook, { type: "array", bookType: "xlsx" });

  // Build image placements list
  const imagePlacements: ImagePlacement[] = [
    // Image 1 — Company Logo (B2:E11)
    logoBImageBase64 && {
      base64: logoBImageBase64,
      type: "png",
      anchor: "one",
      fromCol: 1, fromColOff: 70000,
      fromRow: 1, fromRowOff: 30000,
      toCol: 4, toColOff: 0,
      toRow: 10, toRowOff: 0,
      cx: 2268378, cy: 2502652,
      name: "Company Logo",
    },

    // Image 2 — Safety Poster (Your Work When Trouble area - rows 4-10, cols 23-32)
    stopCallWaitImageBase64 && {
      base64: stopCallWaitImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 23, fromColOff: 100000,
      fromRow: 4, fromRowOff: 50000,
      toCol: 32, toColOff: 0,
      toRow: 10, toRowOff: 0,
      cx: 3510396, cy: 3528002,
      name: "Safety Poster",
    },

    // Image 3 — Shift Timing (AM14:BC21)
    shiftTimingImageBase64 && {
      base64: shiftTimingImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 38, fromColOff: 60000,
      fromRow: 13, fromRowOff: 30000,
      toCol: 54, toColOff: 0,
      toRow: 20, toRowOff: 0,
      cx: 16310303, cy: 3799417,
      name: "Shift Timing",
    },

    // Image 4 — 5S Guidelines (BL13:BW21)
    fiveSImageBase64 && {
      base64: fiveSImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 63, fromColOff: 130000,
      fromRow: 12, fromRowOff: 80000,
      toCol: 74, toColOff: 0,
      toRow: 20, toRowOff: 0,
      cx: 8266952, cy: 4184651,
      name: "5S Guidelines",
    },

    // Image 5 — Don't Run icon (col 18-20, rows 16-20)
    noSymbolImageBase64 && {
      base64: noSymbolImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 18, fromColOff: 60000,
      fromRow: 16, fromRowOff: 30000,
      toCol: 20, toColOff: 0,
      toRow: 20, toRowOff: 0,
      cx: 2113972, cy: 2135908,
      name: "No Run",
    },

    // Image 6 — Don't Use Mobile icon (col 13-17, rows 16-20)
    noMobileImageBase64 && {
      base64: noMobileImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 13, fromColOff: 60000,
      fromRow: 16, fromRowOff: 30000,
      toCol: 17, toColOff: 0,
      toRow: 20, toRowOff: 0,
      cx: 2116283, cy: 2147454,
      name: "No Mobile",
    },

    // Image 7 — PPE Guide (col 21-29, rows 13-20)
    ppeGuideImageBase64 && {
      base64: ppeGuideImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 21, fromColOff: 80000,
      fromRow: 13, fromRowOff: 50000,
      toCol: 29, toColOff: 0,
      toRow: 20, toRowOff: 0,
      cx: 4028994, cy: 3847875,
      name: "PPE Guide",
    },
    // Image 8 — PPE Guide (col 21-29, rows 13-20)
    ppeGlovesImageBase64 && {
      base64: ppeGlovesImageBase64,
      type: "png",
      anchor: "two",
      fromCol: 30,    // AE = col 30 (0-indexed)
      fromColOff: 80000,
      fromRow: 13,    // row 14 (0-indexed = 13)
      fromRowOff: 50000,
      toCol: 37,      // AL = col 37 (0-indexed)
      toColOff: 0,
      toRow: 20,      // row 21 (0-indexed = 20)
      toRowOff: 0,
      cx: 3777210,
      cy: 3862213,
      name: "PPE Gloves Image",
    },
    qrCodeImageBase64 && {
      base64: qrCodeImageBase64,
      type: "png",
      anchor: "one",      // Use "one" for single cell anchor
      fromCol: 82,        // CE column
      fromRow: 8,         // Row 9
      toCol: 82,          // Same column (single cell)
      toRow: 8,           // Same row (single cell)
      fromColOff: 900000,  // ← Increase this to move further right (max ~950000)
      fromRowOff: 50000,
      toColOff: 0,
      toRowOff: 0,
      cx: 500000,         // ← Reduced width to fit better
      cy: 500000,         // ← Reduced height
      name: "QR Code",
    },
  ].filter(Boolean) as ImagePlacement[];


  // Inject images and trigger download
  const patchedBlob = await injectImagesIntoXlsx(xlsxBuffer, imagePlacements);
  const url = URL.createObjectURL(patchedBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "OPS_Format.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

export async function parseFormWorkbook(file: File) {
  const buffer = await file.arrayBuffer();
  const workbook = read(buffer, { type: "array" });

  if (isOPSFormat(workbook)) {
    console.log("Detected OPS Format Excel file. Parsing using custom OPS parser...");
    const parsedData = parseOPSExcel(buffer);
    const converted = convertOPSToFormQuestions(parsedData);
    return {
      title: parsedData.basicInfo.deptSection + " - Operation Standard",
      description: "Operation Standard Form imported from Excel.",
      sections: converted.sections,
      parametersToCreate: [],
      isOPS: true,
    };
  }

  const sheetNames = workbook.SheetNames;

  // Find Parameters sheet
  const parametersSheetIndex = sheetNames.findIndex((name) =>
    name.toLowerCase().includes("parameter")
  );
  const parametersSheet =
    parametersSheetIndex >= 0
      ? workbook.Sheets[sheetNames[parametersSheetIndex]]
      : null;

  // Find Form Template sheet
  const formSheetIndex = sheetNames.findIndex(
    (name) =>
      name.toLowerCase().includes("form") ||
      name.toLowerCase().includes("template")
  );
  const formSheet =
    formSheetIndex >= 0
      ? workbook.Sheets[sheetNames[formSheetIndex]]
      : workbook.Sheets[sheetNames[0]];

  if (!formSheet) {
    throw new Error("Workbook must have a Form Template sheet");
  }

  // Parse parameters from Parameters sheet
  let parametersToCreate: Array<{ name: string; type: "main" | "followup" }> =
    [];
  if (parametersSheet) {
    const parametersRawData = utils.sheet_to_json<Record<string, any>>(
      parametersSheet,
      {
        defval: "",
      }
    );

    // Extract parameters - find the parameter name column (case-insensitive)
    parametersToCreate = parametersRawData
      .filter((row) => {
        // Find the first non-empty value that isn't a header
        const firstValue = Object.values(row)[0]?.toString().trim() || "";
        return (
          firstValue &&
          firstValue.toLowerCase() !== "parameter name" &&
          firstValue !== ""
        );
      })
      .map((row) => {
        // Get the first non-empty value as parameter name
        const paramName = Object.values(row)[0]?.toString().trim() || "";
        // Get the second value as type, default to 'main'
        const typeValue =
          Object.values(row)[1]?.toString().trim().toLowerCase() || "main";

        return {
          name: paramName,
          type: (typeValue === "followup" ? "followup" : "main") as
            | "main"
            | "followup",
        };
      })
      .filter((p) => p.name && p.name.toLowerCase() !== "parameter name"); // Final validation
  }

  // Parse form data from Form Template sheet
  const rawData = utils.sheet_to_json<Record<string, any>>(formSheet, {
    defval: "",
  });

  if (rawData.length === 0) {
    throw new Error("Form Template sheet is empty");
  }

  // Find the header row (containing "Question" or "Question Type")
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(rawData.length, 10); i++) {
    const row = rawData[i];
    const values = Object.values(row).map(v => String(v).toLowerCase());
    if (values.some(v => v === "question" || v === "question type" || v === "question text")) {
      headerRowIndex = i;
      break;
    }
  }

  // If we find the header, start after it. 
  // We also skip rows that look like descriptions or are empty.
  let startIndex = headerRowIndex !== -1 ? headerRowIndex + 1 : 0;

  while (startIndex < rawData.length && startIndex < 20) {
    const row = rawData[startIndex];
    const values = Object.values(row).map(v => String(v).toLowerCase());
    const isDescription = values.some(v =>
      v.includes("the question text to ask") ||
      v.includes("which section (1, 2, 3") ||
      v.includes("name of the form") ||
      v.includes("title of the section") ||
      v.includes("overview/purpose of the form") ||
      v.includes("percentage weight") ||
      v.includes("action after section") ||
      v.includes("true/false") ||
      v.includes("comma-separated") ||
      v.includes("jump to section")
    );
    const isEmpty = values.every(v => !v || v.trim() === "" || v === "undefined" || v === "null");

    if (isDescription || isEmpty) {
      startIndex++;
    } else {
      break;
    }
  }

  const dataRows = rawData.slice(startIndex);

  if (dataRows.length === 0) {
    throw new Error(
      "No data rows found in Form Template. Please ensure your questions start after the header and description rows."
    );
  }

  const formData = parseNewTemplateFormat(dataRows, parametersToCreate);

  // Return combined data
  return {
    ...formData,
    parametersToCreate,
  };
}

function parseNewTemplateFormat(
  rows: FormRowData[],
  parametersToCreate: Array<{ name: string; type: "main" | "followup" }>
): Partial<Question> & { sections: Section[] } {
  // Helper function to find column name (case-insensitive and flexible)
  const findColumnName = (
    availableColumns: string[],
    searchPatterns: string[]
  ): string | null => {
    // First try exact match
    const exactMatch = availableColumns.find((col) =>
      searchPatterns.some((p) => col === p)
    );
    if (exactMatch) return exactMatch;

    // Try case-insensitive match
    const caseInsensitiveMatch = availableColumns.find((col) =>
      searchPatterns.some((p) => col.toLowerCase() === p.toLowerCase())
    );
    if (caseInsensitiveMatch) return caseInsensitiveMatch;

    // Try loose match (contains any search term)
    const looseMatch = availableColumns.find((col) =>
      searchPatterns.some((p) => col.toLowerCase().includes(p.toLowerCase()))
    );
    if (looseMatch) {
      console.log(
        `[Excel Import] Using approximate column match: "${looseMatch}" for "${searchPatterns.join(
          ", "
        )}"`
      );
      return looseMatch;
    }

    return null;
  };

  const sectionsMap = new Map<string, Section>();
  const availableColumns = rows.length > 0 ? Object.keys(rows[0]) : [];

  const formTitleColumn = findColumnName(availableColumns, ["Form Title", "Title", "FormTitle"]);
  const formDescColumn = findColumnName(availableColumns, ["Form Description", "Description", "FormDescription"]);
  const questionTextColumn = findColumnName(availableColumns, ["Question", "Question Text", "QuestionText"]);

  // Filter out any remaining instruction/description rows from the data rows
  const dataRows = rows.filter(row => {
    const qText = (questionTextColumn ? row[questionTextColumn] : row["Question"])?.toString().trim() || "";
    const isInstruction =
      !qText ||
      qText.toLowerCase().includes("the question text to ask") ||
      qText.toLowerCase().includes("initial qualification questions") ||
      qText.toLowerCase() === "question";
    return !isInstruction;
  });

  if (dataRows.length === 0) {
    return {
      title: "Imported Form",
      description: "",
      sections: []
    };
  }

  const formTitle = (formTitleColumn ? dataRows[0][formTitleColumn] : dataRows[0]["Form Title"])?.toString().trim() || "Imported Form";
  const formDescription = (formDescColumn ? dataRows[0][formDescColumn] : dataRows[0]["Form Description"])?.toString().trim() || "";

  let currentSectionNo: string | null = null;
  const sectionLinkMap = new Map<
    string,
    { questionId: string; option: string }
  >();
  const questionMap = new Map<string, FollowUpQuestion>();
  const sectionMergingMap = new Map<string, string>(); // Map to store section merging info

  const normalizeQuestionType = (type: string): string => {
    if (!type) return "text";

    const normalizedType = String(type)
      .toLowerCase()
      .trim()
      .replace(/\s+/g, " ") // normalize multiple spaces to single space
      .replace(/\s*\/\s*/g, ""); // remove slashes and surrounding spaces

    const typeMap: Record<string, string> = {
      // Legacy/UI type names - without spaces/slashes
      shorttext: "text",
      shortint: "text",
      multiplechoice: "radio", // This should be "radio"
      longtext: "paragraph",
      longinput: "paragraph",
      dropdown: "search-select", // Defaulting dropdown to search-select
      checkboxes: "checkbox",
      fileupload: "file",
      "file upload": "file",
      fileuploader: "file",

      // Yes/No variations
      yesnona: "yesNoNA",
      "yes/no/na": "yesNoNA",
      "yes/no/n/a": "yesNoNA",

      // Core types - pass through
      text: "text",
      radio: "radio", // Ensure radio stays radio
      paragraph: "paragraph",
      select: "search-select", // Map select to search-select for better UX
      checkbox: "checkbox",
      file: "file",

      // Add these mappings for common variations
      "multiple choice": "radio",
      "drop down": "search-select",
      "drop-down": "search-select",
      "multi choice": "radio",
      "multi-choice": "radio",

      // New types
      "searchable select": "search-select",
      "search-select": "search-select",
      "searchableselect": "search-select",
      "product nps buckets": "productNPSTGWBuckets",
      "productnpstgwbuckets": "productNPSTGWBuckets",
      "hierarchy": "productNPSTGWBuckets",

      // Chassis types
      "chassis with zone": "chassis-with-zone",
      "chassis-with-zone": "chassis-with-zone",
      "chassiswithzone": "chassis-with-zone",
      "chassis without zone": "chassis-without-zone",
      "chassis-without-zone": "chassis-without-zone",
      "chassiswithoutzone": "chassis-without-zone",
      "zone in": "zone-in",
      "zone-in": "zone-in",
      "zonein": "zone-in",
      "zone out": "zone-out",
      "zone-out": "zone-out",
      "zoneout": "zone-out",
    };

    // First try exact match after normalization
    if (typeMap[normalizedType]) {
      console.log(
        `[TYPE MAPPING] "${type}" → "${normalizedType}" → "${typeMap[normalizedType]}"`
      );
      return typeMap[normalizedType];
    }

    // If not found, try with spaces removed entirely
    const noSpaces = normalizedType.replace(/\s/g, "");
    if (typeMap[noSpaces]) {
      console.log(
        `[TYPE MAPPING] "${type}" → "${noSpaces}" → "${typeMap[noSpaces]}"`
      );
      return typeMap[noSpaces];
    }

    console.warn(
      `[TYPE MAPPING] Unknown type: "${type}", defaulting to "text"`
    );
    return "text";
  };
  function parseNestedFollowUps(
    row: FormRowData,
    parentQuestion: FollowUpQuestion,
    parentId: string,
    levelPath: string
  ) {
    console.log(
      `[NESTED] Parsing level ${levelPath} for parent: ${parentQuestion.text}`
    );

    // Find all child columns for this level
    const childPattern = new RegExp(`^FU${levelPath}\\.(\\d+):\\s*(.+)$`);
    const childColumns = Object.keys(row).filter((key) =>
      childPattern.test(key)
    );

    console.log(
      `[NESTED] Found ${childColumns.length} child columns:`,
      childColumns
    );

    if (childColumns.length === 0) return;

    // Group by child number (1, 2, 3, etc.)
    const childGroups = new Map<number, Record<string, string>>();

    childColumns.forEach((column) => {
      const match = column.match(childPattern);
      if (!match) return;

      const childNum = parseInt(match[1]);
      const columnType = match[2].trim(); // "Option", "Question Type", etc.

      if (!childGroups.has(childNum)) {
        childGroups.set(childNum, {});
      }

      const childData = childGroups.get(childNum)!;
      childData[columnType] = row[column]?.toString().trim() || "";
    });

    console.log(`[NESTED] Child groups:`, Array.from(childGroups.entries()));

    // Process each child group
    childGroups.forEach((childData, childNum) => {
      const childOption = childData["Option"];
      const childType = childData["Question Type"];
      const childText = childData["Question Text"];

      console.log(`[NESTED] Processing child ${childNum}:`, {
        childOption,
        childType,
        childText,
      });

      if (!childOption || !childType || !childText) {
        console.warn(`[NESTED] Missing required fields for child ${childNum}`);
        return;
      }

      const childRequired =
        (childData["Required"] || "FALSE").toLowerCase() === "true";
      const childSubParam1 = childData["SubParam1"] || childData["Main Parameter"] || childData["Sub Parameter 1"] || undefined;
      const childSubParam2 = childData["SubParam2"] || childData["Followup Parameter"] || childData["Sub Parameter 2"] || childData["Follow up Parameter"] || undefined;
      const childOptionsStr = childData["Options"] || "";
      const childCorrectAnswer = childData["Correct Answer"] || undefined;
      const childRankingLogicRaw = childData["Ranking Logic"] || "FALSE";
      const childRankingLogic = childRankingLogicRaw.toLowerCase() === "true" || childRankingLogicRaw === "1";
      const childTrackQuestionRaw = childData["Track Question"] || "FALSE";
      const childTrackQuestion = childTrackQuestionRaw.toLowerCase() === "true" || childTrackQuestionRaw === "1";
      const childImageUrl = childData["Image/File URL"] || undefined;
      const childDescription = childData["Description"] || undefined;

      const childOptions = childOptionsStr
        ? childOptionsStr
          .split(",")
          .map((opt) => opt.trim())
          .filter(Boolean)
        : undefined;

      const childQuestionId = generateId();
      const childQuestion: FollowUpQuestion = {
        id: childQuestionId,
        text: childText,
        type: normalizeQuestionType(childType) as FollowUpQuestion["type"],
        required: childRequired,
        options: childOptions,
        description: childDescription,
        imageUrl: childImageUrl,
        followUpQuestions: [],
        sectionId: parentQuestion.sectionId,
        correctAnswer: childCorrectAnswer,
        showWhen: {
          questionId: parentId,
          value: childOption,
        },
        subParam1: childSubParam1,
        subParam2: childSubParam2,
        trackResponseRank: childRankingLogic,
        trackResponseQuestion: childTrackQuestion,
        allowedFileTypes: undefined,
      };

      // Add to parent's follow-ups
      parentQuestion.followUpQuestions = parentQuestion.followUpQuestions || [];
      parentQuestion.followUpQuestions.push(childQuestion);

      console.log(
        `[NESTED] Added child question: ${childText} to parent: ${parentQuestion.text}`
      );

      // Recursively parse deeper levels
      const nextLevelPath = levelPath
        ? `${levelPath}.${childNum}`
        : childNum.toString();
      parseNestedFollowUps(row, childQuestion, childQuestionId, nextLevelPath);
    });
  }

  // Log available columns for debugging
  let mergingColumnName = "Section Merging";
  let nextSectionColumnName = "Next Section";
  let branchingColumnName = "Branching";
  let sectionNoColumnName = "Section Number";
  let questionTextColumnName = "Question";

  if (rows.length > 0) {
    console.log("[Excel Import] Available columns:", availableColumns);

    const foundSectionNoColumn = findColumnName(availableColumns, [
      "Section Number",
      "Section No",
      "Section #",
      "Section",
    ]);
    if (foundSectionNoColumn) {
      sectionNoColumnName = foundSectionNoColumn;
    }

    const foundQuestionTextColumn = findColumnName(availableColumns, [
      "Question",
      "Question Text",
      "QuestionText",
    ]);
    if (foundQuestionTextColumn) {
      questionTextColumnName = foundQuestionTextColumn;
    }

    const foundMergingColumn = findColumnName(availableColumns, [
      "Subsection Of",
      "Parent Section",
      "Section Merging",
      "Merge",
      "Merging",
    ]);
    if (foundMergingColumn) {
      mergingColumnName = foundMergingColumn;
      console.log(
        `[Excel Import] "Section Merging" column found: ${mergingColumnName}`
      );
    } else {
      console.warn(
        "[Excel Import] Warning: 'Section Merging' column not found. Make sure your Excel has this column for section merging to work."
      );
    }

    const foundNextSectionColumn = findColumnName(availableColumns, [
      "After Section Action",
      "Next Section",
      "NextSection",
      "NavigateTo",
      "Navigate To",
      "Next Section Action",
      "Form Ending",
      "Form Action",
    ]);
    if (foundNextSectionColumn) {
      nextSectionColumnName = foundNextSectionColumn;
      console.log(
        `[Excel Import] "Next Section" column found (as ${nextSectionColumnName})`
      );
    }

    const foundBranchingColumn = findColumnName(availableColumns, [
      "Section Routing",
      "Option Mapping",
      "Branching",
      "Routing",
      "Section Navigation",
      "Branching Rule",
      "Condition Mapping",
    ]);
    if (foundBranchingColumn) {
      branchingColumnName = foundBranchingColumn;
      console.log(
        `[Excel Import] "Branching" column found (as ${branchingColumnName})`
      );
    }
  }

  const sectionNavigationMap = new Map<string, string>();
  const rawBranchingRules: Array<{
    sectionNo: string;
    questionId: string;
    optionLabel: string;
    targetSectionNo: string;
  }> = [];

  dataRows.forEach((row: FormRowData) => {
    const sectionNo = row[sectionNoColumnName]?.toString().trim();
    const sectionTitle = row["Section Title"]?.toString().trim();
    const sectionDesc = row["Section Description"]?.toString().trim();
    const questionText = row[questionTextColumnName]?.toString().trim();

    // Secondary safety check for instructions
    if (!questionText ||
      questionText.toLowerCase().includes("the question text to ask") ||
      questionText.toLowerCase().includes("initial qualification questions") ||
      questionText.toLowerCase() === "question") {
      return;
    }

    const questionId = generateId();
    const branchingRules: any[] = [];

    if (sectionNo) {
      currentSectionNo = sectionNo;
      const sectionMerging = row[mergingColumnName]?.toString().trim() || "";
      const nextSection = row[nextSectionColumnName]?.toString().trim() || "";

      console.log(
        `[Excel Import] Section ${sectionNo}: Title="${sectionTitle}", Merging="${sectionMerging}", NextSection="${nextSection}"`
      );

      if (sectionMerging) {
        sectionMergingMap.set(sectionNo, sectionMerging);
        console.log(
          `[Excel Import] Stored merging data for section ${sectionNo}: "${sectionMerging}"`
        );
      }

      if (nextSection) {
        sectionNavigationMap.set(sectionNo, nextSection);
        console.log(
          `[Excel Import] Stored navigation data for section ${sectionNo}: "${nextSection}"`
        );
      }

      if (!sectionsMap.has(sectionNo)) {
        const newSection = {
          id: generateId(),
          title: sectionTitle || `Section ${sectionNo}`,
          description: sectionDesc || "Section description",
          questions: [],
          merging: sectionMerging || undefined,
          parentSectionId: undefined,
          isSubsection: false,
          sectionNo: sectionNo, // Added for easier identification
        };
        sectionsMap.set(sectionNo, newSection);
        console.log(
          `[Excel Import] Created new section ${sectionNo} with ID: ${newSection.id
          }, Merging: ${sectionMerging || "none"}`
        );
      } else {
        const existingSection = sectionsMap.get(sectionNo);
        if (existingSection) {
          if (sectionMerging) {
            existingSection.merging = sectionMerging;
            console.log(
              `[Excel Import] Updated merging for section ${sectionNo}: "${sectionMerging}"`
            );
          }
        }
      }
    }

    if (!currentSectionNo) {
      return;
    }

    const section = sectionsMap.get(currentSectionNo);
    if (!section) return;
    const suggestion = row["Suggestion"]?.toString().trim();
    const questionDescColumn = findColumnName(availableColumns, ["Question Description", "Description"]);
    const questionDesc = questionDescColumn ? row[questionDescColumn]?.toString().trim() : undefined;
    const questionTypeColumn = findColumnName(availableColumns, ["Question Type", "Type", "QuestionType"]);
    const questionTypeRaw = (questionTypeColumn ? row[questionTypeColumn] : row["Question Type"])?.toString().trim() || "text";
    const questionType = normalizeQuestionType(questionTypeRaw);
    const requiredColumn = findColumnName(availableColumns, ["Required", "Is Required", "Mandatory"]);
    const requiredStr = (requiredColumn ? row[requiredColumn] : row["Required"])?.toString().trim().toLowerCase();
    const required =
      requiredStr === "true" || requiredStr === "yes" || requiredStr === "1";
    const optionsColumn = findColumnName(availableColumns, ["Options", "Choices", "Values"]);
    const optionsStr = (optionsColumn ? row[optionsColumn] : row["Options"])?.toString().trim() || "";
    const correctAnswer = row["Correct Answer"]?.toString().trim();
    const correctAnswersStr = row["Correct Answers"]?.toString().trim();
    const correctAnswers = correctAnswersStr
      ? correctAnswersStr
        .split("|")
        .map((ans) => ans.trim())
        .filter(Boolean)
      : undefined;

    const options = optionsStr
      ? optionsStr
        .split(",")
        .map((opt) => opt.trim())
        .filter(Boolean)
      : undefined;

    const followUpConfig: Record<
      string,
      { hasFollowUp: boolean; required: boolean }
    > = {};

    if (options && options.length > 0) {
      options.forEach((option) => {
        followUpConfig[option] = { hasFollowUp: false, required: false };
      });
    }

    const subParam1Column = findColumnName(availableColumns, ["SubParam1", "Main Parameter", "Sub Parameter 1", "Main Parameters"]);
    const subParam2Column = findColumnName(availableColumns, ["SubParam2", "Followup Parameter", "Sub Parameter 2", "Followup Parameters", "Follow up Parameter"]);

    const subParam1 = (subParam1Column ? row[subParam1Column] : row["SubParam1"])?.toString().trim();
    const subParam2 = (subParam2Column ? row[subParam2Column] : row["SubParam2"])?.toString().trim();

    // Validate SubParam1 and SubParam2 against parameters from the Parameters sheet (if parameters exist)
    // Allow SubParam values even if no parameters are defined - they will be auto-created if needed
    if (subParam1 && parametersToCreate.length > 0) {
      const isSubParam1Valid = parametersToCreate.some(
        (p) => p.name.toLowerCase() === subParam1.toLowerCase()
      );
      if (!isSubParam1Valid) {
        console.warn(
          `SubParam1 "${subParam1}" not found in parameters. Will be treated as custom value.`
        );
      }
    }

    if (subParam2 && parametersToCreate.length > 0) {
      const isSubParam2Valid = parametersToCreate.some(
        (p) => p.name.toLowerCase() === subParam2.toLowerCase()
      );
      if (!isSubParam2Valid) {
        console.warn(
          `SubParam2 "${subParam2}" not found in parameters. Will be treated as custom value.`
        );
      }
    }

    const allowedFileTypesStr = row["Allowed File Types"]?.toString().trim();
    const allowedFileTypes = allowedFileTypesStr
      ? allowedFileTypesStr
        .split(",")
        .map((type) => type.trim().toLowerCase())
        .filter((type) => VALID_FILE_TYPES.includes(type))
      : undefined;

    // Handle the case where the user might have accidentally put a URL or other invalid text in this column
    const filteredAllowedFileTypes = (allowedFileTypes && allowedFileTypes.length > 0) ? allowedFileTypes : undefined;

    const rankingLogicRaw = row["Ranking Logic"]?.toString().trim().toLowerCase() || "false";
    const rankingLogic = rankingLogicRaw === "true" || rankingLogicRaw === "yes" || rankingLogicRaw === "1";

    const trackQuestionRaw = row["Track Question"]?.toString().trim().toLowerCase() || "false";
    const trackQuestion = trackQuestionRaw === "true" || trackQuestionRaw === "yes" || trackQuestionRaw === "1";

    const imageUrl = row["Image/File URL"]?.toString().trim() || undefined;

    // Process branching rules (Option 1 Route to Option 5 Route)
    for (let i = 1; i <= 5; i++) {
      const routeKey = `Option ${i} Route`;
      const targetSectionNo = row[routeKey]?.toString().trim();
      if (targetSectionNo && targetSectionNo !== "0" && options && options[i - 1]) {
        branchingRules.push({
          questionId: questionId,
          optionLabel: options[i - 1],
          targetSectionId: targetSectionNo,
        });

        rawBranchingRules.push({
          sectionNo: currentSectionNo!,
          questionId: questionId,
          optionLabel: options[i - 1],
          targetSectionNo: targetSectionNo
        });
      }
    }

    // Also handle combined Section Routing column
    const sectionRouting = row["Section Routing"]?.toString().trim();
    if (sectionRouting && options) {
      const routes = sectionRouting.split(",").map(r => r.trim());
      routes.forEach((targetSectionNo, idx) => {
        if (targetSectionNo && targetSectionNo !== "0" && options[idx]) {
          const alreadyAdded = branchingRules.some(r => r.optionLabel === options[idx]);
          if (!alreadyAdded) {
            branchingRules.push({
              questionId: questionId,
              optionLabel: options[idx],
              targetSectionId: targetSectionNo,
            });

            rawBranchingRules.push({
              sectionNo: currentSectionNo!,
              questionId: questionId,
              optionLabel: options[idx],
              targetSectionNo: targetSectionNo
            });
          }
        }
      });
    }

    const question: FollowUpQuestion = {
      id: questionId,
      text: questionText,
      type: questionType as FollowUpQuestion["type"],
      required: required,
      options: options || undefined,
      description: questionDesc || undefined,
      imageUrl: imageUrl,
      suggestion: suggestion || undefined,
      subParam1: subParam1 || undefined,
      subParam2: subParam2 || undefined,
      allowedFileTypes: filteredAllowedFileTypes,
      followUpQuestions: [],
      sectionId: section.id,
      correctAnswer: correctAnswer || undefined,
      correctAnswers: correctAnswers,
      trackResponseRank: rankingLogic,
      trackResponseQuestion: trackQuestion,
      ...(branchingRules.length > 0 && { branchingRules }),
    };

    // Process level 1 follow-ups (Support up to 99 main-level follow-ups)
    for (let fuIndex = 1; fuIndex <= 99; fuIndex++) {
      const fuOptionKey = `FU${fuIndex}: Option` as const;
      const fuTypeKey = `FU${fuIndex}: Question Type` as const;
      const fuRequiredKey = `FU${fuIndex}: Required` as const;
      const fuSubParam1Key = `FU${fuIndex}: SubParam1` as const;
      const fuSubParam2Key = `FU${fuIndex}: SubParam2` as const;
      const fuTextKey = `FU${fuIndex}: Question Text` as const;
      const fuDescriptionKey = `FU${fuIndex}: Description` as const;
      const fuOptionsKey = `FU${fuIndex}: Options` as const;
      const fuCorrectAnswerKey = `FU${fuIndex}: Correct Answer` as const;
      const fuRankingLogicKey = `FU${fuIndex}: Ranking Logic` as const;
      const fuTrackQuestionKey = `FU${fuIndex}: Track Question` as const;
      const fuImageUrlKey = `FU${fuIndex}: Image/File URL` as const;

      // Type-safe property access
      const fuOption = (row[fuOptionKey] as string)?.toString().trim();
      const fuTypeRaw = (row[fuTypeKey] as string)?.toString().trim();
      const fuType = fuTypeRaw ? normalizeQuestionType(fuTypeRaw) : "text";
      const fuText = (row[fuTextKey] as string)?.toString().trim();
      const fuDescription = (row[fuDescriptionKey] as string)?.toString().trim();
      const fuImageUrl = (row[fuImageUrlKey] as string)?.toString().trim();

      if (fuOption && fuType && fuText) {
        const fuRequiredStr =
          (row[fuRequiredKey] as string)?.toString().trim() || "FALSE";
        const fuRequired = fuRequiredStr.toLowerCase() === "true";
        const fuSubParam1Column = Object.keys(row).find(k => k.startsWith(`FU${fuIndex}:`) && (k.includes("SubParam1") || k.includes("Main Parameter") || k.includes("Sub Parameter 1")));
        const fuSubParam2Column = Object.keys(row).find(k => k.startsWith(`FU${fuIndex}:`) && (k.includes("SubParam2") || k.includes("Followup Parameter") || k.includes("Sub Parameter 2") || k.includes("Follow up Parameter")));

        const fuSubParam1 = (fuSubParam1Column ? row[fuSubParam1Column] : row[fuSubParam1Key])?.toString().trim();
        const fuSubParam2 = (fuSubParam2Column ? row[fuSubParam2Column] : row[fuSubParam2Key])?.toString().trim();
        const fuOptionsStr =
          (row[fuOptionsKey] as string)?.toString().trim() || "";
        const fuCorrectAnswer =
          (row[fuCorrectAnswerKey] as string)?.toString().trim() || "";
        const fuRankingLogicRaw = (row[fuRankingLogicKey] as string)?.toString().trim() || "FALSE";
        const fuRankingLogic = fuRankingLogicRaw.toLowerCase() === "true" || fuRankingLogicRaw === "1";
        const fuTrackQuestionRaw = (row[fuTrackQuestionKey] as string)?.toString().trim() || "FALSE";
        const fuTrackQuestion = fuTrackQuestionRaw.toLowerCase() === "true" || fuTrackQuestionRaw === "1";

        const fuOptions = fuOptionsStr
          ? fuOptionsStr
            .split(",")
            .map((opt) => opt.trim())
            .filter(Boolean)
          : undefined;

        const followUpId = generateId();
        const followUp: FollowUpQuestion = {
          id: followUpId,
          text: fuText,
          type: fuType as FollowUpQuestion["type"],
          required: fuRequired,
          options: fuOptions,
          description: fuDescription,
          imageUrl: fuImageUrl || undefined,
          followUpQuestions: [],
          sectionId: section.id,
          correctAnswer: fuCorrectAnswer || undefined,
          showWhen: {
            questionId: questionId,
            value: fuOption,
          },
          subParam1: fuSubParam1 || undefined,
          subParam2: fuSubParam2 || undefined,
          trackResponseRank: fuRankingLogic,
          trackResponseQuestion: fuTrackQuestion,
          allowedFileTypes: undefined,
        };

        question.followUpQuestions = question.followUpQuestions || [];
        question.followUpQuestions.push(followUp);

        // Parse nested follow-ups for this level 1 follow-up
        parseNestedFollowUps(row, followUp, followUpId, fuIndex.toString());

        if (!followUpConfig[fuOption]) {
          followUpConfig[fuOption] = {
            hasFollowUp: true,
            required: fuRequired,
          };
        } else {
          followUpConfig[fuOption].hasFollowUp = true;
          followUpConfig[fuOption].required =
            fuRequired || followUpConfig[fuOption].required;
        }
      }
    }

    if (Object.keys(followUpConfig).length > 0) {
      (question as any).followUpConfig = followUpConfig;
    }

    section.questions.push(question);
    questionMap.set(questionText, question);
  });

  const sections = Array.from(sectionsMap.values());

  // Create a mapping from section numbers to section IDs for branching
  const sectionNumberToIdMap = new Map<string, string>();
  sections.forEach((section, idx) => {
    const sectionNo = Array.from(sectionsMap.entries()).find(
      ([_, s]) => s.id === section.id
    )?.[0];
    if (sectionNo) {
      sectionNumberToIdMap.set(sectionNo, section.id);
    }
  });

  // Update section navigation to use section IDs instead of section numbers
  sections.forEach((section) => {
    // Get the section number for this section
    const sectionNo = Array.from(sectionsMap.entries()).find(
      ([_, s]) => s.id === section.id
    )?.[0];

    if (sectionNo && sectionNavigationMap.has(sectionNo)) {
      const targetSectionNo = sectionNavigationMap.get(sectionNo);
      if (targetSectionNo && targetSectionNo.toLowerCase() === "end") {
        (section as any).nextSectionId = "end";
      } else if (targetSectionNo) {
        const targetSectionId = sectionNumberToIdMap.get(targetSectionNo);
        if (targetSectionId) {
          (section as any).nextSectionId = targetSectionId;
          console.log(
            `[Navigation] Mapping section ${sectionNo} nextSection ${targetSectionNo} to ID ${targetSectionId}`
          );
        }
      }
    }
  });

  // Update branching rules to use section IDs instead of section numbers
  sections.forEach((section) => {
    section.questions.forEach((question) => {
      if (
        (question as any).branchingRules &&
        (question as any).branchingRules.length > 0
      ) {
        (question as any).branchingRules = (question as any).branchingRules.map(
          (rule: any) => {
            if (rule.targetSectionId && rule.targetSectionId.toLowerCase() === 'end') {
              return { ...rule, targetSectionId: 'end' };
            }
            const sectionId = sectionNumberToIdMap.get(rule.targetSectionId);
            if (sectionId) {
              console.log(
                `[Branching] Mapping section number ${rule.targetSectionId} to ID ${sectionId}`
              );
              return {
                ...rule,
                targetSectionId: sectionId,
              };
            }
            return rule;
          }
        );
      }
    });
  });

  sectionLinkMap.forEach((linkInfo, targetSectionNo) => {
    const targetSectionIdx = parseInt(targetSectionNo) - 1;
    if (targetSectionIdx >= 0 && targetSectionIdx < sections.length) {
      const targetSection = sections[targetSectionIdx];
      if ((targetSection as any).linkedToQuestionId === undefined) {
        (targetSection as any).linkedToQuestionId = linkInfo.questionId;
        (targetSection as any).linkedToOption = linkInfo.option;
      }
    }
  });

  // Process section merging
  // Format: "1,2" means section 1 is parent, section 2 is subsection
  console.log(
    `[Section Merging] Processing merging data. Map size: ${sectionMergingMap.size}`
  );
  console.log(
    `[Section Merging] Merging map entries:`,
    Array.from(sectionMergingMap.entries())
  );

  sectionMergingMap.forEach((mergingStr, currentSectionNo) => {
    console.log(
      `[Section Merging] Processing section ${currentSectionNo}: "${mergingStr}"`
    );

    if (!mergingStr) {
      console.log(
        `[Section Merging] Section ${currentSectionNo} has empty merging string, skipping`
      );
      return;
    }

    const sectionNumbers = mergingStr
      .split(",")
      .map((n) => n.trim())
      .filter(Boolean);

    console.log(
      `[Section Merging] Parsed section numbers: ${sectionNumbers.join(", ")}`
    );

    let parentSectionNo: string;
    let childSectionNumbers: string[];

    if (sectionNumbers.length === 1) {
      // New simplified format: "Subsection Of: 1" means current section is child of 1
      parentSectionNo = sectionNumbers[0];
      childSectionNumbers = [currentSectionNo];
    } else {
      // Legacy format: "1,2" means 1 is parent, 2 is child
      parentSectionNo = sectionNumbers[0];
      childSectionNumbers = sectionNumbers.slice(1);
    }

    const parentSectionEntry = Array.from(sectionsMap.entries()).find(
      ([sectionNo]) => sectionNo === parentSectionNo
    );
    const parentSection = parentSectionEntry?.[1];

    if (!parentSection) {
      console.warn(
        `[Section Merging] Parent section ${parentSectionNo} not found in sections map`
      );
      return;
    }

    console.log(
      `[Section Merging] Parent section found: ${parentSectionNo} (ID: ${parentSection.id
      }), Children: ${childSectionNumbers.join(", ")}`
    );

    // Set children as subsections
    childSectionNumbers.forEach((childSectionNo) => {
      const childSectionEntry = Array.from(sectionsMap.entries()).find(
        ([sectionNo]) => sectionNo.toString() === childSectionNo.toString()
      );
      const childSection = childSectionEntry?.[1];

      if (childSection && parentSection && childSection.id !== parentSection.id) {
        childSection.parentSectionId = parentSection.id;
        childSection.isSubsection = true;

        // Also ensure questions know their new section layout if needed
        childSection.questions.forEach(q => {
          q.sectionId = childSection.id;
        });

        console.log(
          `[Section Merging] ✓ Set section ${childSectionNo} (ID: ${childSection.id}) as subsection of ${parentSectionNo}`
        );
      }
    });
  });

  console.log(
    `[Section Merging] Final sections after merging:`,
    sections.map((s) => ({
      id: s.id,
      title: s.title,
      isSubsection: s.isSubsection,
      parentSectionId: s.parentSectionId,
    }))
  );

  // Process all branching rules into the top-level sectionBranching array
  const sectionBranching = rawBranchingRules.map(rule => {
    const sectionId = sectionNumberToIdMap.get(rule.sectionNo);
    let targetSectionId = rule.targetSectionNo;

    if (targetSectionId && targetSectionId.toLowerCase() !== 'end') {
      const mappedId = sectionNumberToIdMap.get(targetSectionId);
      if (mappedId) targetSectionId = mappedId;
    }

    return {
      questionId: rule.questionId,
      sectionId: sectionId,
      optionLabel: rule.optionLabel,
      targetSectionId: targetSectionId
    };
  }).filter(rule => rule.sectionId && rule.targetSectionId);

  console.log(`[Excel Import] Total branching rules created: ${sectionBranching.length}`);

  const formPayload: Partial<Question> & { sections: Section[]; sectionBranching: any[] } = {
    id: generateId(),
    title: formTitle,
    description: formDescription || "Imported form from Excel template",
    isVisible: true,
    sections,
    followUpQuestions: [],
    sectionBranching
  };

  return formPayload;
}
