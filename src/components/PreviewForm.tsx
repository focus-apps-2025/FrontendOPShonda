
import React, { useEffect, useState, useRef, useMemo, useCallback } from "react";
import { useParams, useNavigate, useSearchParams } from "react-router-dom";
import { apiClient } from "../api/client";
import SectionContent from "./preview/SectionContent";
import ThankYouMessage from "./ThankYouMessage";
import {
  MapPin, CheckCircle2, ChevronUp, Loader2,
  Sun, Moon, Database, Users, Send, Printer,
} from "lucide-react";
import { useTheme } from "../context/ThemeContext";
import { useNotification } from "../context/NotificationContext";
import { useQuestionLogic } from "../hooks/useQuestionLogic";
import { useAuth } from "../context/AuthContext";
import type { Question, Response } from "../types";
import {
  getLevel2Options, getLevel3Options, getLevel4Options,
  getLevel5Options, getLevel6Options,
} from "../config/npsHierarchy";

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────
interface Form {
  id: string; title: string; description: string; sections: any[];
  viewType?: "section-wise" | "question-wise";
  followUpQuestions?: any[];
  chassisNumbers?: Array<{ chassisNumber: string; partDescription: string }> | string[];
  chassisTenantAssignments?: Record<string, string[]>;
}

interface PreviewFormProps {
  questions?: Question[];
  onSubmit?: (response: Response) => Promise<void> | void;
  branchingRules?: any[];
  viewType?: "section-wise" | "question-wise";
  onQuestionChange?: (qId: string, qText: string, qType: string, sId: string, sTitle: string, answer?: any) => void;
  onSectionComplete?: (sId: string, sTitle: string, timeSpent: number, qCount: number) => void;
  formSessionId?: string | null;
  chassisNumbers?: Array<{ chassisNumber: string; partDescription: string }> | string[];
  chassisTenantAssignments?: Record<string, string[]>;
  opsSectionMapping?: {
    headerSectionId: string;
    generalInstructionsSectionId: string;
    pastProblemsSectionId: string;
    processStepsSectionId: string;
    associateSignSectionId: string;
    illustrationsSectionId: string;
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// OPS Template — matches screenshots pixel-for-pixel
// ─────────────────────────────────────────────────────────────────────────────
interface OPSTemplateProps {
  form: Form | null;
  answers: Record<string, any>;
  opsSectionMapping?: PreviewFormProps["opsSectionMapping"];
  onPrint: () => void;
  highlightedField: {
    questionId: string;
    columnName?: string;
    rowIndex?: number;
  } | null;
  historicalAnswers?: any[];
  isLoadingHistory?: boolean;
  submissionHistory?: Array<{ no: number; date: string; issuanceDetails: string }>;
}

function OPSTemplate({ form, answers, opsSectionMapping, onPrint, highlightedField, historicalAnswers = [], isLoadingHistory = false, submissionHistory = [] }: OPSTemplateProps) {
  const opsRef = useRef<HTMLDivElement>(null);

  // Add this function to get highlight style for a field
  const getCellHighlightStyle = (
    questionId: string | null,
    columnName?: string,
    rowIndex: number = 0
  ): React.CSSProperties => {
    if (!highlightedField || !questionId) return {};

    // Check if this is the exact field being edited
    const isExactMatch =
      (highlightedField.questionId === questionId ||
        highlightedField.questionId.includes(questionId) ||
        questionId.includes(highlightedField.questionId)) &&
      (!highlightedField.columnName || highlightedField.columnName === columnName) &&
      highlightedField.rowIndex === rowIndex;

    if (isExactMatch) {
      return {
        outline: '2px solid #3b82f6',
        outlineOffset: '1.5px',
        borderRadius: '2px',
        transition: 'all 0.2s ease-in-out',
      };
    }
    return {};
  };

  // Get highlight style for header fields (only single cell)
  const getHeaderHighlightStyle = (questionId: string | null): React.CSSProperties => {
    if (!highlightedField || !questionId) return {};

    // Only highlight if no columnName (header field) and rowIndex matches
    const isMatch =
      (highlightedField.questionId === questionId ||
        highlightedField.questionId.includes(questionId) ||
        questionId.includes(highlightedField.questionId)) &&
      (!highlightedField.columnName) &&
      highlightedField.rowIndex === 0;

    if (isMatch) {
      return {
        outline: '4px solid #29a131ff',
        outlineOffset: '1.5px',
        borderRadius: '2px',
        transition: 'all 0.2s ease-in-out',
      };
    }
    return {};
  };



  const getQuestionsFromSection = useCallback((sectionId: string | undefined) => {
    if (!form || !sectionId) return [];
    const section = form.sections?.find((s: any) => s.id === sectionId || s._id === sectionId);
    if (!section) return [];
    return (section.questions || []).filter((q: any) => !q.showWhen?.questionId);
  }, [form]);

  const headerQuestions = opsSectionMapping ? getQuestionsFromSection(opsSectionMapping.headerSectionId) : [];
  const instructionQuestions = opsSectionMapping ? getQuestionsFromSection(opsSectionMapping.headerSectionId) : [];
  const processQuestions = opsSectionMapping ? getQuestionsFromSection(opsSectionMapping.processStepsSectionId) : [];
  const pastProblemsQuestions = opsSectionMapping ? getQuestionsFromSection(opsSectionMapping.pastProblemsSectionId) : [];
  const associateQuestions = opsSectionMapping ? getQuestionsFromSection(opsSectionMapping.associateSignSectionId) : [];
  const illustrationQuestions = opsSectionMapping ? getQuestionsFromSection(opsSectionMapping.illustrationsSectionId) : [];

  const getAnswerByIndex = (questions: any[], index: number): string => {
    const question = questions[index];
    if (!question) return "";
    const questionId = question.id || question._id;
    if (!questionId) return "";
    const value = answers[questionId];
    if (value === undefined || value === null) return "";
    if (typeof value === "object") return JSON.stringify(value);
    return String(value);
  };

  const getAnswerOrFallback = (questions: any[], index: number, fallback: string): string => {
    const answer = getAnswerByIndex(questions, index);
    return answer || fallback;
  };

  const getQuestionText = (questions: any[], index: number, defaultText: string): string => {
    const question = questions[index];
    if (question?.text) return question.text;
    if (question?.label) return question.label;
    return defaultText;
  };

  const dept = getAnswerOrFallback(headerQuestions, 0, "—");
  const lineZone = getAnswerOrFallback(headerQuestions, 1, "—");
  const model = getAnswerOrFallback(headerQuestions, 2, "—");
  const station = getAnswerOrFallback(headerQuestions, 3, "—");

  const formatNo = getAnswerOrFallback(headerQuestions, 2, "—");
  const controlNo = getAnswerOrFallback(headerQuestions, 5, "—");



  const fifo = "1. Bin/trolley must be changed only after complete usage of all material in it.\n2. Empty bin/trolley should be replaced with new one.\n3. Don't top up partially filled bin.\n4. Follow FIFO on line during Process.\n5. Do not use next bin / Trolley material until running not consumed.";
  const nonLub = "Do not use any lubrication if not specified in OPS / Process Sheet.";
  const envTxt = "1. Do waste segregation.\n2. Switch off idle lights & machines.\n3. Ensure 3R Principal in daily activities.\n4. If there was any leakage, communicate to Sub Leader.";
  const safeTxt = "1. Follow POS sheet in case of any Chemical.\n2. Follow MSDS/SDS in case of any emergency regarding chemical.\n3. Follow your PPE's.";

  const pastProb = getAnswerByIndex(pastProblemsQuestions, 0);

  // Find questions by field name
  // Find questions by field name - UPDATED to handle underscore patterns
  const findQuestionByField = (fieldName: string) => {
    // Try multiple patterns to match the field
    const patterns = [
      fieldName,                           // reactionPlan
      fieldName.toLowerCase(),             // reactionplan
      fieldName.replace(/([A-Z])/g, '_$1').toLowerCase(), // reaction_plan
      fieldName.replace(/([A-Z])/g, '_$1').toLowerCase().replace(/^_/, ''), // reaction_plan (without leading underscore)
      `_${fieldName.toLowerCase()}`,       // _reactionplan
      `${fieldName.toLowerCase()}`,        // reactionplan
      `step1_${fieldName.toLowerCase()}`,  // step1_reactionplan
      `q_step1_${fieldName.toLowerCase().replace(/_/g, '_')}`, // q_step1_reaction_plan
    ];

    // Also try with the specific pattern for your IDs
    // Your IDs are like: q_step1_reaction_plan, q_step1_part_name_qty
    const specificPattern = `q_step1_${fieldName.replace(/([A-Z])/g, '_$1').toLowerCase().replace(/^_/, '')}`;
    patterns.push(specificPattern);

    const found = processQuestions.find(q => {
      const qId = (q.id || q._id || "").toLowerCase();
      return patterns.some(pattern => qId.includes(pattern.toLowerCase()));
    });

    if (found) {
      console.log(`Found question for field "${fieldName}":`, found.id);
    } else {
      console.log(`No question found for field "${fieldName}" - tried patterns:`, patterns);
    }

    return found;
  };

  const getAnswerByField = (fieldName: string): string => {
    const question = findQuestionByField(fieldName);
    if (!question) return "";
    const qId = question.id || question._id;
    const answer = answers[qId];
    if (answer === undefined || answer === null || answer === "") return "";
    if (Array.isArray(answer)) return answer.join(", ");
    if (typeof answer === "object") return JSON.stringify(answer);
    return String(answer);
  };

  const columnFields = [
    { field: 'importance', defaultLabel: 'Item Importance', width: '7%' },
    { field: 'activity', defaultLabel: 'What / Activity', width: '10%' },
    { field: 'method', defaultLabel: 'Method (How)', width: '8%' },
    { field: 'frequency', defaultLabel: 'Frequency / When', width: '6%' },
    { field: 'standard', defaultLabel: 'Standard (Spec./Criteria)', width: '10%' },
    { field: 'responsibility', defaultLabel: 'Responsibility', width: '6%' },
    { field: 'equipment', defaultLabel: 'Equipment / Measuring Eq.', width: '7%' },
    { field: 'abnormalities', defaultLabel: 'Possible Abnormalities', width: '8%' },
    { field: 'reactionPlan', defaultLabel: 'Reaction Plan', width: '6%' },
    { field: 'partNameQty', defaultLabel: 'Part Name & QTY', width: '6%' },
    { field: 'ppe', defaultLabel: 'PPEs required', width: '8%' },
    { field: 'remarks', defaultLabel: 'Remarks', width: '6%' },
  ];

  const getColumnLabel = (field: string, defaultLabel: string): string => {
    const question = findQuestionByField(field);
    if (question?.text) return question.text;
    if (question?.label) return question.label;
    return defaultLabel;
  };

  const rowAnswers = columnFields.map(col => getAnswerByField(col.field));

  const imgQ = illustrationQuestions.find((q: any) => q.id?.includes("image") || q.type === "file") || illustrationQuestions[0];
  const imgVal = imgQ ? answers[imgQ.id || imgQ._id] : null;

  // Parse image value - could be string URL or array of URLs
  let imageUrls: string[] = [];
  if (imgVal) {
    if (Array.isArray(imgVal)) {
      imageUrls = imgVal.map((item: any) => typeof item === "string" ? item : item.url).filter(Boolean);
    } else if (typeof imgVal === "string") {
      try {
        const parsed = JSON.parse(imgVal);
        if (Array.isArray(parsed)) {
          imageUrls = parsed.map((item: any) => typeof item === "string" ? item : item.url).filter(Boolean);
        } else if (parsed.url) {
          imageUrls = [parsed.url];
        } else {
          imageUrls = [imgVal];
        }
      } catch {
        imageUrls = [imgVal];
      }
    } else if (typeof imgVal === "object" && imgVal.url) {
      imageUrls = [imgVal.url];
    }
  }

  const hasImg = imageUrls.length > 0;
  const live = (v: string) => (v && v !== "—") ? "#15803d" : "#999";
  const BORDER = "1px solid #999";
  const BORDER2 = "2px solid #000";

  const C: React.CSSProperties = { border: BORDER, padding: "1px 1.5px", fontSize: "7pt", verticalAlign: "top", wordBreak: "break-word", lineHeight: "1.4" };
  const H: React.CSSProperties = { ...C, background: "#d9d9d9", fontWeight: 700, textAlign: "center", verticalAlign: "middle", fontSize: "6.5pt" };
  const L: React.CSSProperties = { ...C, background: "#e8e8e8", fontWeight: 700, fontSize: "6.5pt", verticalAlign: "middle", lineHeight: "1" };
  const V: React.CSSProperties = { ...C, background: "#fff", verticalAlign: "middle" };
  const T: React.CSSProperties = { width: "100%", borderCollapse: "collapse" as const, tableLayout: "fixed" as const, fontSize: "7pt", lineHeight: "1" };

  const LOGO = "/assets/Companylogo.png";
  const STOP = "/assets/Safetyposter.png";
  const NO_MOB = "/assets/dontusemobile.png";
  const NO_RUN = "/assets/Dontrun.png";
  const PPE_UNI = "/assets/PPEGuide.png";
  const PPE_STA = "/assets/PPEGUIDE2.png";
  const FIVE_S = "/assets/5S_Guidelines.png";
  const QR = "/assets/Qrcode.png";
  const SHIFT = "/assets/Shift_timing.png";

  const procInstructions = [
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

  const troubleRows = [
    "Equipment Trouble / Machine Break Down",
    "A Trouble You Are Responsible For",
    "Empty Marshal Carrier",
    "Stock Out / Material Shortage",
    "A Trouble From Different Section",
  ];

  const PROC_ROWS = 5;
  const ACOLS = 22;

  const LinedBox = () => (
    <table style={{ width: "100%", height: "100%", borderCollapse: "collapse" }}>
      {Array.from({ length: 10 }).map((_, i) => (
        <tr key={i}>
          <td style={{ borderBottom: i < 13 ? "0.5px solid #ccc" : "none", height: 8 }}>&nbsp;</td>
        </tr>
      ))}
    </table>
  );

  return (
    <div
      ref={opsRef}
      id="ops-template-root"
      style={{ fontFamily: "Arial, sans-serif", fontSize: "7pt", color: "#000", background: "#fff", padding: "2px" }}
    >
      {/* Retention Bar */}
      <table style={{ ...T, border: BORDER2 }}>
        <tbody>
          <tr>
            <td style={{ padding: "1px 6px", textAlign: "right", fontWeight: 700, fontSize: "6.5pt", border: BORDER2 }}>
              Retention Period : 20 years after Model is discontinued
            </td>
          </tr>
        </tbody>
      </table>

      {/* Top Header Table - COMPACT VERSION */}
      <table style={{ ...T, border: BORDER2, borderTop: "none", }}>
        <colgroup>
          <col style={{ width: "5.5%" }} />
          <col style={{ width: "3.5%" }} />
          <col style={{ width: "5%" }} />
          <col style={{ width: "3.5%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "7%" }} />
          <col style={{ width: "4%" }} />
          <col style={{ width: "11%" }} />
          <col style={{ width: "9%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "6%" }} />
          <col style={{ width: "4%" }} />
          <col style={{ width: "4%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "9%" }} />
        </colgroup>
        <tbody>
          <tr style={{ height: 4 }}> {/* Reduced from 14 */}
            <td rowSpan={8} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", padding: 0, background: "#1d4ed8" }}> {/* Reduced padding */}
              <img src={LOGO} alt="Logo" style={{ width: "100%", maxHeight: 80, objectFit: "contain" }} /> {/* Reduced from 80 */}
            </td>
            <td style={{ ...L, marginBottom: 0, lineHeight: "1.5" }}>{getQuestionText(headerQuestions, 0, "Dept. / Section")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(dept), marginBottom: 0, lineHeight: "1.5", ...getHeaderHighlightStyle(headerQuestions[0]?.id || headerQuestions[0]?._id) }}>
              {dept || "—"}
            </td>
            <td style={{ ...L, marginBottom: 0, lineHeight: "1.5" }}>{getQuestionText(headerQuestions, 1, "Line / Zones")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(lineZone), marginBottom: 0, lineHeight: "1.5", ...getHeaderHighlightStyle(headerQuestions[1]?.id || headerQuestions[1]?._id) }}>
              {lineZone || "—"}
            </td>
            <td colSpan={4} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", padding: 2 }}> {/* Reduced padding */}
              <div style={{ fontSize: "9pt", fontWeight: 700, letterSpacing: 1 }}>Operation Standard</div> {/* Reduced from 14pt */}
            </td>
            {/* Prepared, Checked, Approved — plain empty cells (no ruled lines) */}
            {[0, 1, 2].map(i => (
              <td key={i} rowSpan={7} style={{ border: BORDER2, verticalAlign: "top", padding: 0 }} />
            ))}
            {/* No., DD/MM/YY, Issuance/Revision details — keep ruled lines */}
            {[3, 4, 5].map(i => (
              <td key={i} rowSpan={7} style={{ border: BORDER2, verticalAlign: "top", padding: 0, background: "#fff" }}>
                <table style={{ width: "100%", height: "100%", borderCollapse: "collapse" }}>
                  {Array.from({ length: 12 }).map((_, rIndex) => {
                    const entry = submissionHistory?.[rIndex];
                    let displayValue: string | number = "";
                    if (entry) {
                      if (i === 3) displayValue = entry.no;
                      else if (i === 4) {
                        try {
                          displayValue = new Date(entry.date).toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "2-digit" });
                        } catch (e) {
                          displayValue = entry.date;
                        }
                      }
                      else if (i === 5) displayValue = entry.issuanceDetails;
                    }
                    return (
                      <tr key={rIndex}>
                        <td style={{
                          borderBottom: "1px solid #ccc",
                          height: 12,
                          padding: "0 2px",
                          fontSize: "5.5pt",
                          textAlign: "center",
                          color: entry ? "#15803d" : "transparent",
                          fontWeight: entry ? 700 : 400,
                          whiteSpace: "nowrap",
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                        }}>
                          {displayValue || "\u00A0"}
                        </td>
                      </tr>
                    );
                  })}
                </table>
              </td>
            ))}
            <td rowSpan={8} style={{ border: BORDER2, verticalAlign: "top", padding: "2px 3px", fontSize: "5.5pt" }}>
              {/* Format No. - Use actual question text from form */}
              <div style={{ fontWeight: 700, color: "#c00", lineHeight: "1.5", marginBottom: 2, height: 20 }}>
                {getQuestionText(headerQuestions, 4, "Format No.")} :
              </div>
              <div style={{ fontWeight: 700, fontSize: "7pt", marginBottom: 2, color: live(formatNo), ...getHeaderHighlightStyle(headerQuestions[4]?.id || headerQuestions[4]?._id) }}>
                {formatNo || "—"}
              </div>


              <div style={{ borderTop: "0.5px solid #999", margin: "2px 0" }} />

              {/* Control No. - Use actual question text from form */}
              <div style={{ fontWeight: 700, color: "#c00", lineHeight: "1.5", marginBottom: 2, height: 20 }}>
                {getQuestionText(headerQuestions, 5, "Control No.")} :
              </div>
              <div style={{ fontWeight: 700, fontSize: "7pt", marginBottom: 2, color: live(controlNo), ...getHeaderHighlightStyle(headerQuestions[5]?.id || headerQuestions[5]?._id) }}>
                {controlNo || "—"}
              </div>

              <div style={{ borderTop: "0.5px solid #999", margin: "2px 0" }} />

              {/* QR Code - This might be a static label or from form */}
              <div style={{ fontWeight: 700, marginBottom: 4, alignItems: "center", justifyContent: "center" }}>
                {getQuestionText(instructionQuestions, 6, "QR Code")} :
              </div>
              <img src={QR} alt="QR" style={{ width: 60, height: 45, objectFit: "contain" }} />
            </td>
          </tr>

          <tr style={{ height: 4 }}> {/* Reduced from 13 */}
            <td style={{ ...L, lineHeight: "1.5" }}>{getQuestionText(headerQuestions, 2, "Model")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(model), ...getHeaderHighlightStyle(headerQuestions[2]?.id || headerQuestions[2]?._id) }}>
              {model || "—"}
            </td>
            <td style={{ ...L, lineHeight: "1.5" }}>{getQuestionText(headerQuestions, 3, "Process / Station")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(station), ...getHeaderHighlightStyle(headerQuestions[3]?.id || headerQuestions[3]?._id) }}>
              {station || "—"}
            </td>
            <td colSpan={4} style={{ ...H, border: BORDER2, fontSize: "5.5pt" }}> {/* Reduced from 6.5pt */}
              Your Work When Trouble Stopped The Production Line
            </td>
          </tr>

          <tr style={{ height: 4 }}> {/* Reduced from 13 */}
            <td rowSpan={6} colSpan={2} style={{ border: BORDER2, verticalAlign: "top", fontSize: "5pt", padding: "2px 3px" }}> {/* Reduced padding */}
              <div style={{ fontWeight: 900, marginBottom: 5, fontSize: 8 }}>REJECTION HANDLING :-</div>
              <div style={{ marginBottom: 2, fontSize: 7 }}>Clearly Identify Rejected / NG parts.</div>
              <div style={{ lineHeight: "1.5", fontSize: 7 }}>Keep them properly with proper identification at defined Location.</div>
            </td>
            <td rowSpan={6} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", fontWeight: 700, fontSize: "6pt", padding: 1, lineHeight: 1.5 }}> {/* Reduced padding */}
              Measuring<br />Instruments<br />or Gauges
            </td>
            <td rowSpan={6} style={{ border: BORDER2, verticalAlign: "top", fontSize: "5pt", padding: 0 }}>
              {[
                "Always use Calibrated Measuring Instruments / Gauges.",
                "Ensure Zero setting before use.",
                "Do Not Use Unidentified Measuring Tool / Gauges.",
                "In case of any abnormality, inform Line leader and Quality Engineer.",
              ].map((txt, i, arr) => (
                <div key={i} style={{ padding: "2px 1px", borderBottom: i < arr.length - 1 ? "0.5px solid #ccc" : "none", lineHeight: "1.5" }}>{txt}</div> // Reduced padding
              ))}
            </td>
            <td rowSpan={6} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", padding: 2 }}> {/* Reduced padding */}
              <img src={STOP} alt="Stop Call Wait" style={{ maxWidth: "100%", height: 90, objectFit: "contain" }} /> {/* Reduced from 180 */}
            </td>
            <td style={{ ...H, }}>S. No.</td>
            <td style={H}>Trouble</td>
            <td style={H}>Your task</td>
          </tr>

          <tr style={{ height: 4 }}> {/* Reduced from 12 */}
            <td style={{ ...C, textAlign: "center" }}>1</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{troubleRows[0]}</td> {/* Reduced from 6pt */}
            <td rowSpan={5} style={{ fontSize: "5.5pt", textAlign: "center", verticalAlign: "middle", lineHeight: "2" }}> {/* Reduced from 5.5pt */}
              Stop The Line<br />Inform the Zone Leader<br />Write on card if mentioned in OPS
            </td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>2</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{troubleRows[1]}</td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>3</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{troubleRows[2]}</td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>4</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{troubleRows[3]}</td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>5</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{troubleRows[4]}</td>
            <td style={{ ...H, border: BORDER2, fontSize: "6pt" }}>Prepared</td>
            <td style={{ ...H, fontSize: "6pt" }}>Checked</td>
            <td style={{ ...H, fontSize: "6pt" }}>Approved</td>
            <td style={{ ...H, fontSize: "6pt" }}>No.</td>
            <td style={{ ...H, fontSize: "6pt" }}>DD/MM/YY</td>
            <td style={{ ...H, border: BORDER2, fontSize: "6pt" }}>Issuance / Revision details</td>
          </tr>
        </tbody>
      </table>

      {/* General Instructions Banner - COMPACT */}
      <table style={{ ...T, border: BORDER2, borderTop: "none", }}>
        <tbody>
          <tr>
            <td style={{ padding: "1px 6px", fontWeight: 700, fontSize: "8pt", textAlign: "center", background: "#d9d9d9" }}> {/* Reduced padding and font */}
              General Instructions
            </td>
          </tr>
        </tbody>
      </table>

      {/* General Instructions Body - COMPACT */}
      <table style={{ ...T, border: BORDER2, borderTop: "none" }}>
        <colgroup>
          <col style={{ width: "10%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "9%" }} />
          <col style={{ width: "10%" }} />
          <col style={{ width: "15%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "8%" }} />
          <col style={{ width: "13%" }} />
          <col style={{ width: "12%" }} />
        </colgroup>
        <tbody>
          <tr>
            <td style={H}>FIFO System</td>
            <td style={H}>Non Lubrication Rule</td>
            <td style={{ ...H, fontSize: "5pt" }}>Always wear PPEs / Proper uniform</td> {/* Reduced from 5.5pt */}
            <td style={{ ...H, fontSize: "5pt" }}>Wear PPEs as per<br />station requirements</td>
            <td style={H}>Shift Timings</td>
            <td style={H}>Environmental Issues</td>
            <td style={H}>Safety Issues</td>
            <td style={H}>5S Guidelines</td>
            <td style={H}>Process Instructions</td>
          </tr>
          <tr>
            <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}> {/* Reduced font */}
              <div style={{ fontWeight: 800, marginBottom: 1, fontSize: "6pt" }}>FIFO System</div> {/* Reduced margin */}
              {fifo.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, lineHeight: "2", fontSize: "6pt" }}>{l}</div>)} {/* Reduced margin */}
            </td>
            <td style={{ ...C, verticalAlign: "top", padding: 0 }}>
              <div style={{ padding: "2px 3px", borderBottom: "0.5px solid #ccc", fontSize: "6pt", lineHeight: "2" }}>{nonLub}</div> {/* Reduced padding */}
              <div style={{ display: "flex", borderBottom: "0.5px solid #ccc" }}>
                <div style={{ flex: 1, borderRight: "0.5px solid #ccc", padding: "1px 2px", textAlign: "center", fontWeight: 700, fontSize: "5pt", background: "#d9d9d9", lineHeight: "2" }}>No mobile on shopfloor</div>
                <div style={{ flex: 1, padding: "1px 2px", textAlign: "center", fontWeight: 700, fontSize: "5pt", background: "#d9d9d9", lineHeight: "2" }}>Do not run on shopfloor</div>
              </div>
              <div style={{ display: "flex" }}>
                <div style={{ flex: 1, borderRight: "0.5px solid #ccc", padding: 2, textAlign: "center" }}> {/* Reduced padding */}
                  <img src={NO_MOB} alt="No Mobile" style={{ width: 60, height: 90, objectFit: "contain" }} /> {/* Reduced from 72x72 */}
                </div>
                <div style={{ flex: 1, padding: 2, textAlign: "center" }}>
                  <img src={NO_RUN} alt="No Run" style={{ width: 60, height: 90, objectFit: "contain" }} />
                </div>
              </div>
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 2 }}> {/* Reduced padding */}
              <img src={PPE_UNI} alt="Full PPE Uniform" style={{ width: "100%", height: 140, objectFit: "contain" }} /> {/* Reduced from 190 */}
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 2 }}>
              <img src={PPE_STA} alt="Station PPE" style={{ width: "100%", maxHeight: 140, objectFit: "contain" }} /> {/* Reduced from 190 */}
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "top", padding: 2, marginTop: 6 }}> {/* Reduced padding */}
              <img src={SHIFT} alt="Shift Timings" style={{ width: "100%", maxHeight: 160, objectFit: "contain" }} /> {/* Reduced from 190 */}
            </td>
            <td style={{ ...C, verticalAlign: "top", fontSize: "5.5pt" }}>
              <div style={{ fontWeight: 700, color: "#166534", marginBottom: 2 }}>Environmental Issues</div> {/* Reduced margin */}
              {envTxt.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 1, fontSize: "6pt", lineHeight: "2" }}>{l}</div>)} {/* Reduced margin */}
            </td>
            <td style={{ ...C, verticalAlign: "top", fontSize: "5.5pt" }}>
              <div style={{ fontWeight: 700, color: "#991b1b", marginBottom: 2 }}>Safety Issues</div>
              {safeTxt.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, lineHeight: "2", fontSize: "6pt" }}>{l}</div>)}
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 2 }}>
              <img src={FIVE_S} alt="5S Guidelines" style={{ width: "100%", maxHeight: 100, objectFit: "fill" }} /> {/* Reduced from 140 */}
            </td>
            <td style={{ ...C, verticalAlign: "top", fontSize: "6pt" }}>
              {procInstructions.map((l, i) => (
                <div key={i} style={{
                  marginBottom: "0px",
                  lineHeight: "2"  // THIS is the fix for wrapped line spacing
                }}>
                  {l}
                </div>
              ))} {/* Reduced margin */}
            </td>
          </tr>
        </tbody>
      </table>

      {/* Process Steps Table */}
      <table style={{ ...T, border: BORDER2, borderTop: "none" }}>
        <colgroup>
          <col style={{ width: "8%" }} />
          <col style={{ width: "2.5%" }} />
          {columnFields.map((col) => <col key={col.field} style={{ width: col.width }} />)}
        </colgroup>
        <thead>
          <tr>
            <th style={{ ...H, background: "#ffff00", color: "#000" }}>
              {getQuestionText(illustrationQuestions, 0, "Illustrations &\nProcess Details")}
            </th>
            <th style={H}>SN</th>
            {columnFields.map((col) => (
              <th key={col.field} style={{ ...H, whiteSpace: "pre-line" }}>
                {getColumnLabel(col.field, col.defaultLabel)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {/* ── Loading indicator ── */}
          {isLoadingHistory && (
            <tr>
              <td colSpan={2 + columnFields.length} style={{ ...C, textAlign: 'center', background: '#f0f9ff', color: '#3b82f6', fontStyle: 'italic', height: 28 }}>
                ⏳ Fetching historical responses…
              </td>
            </tr>
          )}

          {/* ── Historical rows (previous responses) ── */}
          {!isLoadingHistory && historicalAnswers.map((histResp, histIdx) => {
            const histAnswers: Record<string, any> = histResp.answers || histResp;
            return (
              <tr key={`hist-${histIdx}`}>
                {/* Illustration cell – blank for historical */}
                <td style={{ ...C, background: '#f3f4f6', height: 50, minHeight: 50 }} />
                {/* SN */}
                <td style={{ ...C, background: '#f3f4f6', textAlign: 'center', fontWeight: 700, fontSize: '9pt', verticalAlign: 'middle', fontStyle: 'italic', color: '#6b7280' }}>
                  {histIdx + 1}
                </td>
                {columnFields.map((col) => {
                  // Extract this column's answer from the historical response
                  const fieldQuestion = findQuestionByField(col.field);
                  const qId = fieldQuestion?.id || fieldQuestion?._id;
                  const val = qId ? histAnswers[qId] : '';
                  const displayVal = val !== undefined && val !== null && val !== '' ? String(val) : '';
                  return (
                    <td
                      key={col.field}
                      style={{
                        ...C,
                        background: '#f3f4f6',
                        fontStyle: 'italic',
                        color: displayVal ? '#374151' : '#d1d5db',
                        height: 50,
                        minHeight: 50,
                        fontSize: '6.5pt',
                      }}
                    >
                      {displayVal || '\u00A0'}
                    </td>
                  );
                })}
              </tr>
            );
          })}



          {/* ── Current / new response row ── */}
          <tr>
            <td style={{ ...C, background: "#ffff00", textAlign: "center", verticalAlign: "middle", padding: 2, minHeight: 50, height: 50 }}>
              {hasImg && (
                <div
                  style={{
                    display: 'flex',
                    flexWrap: 'wrap',
                    gap: '4px',
                    ...getCellHighlightStyle(imgQ?.id || imgQ?._id)
                  }}
                >
                  {imageUrls.map((url, idx) => (
                    <img
                      key={idx}
                      src={url}
                      alt={`Illustration ${idx + 1}`}
                      style={{ width: '48%', height: 40, objectFit: 'contain', cursor: 'pointer' }}
                      onClick={() => window.open(url, '_blank')}
                    />
                  ))}
                </div>
              )}
            </td>
            {/* SN = after all historical rows */}
            <td style={{ ...C, textAlign: "center", fontWeight: 700, fontSize: "9pt", verticalAlign: "middle" }}>
              {historicalAnswers.length + 1}
            </td>
            {columnFields.map((col, colIdx) => {
              const cellVal = rowAnswers[colIdx];
              const fieldQuestion = findQuestionByField(col.field);
              const questionId = fieldQuestion?.id || fieldQuestion?._id;
              const shouldShowValue = !!cellVal;
              return (
                <td
                  key={col.field}
                  style={{
                    ...C,
                    fontWeight: 700,
                    color: shouldShowValue ? '#000' : 'transparent',
                    height: 50,
                    minHeight: 50,
                    ...getCellHighlightStyle(questionId, col.field, 0)
                  }}
                >
                  {shouldShowValue ? cellVal : '\u00A0'}
                </td>
              );
            })}
          </tr>

          {/* ── Extra blank rows to maintain minimum height ── */}
          {Array.from({ length: Math.max(0, PROC_ROWS - historicalAnswers.length - 1) }).map((_, i) => (
            <tr key={`blank-${i}`}>
              <td style={{ ...C, background: "#ffff00", height: 50, minHeight: 50 }} />
              <td style={{ ...C, textAlign: 'center', fontWeight: 700, fontSize: '9pt', verticalAlign: 'middle', color: '#ccc' }}>
                {historicalAnswers.length + 2 + i}
              </td>
              {columnFields.map(col => (
                <td key={col.field} style={{ ...C, height: 50, minHeight: 50 }}>&nbsp;</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      {/* Abnormality + Past Problems */}
      <table style={{ ...T, border: BORDER2, borderTop: "none" }}>
        <colgroup>
          <col style={{ width: "14%" }} />
          <col style={{ width: "86%" }} />
        </colgroup>
        <tbody>
          <tr>
            <td rowSpan={2} style={{ ...C, border: BORDER2, verticalAlign: "top", fontSize: "5.5pt", padding: "4px 5px" }}>
              <div style={{ fontWeight: 700, marginBottom: 1 }}>{getQuestionText(pastProblemsQuestions, 0, "Abnormality Handling Route")} :</div>
            </td>
            <td style={{ ...H, textAlign: "center", fontSize: "7pt", height: 26 }}>
              {getQuestionText(pastProblemsQuestions, 1, "Past Problem Details")}
            </td>
          </tr>
          <tr>
            <td style={{ ...C, minHeight: 55, height: 55, verticalAlign: "top", fontSize: "7pt", color: pastProb ? "#000" : "#bbb" }}>
              {pastProb || "—"}
            </td>
          </tr>
        </tbody>
      </table>


      {/* Associate Name & Sign */}
      <table style={{ ...T, border: BORDER2, borderTop: "none" }}>
        <colgroup>
          <col style={{ width: "5%" }} />
          {Array.from({ length: ACOLS }).map((_, i) => (
            <col key={i} style={{ width: `${(95 / ACOLS).toFixed(2)}%` }} />
          ))}
        </colgroup>
        <tbody>
          <tr>
            <td style={{ ...L, textAlign: "center", fontSize: "5.5pt", border: BORDER2, verticalAlign: "middle" }}>
              {associateQuestions[2]?.text || associateQuestions[2]?.label || "Associate Name &\nEmp. Code"}
            </td>
            {Array.from({ length: ACOLS }).map((_, i) => (
              <td key={i} style={{ border: BORDER, height: 22, textAlign: "center", fontSize: "7pt" }}>

              </td>
            ))}
          </tr>
          <tr>
            <td style={{ ...L, textAlign: "center", fontSize: "5.5pt", border: BORDER2, verticalAlign: "middle" }}>
              {associateQuestions[3]?.text || associateQuestions[3]?.label || "Sign &\nDate"}
            </td>
            {Array.from({ length: ACOLS }).map((_, i) => (
              <td key={i} style={{ border: BORDER, height: 26, textAlign: "center", fontSize: "7pt" }}>
              </td>
            ))}
          </tr>
        </tbody>
      </table>

      {/* Page Number */}
      <table style={{ ...T, border: BORDER2, borderTop: "none" }}>
        <colgroup>
          <col style={{ width: "82%" }} />
          <col style={{ width: "18%" }} />
        </colgroup>
        <tbody>
          <tr>
            <td style={{ ...C }}>&nbsp;</td>
            <td style={{ ...C, fontWeight: 700, fontSize: "8pt", textAlign: "center" }}>Page Number : XX / XX</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Main PreviewForm Component
// ─────────────────────────────────────────────────────────────────────────────
export default function PreviewForm({
  questions: propQuestions,
  onSubmit: propOnSubmit,
  branchingRules: propBranchingRules = [],
  viewType = "section-wise",
  onQuestionChange,
  onSectionComplete,
  formSessionId,
  chassisNumbers: propChassisNumbers,
  chassisTenantAssignments: propChassisTenantAssignments,
  opsSectionMapping,
}: PreviewFormProps) {
  const { id: formId } = useParams<{ id: string }>();
  const { tenant, user } = useAuth();
  const tenantSlug = tenant?.slug;
  const navigate = useNavigate();
  const [searchParams] = useSearchParams();
  const inviteId = searchParams.get("inviteId");
  const [form, setForm] = useState<Form | null>(null);
  const isMounted = useRef(true);
  const opsContainerRef = useRef<HTMLDivElement>(null);
  useEffect(() => { return () => { isMounted.current = false; }; }, []);

  const [answers, setAnswers] = useState<Record<string, any>>({});
  const [branchingRules, setBranchingRules] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [location, setLocation] = useState<{ latitude: number; longitude: number; accuracy: number } | null>(null);
  const [locationError, setLocationError] = useState<string | null>(null);
  const [locationConfirmed] = useState(true);
  const [locationDisplayName, setLocationDisplayName] = useState<string | null>(null);
  const [reverseGeocoding, setReverseGeocoding] = useState(false);
  const [currentSectionIndex, setCurrentSectionIndex] = useState(0);
  const [navigationHistory, setNavigationHistory] = useState<number[]>([0]);
  const [validationErrors, setValidationErrors] = useState<Set<string>>(new Set());
  const [parentSectionIndex, setParentSectionIndex] = useState<number | null>(null);
  const [sectionStartTime, setSectionStartTime] = useState<Date>(new Date());
  const [sectionNavigationHistory, setSectionNavigationHistory] = useState<number[]>([0]);
  const [visitedSectionIndices, setVisitedSectionIndices] = useState<Set<number>>(new Set([0]));
  const [sectionSubmitting, setSectionSubmitting] = useState(false);
  const { darkMode, toggleDarkMode } = useTheme();
  const { showSuccess, showConfirm, showError: showNotifyError } = useNotification();
  const { getOrderedVisibleQuestions } = useQuestionLogic();

  const [highlightTimeout, setHighlightTimeout] = useState<NodeJS.Timeout | null>(null);
  const [historicalAnswers, setHistoricalAnswers] = useState<any[]>([]);
  const [isLoadingHistory, setIsLoadingHistory] = useState(false);

  const [showSubmitPopup, setShowSubmitPopup] = useState(false);
  const [popupDate, setPopupDate] = useState("");
  const [popupIssuanceDetails, setPopupIssuanceDetails] = useState("");
  const [submissionHistory, setSubmissionHistory] = useState<Array<{ no: number; date: string; issuanceDetails: string }>>([]);

  const [highlightedField, setHighlightedField] = useState<{
    questionId: string;
    columnName?: string;  // For process steps columns
    rowIndex?: number;    // For which row (0 = first row)
  } | null>(null);

  const chassisNumbers = useMemo(() => {
    const raw = propChassisNumbers || form?.chassisNumbers || [];
    return raw.map((cn: any) => typeof cn === "string" ? { chassisNumber: cn, partDescription: "" } : cn);
  }, [propChassisNumbers, form?.chassisNumbers]);

  const chassisTenantAssignments = useMemo(() =>
    propChassisTenantAssignments || form?.chassisTenantAssignments || {}
    , [propChassisTenantAssignments, form?.chassisTenantAssignments]);

  const allFormQuestions = useMemo(() => {
    const flatten = (qs: any[]): any[] => {
      let all: any[] = [];
      (qs || []).forEach((q) => { all.push(q); if (q.followUpQuestions?.length) all = all.concat(flatten(q.followUpQuestions)); });
      return all;
    };
    const allQs: any[] = [];
    form?.sections?.forEach((s: any) => {
      allQs.push(...flatten(s.questions));
      s.subsections?.forEach((ss: any) => { allQs.push(...flatten(ss.questions)); });
    });
    if (form?.followUpQuestions) allQs.push(...flatten(form.followUpQuestions));
    return allQs;
  }, [form]);

  // ── Print OPS template as A3 landscape PDF ──
  const handlePrintOPS = useCallback(() => {
    const el = opsContainerRef.current;
    if (!el) return;
    const pw = window.open("", "_blank");
    if (!pw) return;
    pw.document.write(`<!DOCTYPE html><html><head>
      <style>
        @page { size: A3 landscape; margin: 3mm; }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: Arial, sans-serif; font-size: 7pt; background: #fff; color: #000; }
        .no-print { display: none !important; }
        img { max-width: 100%; }
      </style>
    </head><body>${el.innerHTML}</body></html>`);
    pw.document.close(); pw.focus();
    setTimeout(() => { pw.print(); pw.close(); }, 500);
  }, []);

  // ── Answer validation ──
  const isValidFileInput = (value: any): boolean => {
    if (!value) return false;
    if (Array.isArray(value)) return value.length > 0 && value.some(v => isValidFileInput(v)); if (typeof value === "string") {
      try { const p = JSON.parse(value); if (p?.url && p?.location) return !!p.url; } catch { }
      return value.trim().length > 0;
    }
    return false;
  };

  const isAnswerProvided = (q: any, answer: any) => {
    if (q.type === "file") return isValidFileInput(answer);
    if (q.type === "checkbox") return Array.isArray(answer) && answer.length > 0;
    if (q.type === "radio-grid" || q.type === "checkbox-grid" || q.type === "grid") {
      if (!answer || typeof answer !== "object") return false;
      const rows = q.gridOptions?.rows || q.rows || [];
      if (!rows.length) return true;
      return rows.every((row: any) => {
        const rowId = typeof row === "string" ? row : row.id || row;
        const ra = answer[rowId];
        return q.type === "checkbox-grid" ? (Array.isArray(ra) && ra.length > 0) : (ra !== undefined && ra !== null && String(ra).trim() !== "");
      });
    }
    return answer !== undefined && answer !== null && String(answer).trim() !== "";
  };

  const validateSections = (sectionsToValidate: any[]) => {
    let isValid = true;
    const newErrors = new Set<string>();
    sectionsToValidate.forEach((section) => {
      if (!section) return;
      const allQs = [...section.questions];
      const subs = form?.sections.filter((s) => s.isSubsection && s.parentSectionId === section.id) || [];
      subs.forEach((ss) => { allQs.push(...ss.questions); });
      const visible = getOrderedVisibleQuestions(allQs, answers);
      visible.forEach((q) => {
        if (!q.required) return;
        const qId = q.id || (q as any)._id;
        if (!qId) return;
        if (!isAnswerProvided(q, answers[qId])) { isValid = false; newErrors.add(qId); }
      });
    });
    setValidationErrors(newErrors);
    if (!isValid) {
      showNotifyError("Please fill in all required questions");
      setTimeout(() => { document.querySelector('[data-error="true"]')?.scrollIntoView({ behavior: "smooth", block: "center" }); }, 100);
    }
    return isValid;
  };

  // ── Flatten sections helper ──
  const flattenSections = (sections: any[]) =>
    sections.map((section: any) => {
      const allQs: any[] = [];
      const flatten = (qs: any[], parentId?: string) => {
        (qs || []).forEach((q: any) => {
          const { followUpQuestions, ...main } = q;
          if (parentId && !main.showWhen) main.showWhen = { questionId: parentId, value: main.showWhen?.value || "" };
          allQs.push(main);
          if (followUpQuestions?.length) flatten(followUpQuestions, q.id);
        });
      };
      flatten(section.questions || []);
      return { ...section, id: section.id || section._id, nextSectionId: section.nextSectionId || section._nextSectionId, questions: allQs };
    });
  useEffect(() => {
    if (form && form.sections) {
      console.log("=== FORM STRUCTURE DEBUG ===");
      console.log("Form Title:", form.title);
      console.log("Number of sections:", form.sections.length);
      form.sections.forEach((section, idx) => {
        console.log(`Section ${idx + 1}:`, {
          id: section.id || section._id,
          title: section.title,
          isSubsection: section.isSubsection,
          parentSectionId: section.parentSectionId,
          questionCount: section.questions?.length || 0
        });
        if (section.questions) {
          section.questions.forEach((q, qIdx) => {
            console.log(`  Question ${qIdx + 1}:`, {
              id: q.id || q._id,
              text: q.text || q.label,
              type: q.type
            });
          });
        }
      });
      console.log("opsSectionMapping received:", opsSectionMapping);
    }
  }, [form, opsSectionMapping]);
  const effectiveOpsMapping = useMemo(() => {
    if (!form?.sections) return undefined;

    console.log("=== Creating mapping for sections ===");
    console.log("Available sections:", form.sections.map(s => ({
      id: s.id,
      title: s.title,
      questionCount: s.questions?.length
    })));

    const mapping = {
      // Header section - contains Dept, Line, Model, Station
      headerSectionId: "sec_basic_doc_control",

      // General Instructions section - ALSO use sec_basic_doc_control since it has Format No, Control No
      // Or create a separate section for instructions
      generalInstructionsSectionId: "sec_basic_doc_control",  // ← CHANGE THIS

      // Past Problems - from sec_abnormality
      pastProblemsSectionId: "sec_abnormality",

      // Process Steps - from sec_process_steps
      processStepsSectionId: "sec_process_steps",

      // Associate Sign - you might need to add this to your form or use a fallback
      associateSignSectionId: form.sections.find(s => s.title?.includes("Associate"))?.id || "sec_abnormality",

      // Illustrations - from sec_illustrations
      illustrationsSectionId: "sec_illustrations",
    };

    console.log("Final mapping:", mapping);

    // Log what questions are found in each section
    const headerQs = form.sections.find(s => s.id === mapping.headerSectionId)?.questions || [];
    const instructionQs = form.sections.find(s => s.id === mapping.generalInstructionsSectionId)?.questions || [];
    console.log("Header questions:", headerQs.map(q => ({ id: q.id, text: q.text })));
    console.log("Instruction questions:", instructionQs.map(q => ({ id: q.id, text: q.text })));

    return mapping;
  }, [form]);

  // If no mapping could be created, show all questions in a simple debug view
  const showDebugView = !effectiveOpsMapping ||
    (!effectiveOpsMapping.headerSectionId &&
      !effectiveOpsMapping.generalInstructionsSectionId &&
      !effectiveOpsMapping.processStepsSectionId);


  // ── Fetch form ──
  useEffect(() => {
    const fetchForm = async () => {
      if (propQuestions?.length) {
        const firstQ = propQuestions[0];
        const mockForm: Form = {
          id: formId || "preview", title: "Preview Form", description: "",
          sections: [], viewType, chassisNumbers: propChassisNumbers, chassisTenantAssignments: propChassisTenantAssignments,
        };
        if ((firstQ as any).sections?.length) {
          mockForm.title = (firstQ as any).title;
          mockForm.description = (firstQ as any).description;
          mockForm.sections = (firstQ as any).sections;
        } else {
          mockForm.sections = [{ id: "default", title: (firstQ as any).title, description: (firstQ as any).description, questions: propQuestions }];
        }
        mockForm.sections = flattenSections(mockForm.sections);
        setForm(mockForm); setBranchingRules(propBranchingRules); setLoading(false);
        return;
      }
      if (!formId) return;
      try {
        const response = await apiClient.getPublicForm(formId, tenantSlug, inviteId);
        const fetchedForm = response.form;
        if (fetchedForm && (!fetchedForm.sections?.length)) {
          fetchedForm.sections = [{ id: "default", title: fetchedForm.title, description: fetchedForm.description, questions: fetchedForm.followUpQuestions || [] }];
        }
        if (fetchedForm?.sections) fetchedForm.sections = flattenSections(fetchedForm.sections);
        setForm(fetchedForm);
        try {
          const r = await apiClient.getSectionBranchingPublic(formId, tenantSlug);
          if (r?.sectionBranching) setBranchingRules(r.sectionBranching);
        } catch { }
      } catch (err: any) {
        if (err.response?.message === "ALREADY_SUBMITTED") {
          showConfirm("You have already responded. Continue?", async () => {
            try {
              const r = await apiClient.getPublicForm(formId!, tenantSlug);
              setForm(r.form);
            } catch { setError("Failed to load form"); }
          }, "Already Responded", "Yes, Continue", "Go Home", () => navigate("/forms/analytics"));
          return;
        }
        setError("Failed to load form"); console.error(err);
      } finally { setLoading(false); }
    };
    fetchForm();
  }, [formId, tenantSlug, propQuestions, propBranchingRules, viewType]);

  // ── Location ──
  useEffect(() => {
    const go = () => {
      if (!navigator.geolocation) { setLocationError("Geolocation not supported."); return; }
      navigator.geolocation.getCurrentPosition(
        (p) => setLocation({ latitude: p.coords.latitude, longitude: p.coords.longitude, accuracy: p.coords.accuracy }),
        () => setLocationError("Location access denied."),
        { enableHighAccuracy: true, timeout: 10000, maximumAge: 300000 },
      );
    };
    if ("permissions" in navigator) {
      navigator.permissions.query({ name: "geolocation" }).then((s) => { if (s.state === "denied") { setLocationError("Location denied."); return; } go(); }).catch(go);
    } else { go(); }
  }, []);

  useEffect(() => {
    if (!location || locationDisplayName || reverseGeocoding) return;
    setReverseGeocoding(true);
    fetch(`https://api.bigdatacloud.net/data/reverse-geocode-client?latitude=${location.latitude}&longitude=${location.longitude}&localityLanguage=en`)
      .then(r => r.ok ? r.json() : null)
      .then(data => { if (data) { const parts = [data.locality, data.city, data.principalSubdivision, data.countryName].filter(Boolean); if (parts.length) setLocationDisplayName(parts.join(", ")); } })
      .catch(() => { })
      .finally(() => setReverseGeocoding(false));
  }, [location]);

  // ── Section helpers ──
  const getMainSections = useCallback(() => {
    if (!form) return [];
    const base = form.sections.filter((s) => !s.isSubsection);
    const effView = form.viewType || (form as any).view_type || "section-wise";
    if (effView === "question-wise") {
      const vs: any[] = [];
      base.forEach((section, sIdx) => {
        const visible = getOrderedVisibleQuestions(section.questions, answers);
        if (!visible.length) {
          vs.push({ ...section, questions: [], isVirtual: true, originalSectionId: section.id, originalSectionIndex: sIdx, totalOriginalSections: base.length, questionIndex: 0, totalQuestionsInSection: 0 });
        } else {
          visible.forEach((q, qIdx) => vs.push({ ...section, id: `${section.id}_v${qIdx}`, description: qIdx === 0 ? section.description : "", questions: [q], isVirtual: true, originalSectionId: section.id, originalSectionIndex: sIdx, totalOriginalSections: base.length, questionIndex: qIdx, totalQuestionsInSection: visible.length }));
        }
      });
      return vs;
    }
    const map = new Map<string, any>();
    const roots: any[] = [];
    base.forEach((s) => map.set(s.id, { ...s, subsections: [] }));
    base.forEach((s) => {
      const m = map.get(s.id);
      const isSub = s.isSubsection === true || s.isSubsection === "true" || (s.parentSectionId && s.parentSectionId !== "");
      if (isSub && s.parentSectionId) {
        const parent = map.get(s.parentSectionId) || Array.from(map.values()).find((x) => x._id === s.parentSectionId);
        if (parent) parent.subsections.push(m); else roots.push(m);
      } else roots.push(m);
    });
    return roots;
  }, [form, answers, getOrderedVisibleQuestions]);

  const getLinkedSectionIds = (): Set<string> => {
    const ids = new Set<string>();
    if (!form) return ids;
    branchingRules.forEach((r) => { if (r.targetSectionId?.toLowerCase() !== "end") ids.add(r.targetSectionId); });
    form.sections.forEach((s) => {
      if (s.nextSectionId?.toLowerCase() !== "end") ids.add(s.nextSectionId);
      if (s.isSubsection || s.parentSectionId) ids.add(s.id || s._id);
    });
    return ids;
  };

  const getNextSequentialIndex = (ci: number): number => {
    if (!form) return -1;
    const ms = getMainSections();
    const linked = getLinkedSectionIds();
    let n = ci + 1;
    while (n < ms.length) { if (!linked.has(ms[n].id)) return n; n++; }
    return -1;
  };

  const getNextSectionIndex = () => {
    if (!form) return currentSectionIndex + 1;
    const ms = getMainSections();
    const cur = ms[currentSectionIndex];
    if (!cur) return currentSectionIndex + 1;
    const effView = form.viewType || (form as any).view_type || "section-wise";
    if (effView === "question-wise" && cur.isVirtual && cur.questionIndex < cur.totalQuestionsInSection - 1) return currentSectionIndex + 1;
    const groupSections = form.sections.filter((s) => s.id === cur.id || (s.isSubsection && s.parentSectionId === cur.id));
    const groupIds = groupSections.map((s) => s.id);
    const matchRule = (q: any, ans: any) => {
      if (ans === undefined || ans === null) return null;
      const qRules = branchingRules.filter((r) => r.questionId === q.id).filter((r) => groupIds.includes(r.sectionId));
      for (const rule of qRules) {
        const rl = rule.optionLabel?.toLowerCase();
        const al = Array.isArray(ans) ? ans.join(",") : String(ans);
        const all = al.toLowerCase();
        if (rule.isOtherOption) { if (!qRules.some((r) => !r.isOtherOption && r.optionLabel?.toLowerCase() === all) && all) return rule; }
        else if (Array.isArray(ans)) { if (ans.some((v) => String(v).toLowerCase() === rl)) return rule; }
        else if (all === rl) return rule;
      }
      return null;
    };
    for (const section of groupSections) {
      for (const q of getOrderedVisibleQuestions(section.questions, answers)) {
        const rule = matchRule(q, answers[q.id]);
        if (rule) {
          if (rule.targetSectionId?.toLowerCase() === "end") return ms.length;
          const ti = ms.findIndex((s) => s.id === rule.targetSectionId);
          if (ti !== -1) return ti;
        }
      }
    }
    if (cur.nextSectionId) {
      if (cur.nextSectionId.toLowerCase() === "end") return ms.length;
      const ti = ms.findIndex((s) => s.id === cur.nextSectionId);
      if (ti !== -1) return ti;
    }
    const ns = getNextSequentialIndex(currentSectionIndex);
    return ns !== -1 ? ns : ms.length;
  };

  // ── Submit ──
  const handleSubmitClick = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!form || !formId) return;
    const ms = getMainSections();
    const sectionsToValidate = ms.filter((_, idx) => Array.from(new Set(navigationHistory)).includes(idx));
    if (!validateSections(sectionsToValidate)) return;
    setShowSubmitPopup(true);
  };

  const handleConfirmSubmit = async () => {
    if (!popupDate || !popupIssuanceDetails) {
      showNotifyError("Please fill in both Date and Issuance/Revision Details.");
      return;
    }
    const newEntry = {
      no: submissionHistory.length + 1,
      date: popupDate,
      issuanceDetails: popupIssuanceDetails
    };
    const updatedHistory = [...submissionHistory, newEntry];
    setSubmissionHistory(updatedHistory);
    await performSubmission(updatedHistory);
  };

  const performSubmission = async (updatedHistory: Array<{ no: number; date: string; issuanceDetails: string }>) => {
    if (formSessionId && chassisNumbers.length > 0 && !answers["chassis_number"]) {
      showNotifyError("Please select a Chassis Number");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }
    setSubmitting(true);
    try {
      let submittedBy: string | undefined;
      if (user) submittedBy = user.firstName && user.lastName ? `${user.firstName} ${user.lastName}` : user.username || user.email;

      // Store submission history inside answers
      const answersWithHistory = {
        ...answers,
        __submissionHistory: updatedHistory
      };

      const data: any = {
        answers: answersWithHistory,
        inviteId: inviteId || null,
        submittedBy,
        submitterContact: user ? { email: user.email || "", phone: user.phone || "" } : {},
        submissionMetadata: { date: popupDate, issuanceDetails: popupIssuanceDetails }
      };

      if (location) data.location = { latitude: location.latitude, longitude: location.longitude, accuracy: location.accuracy, source: "browser", capturedAt: new Date().toISOString() };
      if (formSessionId) { data.sessionId = formSessionId; data.startedAt = sectionStartTime; data.completedAt = new Date(); }

      if (propOnSubmit) {
        await propOnSubmit({ id: "preview-response", formId: formId || "preview", answers: answersWithHistory, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() } as any);
        showSuccess("Preview submission successful!");
        setShowSubmitPopup(false);
        setPopupDate("");
        setPopupIssuanceDetails("");
        setSubmitted(true);
        return;
      }
      if (!formId) return;
      await apiClient.submitResponse(formId, tenantSlug, data);
      showSuccess("Form submitted successfully!");
      setShowSubmitPopup(false);
      setPopupDate("");
      setPopupIssuanceDetails("");
      setSubmitted(true);
    } catch {
      showNotifyError("Failed to submit form");
    } finally {
      setSubmitting(false);
    }
  };
  const handleSectionSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!form || !formId) return;
    const ms = getMainSections();
    const cur = ms[currentSectionIndex];
    if (onSectionComplete && formSessionId) {
      const secs = Math.floor((Date.now() - sectionStartTime.getTime()) / 1000);
      let qc = cur.questions?.length || 0;
      cur.subsections?.forEach((ss: any) => { qc += ss.questions?.length || 0; });
      onSectionComplete(cur.id, cur.title || "Untitled Section", secs, qc);
    }
    setSectionStartTime(new Date());
    if (currentSectionIndex === 0 && formSessionId && chassisNumbers.length > 0 && !answers["chassis_number"]) {
      showNotifyError("Please select a Chassis Number");
      window.scrollTo({ top: 0, behavior: "smooth" });
      return;
    }

    if (!validateSections([cur])) return;
    setSectionSubmitting(true);
    try {
      const data: any = {
        answers, sectionIndex: currentSectionIndex, isSectionSubmit: true, inviteId: inviteId || null,
        submittedBy: user?.firstName && user?.lastName ? `${user.firstName} ${user.lastName}` : user?.username || user?.email,
        submitterContact: { email: user?.email || "", phone: user?.phone || "" },
      };
      if (location) data.location = { latitude: location.latitude, longitude: location.longitude, accuracy: location.accuracy, source: "browser", capturedAt: new Date().toISOString() };
      if (!propOnSubmit && formId && formId !== "preview") await apiClient.submitResponse(formId, tenantSlug, data);
      else await new Promise((r) => setTimeout(r, 300));
      const next = getNextSectionIndex();
      if (next < ms.length) {
        setNavigationHistory((p) => [...p, next]); setSectionNavigationHistory((p) => [...p, next]);
        setVisitedSectionIndices((p) => new Set(p).add(next)); setCurrentSectionIndex(next);
        window.scrollTo({ top: 0, behavior: "smooth" });
      } else handleSubmitClick(e);
    } catch { showNotifyError("Failed to submit section"); } finally { setSectionSubmitting(false); }
  };
  // Get header questions for auto-sync
  const headerQuestionsIds = useMemo(() => {
    if (!form || !effectiveOpsMapping?.headerSectionId) return null;

    const headerSection = form.sections?.find(s => s.id === effectiveOpsMapping.headerSectionId);
    if (!headerSection?.questions) return null;

    const questions = headerSection.questions;
    return {
      modelId: questions[2]?.id || questions[2]?._id,      // Model at index 2
      formatNoId: questions[4]?.id || questions[4]?._id,  // Format No at index 4
    };
  }, [form, effectiveOpsMapping]);

  // ── Fetch historical responses by model number ──
  const fetchHistoricalResponses = useCallback(async (modelNumber: string) => {
    if (!formId || !headerQuestionsIds?.modelId || !modelNumber.trim()) {
      setHistoricalAnswers([]);
      return;
    }
    setIsLoadingHistory(true);
    try {
      const results = await apiClient.getResponsesByModel(formId, tenantSlug, headerQuestionsIds.modelId, modelNumber.trim());
      setHistoricalAnswers(Array.isArray(results) ? results : []);
    } catch {
      setHistoricalAnswers([]);
    } finally {
      setIsLoadingHistory(false);
    }
  }, [formId, tenantSlug, headerQuestionsIds]);

  // ── Answer change — this updates `answers` state which OPSTemplate reads directly ──
  const handleResponseChange = useCallback((questionId: string, value: any, additionalInfo?: { columnName?: string; rowIndex?: number }) => {
    // Update the current question
    setAnswers((prev) => ({ ...prev, [questionId]: value }));

    // Auto-sync: If Model field is updated, also update Format No. + fetch history
    if (headerQuestionsIds && questionId === headerQuestionsIds.modelId) {
      if (value) {
        setAnswers((prev) => ({ ...prev, [headerQuestionsIds.formatNoId]: value }));
        fetchHistoricalResponses(String(value));
      } else {
        setHistoricalAnswers([]);
      }
    }

    // Set highlighting for the current field
    setHighlightedField({
      questionId,
      columnName: additionalInfo?.columnName,
      rowIndex: additionalInfo?.rowIndex ?? 0,
    });

    // Clear previous timeout
    if (highlightTimeout) {
      clearTimeout(highlightTimeout);
    }

    // Remove highlight after 2 seconds
    const timeout = setTimeout(() => {
      setHighlightedField(null);
    }, 2000);
    setHighlightTimeout(timeout);

    // Call onQuestionChange if needed
    if (!answers[questionId] && value && onQuestionChange && getMainSections()[currentSectionIndex]) {
      const cur = getMainSections()[currentSectionIndex];
      const q = cur.questions?.find((q: any) => q.id === questionId);
      if (q) onQuestionChange(questionId, q.text || "Unknown", q.type || "unknown", cur.id, cur.title || "Untitled Section", value);
    }
  }, [answers, currentSectionIndex, onQuestionChange, getMainSections, highlightTimeout, headerQuestionsIds, fetchHistoricalResponses]);

  const handleLoadSampleData = () => {
    if (!form) return;
    const sample: Record<string, any> = { ...answers };
    allFormQuestions.forEach((q: any) => {
      const qId = q.id || q._id;
      if (!qId || (sample[qId] !== undefined && sample[qId] !== "")) return;
      switch (q.type) {
        case "text": case "paragraph": case "email": case "tel": case "url": sample[qId] = `Sample ${q.text || qId}`; break;
        case "number": sample[qId] = 42; break;
        case "radio": case "select": case "dropdown": case "yesNoNA":
          sample[qId] = q.options?.length ? (q.options[0].label || q.options[0]) : "Yes"; break;
        case "checkbox": sample[qId] = q.options?.length ? [q.options[0].label || q.options[0]] : []; break;
        case "rating": case "rating-number": case "satisfaction-rating": case "scale": sample[qId] = 5; break;
        case "slider": case "slider-feedback": sample[qId] = 50; break;
        case "date": sample[qId] = new Date().toISOString().split("T")[0]; break;
        case "time": sample[qId] = "12:00"; break;
        default: break;
      }
    });
    setAnswers(sample);
    showSuccess("Sample data loaded!");
  };

  const handlePrevSection = () => {
    if (parentSectionIndex !== null) { setCurrentSectionIndex(parentSectionIndex); setParentSectionIndex(null); window.scrollTo({ top: 0, behavior: "smooth" }); return; }
    if (sectionNavigationHistory.length > 1) {
      const nh = [...sectionNavigationHistory]; nh.pop();
      setSectionNavigationHistory(nh);
      setNavigationHistory((p) => { const n = [...p]; n.pop(); return n; });
      setCurrentSectionIndex(nh[nh.length - 1]);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const isLastSection = (() => {
    if (!form) return true;
    const ms = getMainSections();
    const cur = ms[currentSectionIndex];
    if (!cur) return true;
    const check = form.sections.filter((s) => s.id === cur.id || (s.isSubsection && s.parentSectionId === cur.id));
    for (const s of check) {
      for (const q of getOrderedVisibleQuestions(s.questions, answers)) {
        const a = answers[q.id];
        if (a !== undefined && a !== null) {
          const r = branchingRules.find((r) => r.sectionId === s.id && r.questionId === q.id && (Array.isArray(a) ? a.some((v) => v?.toString().toLowerCase() === r.optionLabel?.toLowerCase()) : a.toString().toLowerCase() === r.optionLabel?.toLowerCase()));
          if (r?.targetSectionId?.toLowerCase() === "end") return true;
        }
      }
    }
    if (cur.nextSectionId?.toLowerCase() === "end") return true;
    const ns = getNextSequentialIndex(currentSectionIndex);
    const hasB = check.some((s) => s.questions.some((q: any) => branchingRules.some((r) => r.sectionId === s.id && r.questionId === q.id && r.targetSectionId && r.targetSectionId.toLowerCase() !== "end")));
    const hasD = cur.nextSectionId && cur.nextSectionId.toLowerCase() !== "end" && form.sections.some((s) => s.id === cur.nextSectionId);
    return ns === -1 && !hasB && !hasD;
  })();

  // ── Render ──
  if (loading) {
    return (
      <div className={`min-h-screen flex flex-col items-center justify-center gap-3 ${darkMode ? "bg-slate-950 text-slate-400" : "bg-slate-50 text-slate-500"}`}>
        <div className="relative">
          <div className={`h-12 w-12 rounded-full border-4 ${darkMode ? "border-slate-800" : "border-slate-200"} border-t-blue-500 animate-spin`} />
          <div className="absolute inset-0 flex items-center justify-center"><Loader2 className="h-5 w-5 text-blue-500" /></div>
        </div>
        <p className="text-xs font-medium animate-pulse">Loading form...</p>
      </div>
    );
  }

  if (submitted) return <ThankYouMessage redirectPath={propOnSubmit ? undefined : `/dashboard`} customMessage={form?.description} />;

  if (error || !form) {
    return (
      <div className={`min-h-screen flex items-center justify-center p-4 ${darkMode ? "bg-slate-950" : "bg-slate-50"}`}>
        <div className={`max-w-sm w-full border rounded-xl p-5 text-center ${darkMode ? "bg-red-500/10 border-red-500/20" : "bg-red-50 border-red-100"}`}>
          <div className={`inline-flex items-center justify-center w-10 h-10 rounded-full mb-3 ${darkMode ? "bg-red-500/20 text-red-500" : "bg-red-100 text-red-600"}`}><MapPin className="h-5 w-5" /></div>
          <h3 className={`text-base font-semibold mb-1.5 ${darkMode ? "text-white" : "text-slate-900"}`}>Error</h3>
          <p className={`text-xs mb-5 ${darkMode ? "text-red-400/80" : "text-red-500"}`}>{error || "Form not found"}</p>
          <button onClick={() => window.location.reload()} className="px-5 py-2 bg-red-500 text-white text-sm rounded-lg font-medium hover:bg-red-600">Try Again</button>
        </div>
      </div>
    );
  }

  const mainSections = getMainSections();
  const currentSection = mainSections[currentSectionIndex];
  const effView = form.viewType || (form as any).view_type || "section-wise";
  const subsections = form.sections.filter((s) => s.isSubsection && s.parentSectionId === currentSection?.id);
  const allSectionsToDisplay = currentSection ? [currentSection, ...subsections] : [];

  const progressPct = effView === "question-wise" && currentSection?.isVirtual
    ? ((currentSection.originalSectionIndex + 1) / currentSection.totalOriginalSections) * 100
    : ((currentSectionIndex + 1) / mainSections.length) * 100;
  const progressLabel = effView === "question-wise" && currentSection?.isVirtual
    ? `${currentSection.originalSectionIndex + 1} / ${currentSection.totalOriginalSections}`
    : `${currentSectionIndex + 1} / ${mainSections.length}`;

  return (
    <div className={`h-screen flex flex-col ${darkMode ? "bg-slate-950 text-slate-200" : "bg-slate-50 text-slate-700"}`}>

      {/* ══ TOP HEADER BAR ══ */}
      <div className={`flex-none border-b ${darkMode ? "border-slate-800 bg-slate-950/95" : "border-slate-200 bg-white/95"} backdrop-blur-xl z-40 px-4 py-2`}>
        <div className="flex items-center gap-4">
          {/* Title — fixed to 30% */}
          <div className="flex flex-col min-w-0 flex-shrink-0" style={{ width: "30%" }}>
            <h1 className={`text-sm font-black tracking-tight truncate ${darkMode ? "text-white" : "text-slate-900"}`}>{form.title}</h1>
            <span className="self-start px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-500 border border-amber-500/20 text-[9px] font-black uppercase tracking-[0.2em] animate-pulse mt-0.5">
              Preview Mode
            </span>
          </div>
          {/* Progress */}
          <div className="flex items-center gap-2 flex-1">
            <div className={`flex-1 h-1 rounded-full overflow-hidden ${darkMode ? "bg-slate-800" : "bg-slate-200"}`}>
              <div className="h-full bg-blue-500 transition-all duration-700" style={{ width: `${progressPct}%` }} />
            </div>
            <span className={`text-[10px] font-black whitespace-nowrap ${darkMode ? "text-blue-400" : "text-blue-600"}`}>
              {Math.round(progressPct)}% · {progressLabel}
            </span>
          </div>
          {/* Actions */}
          <div className="flex items-center gap-2">
            <button onClick={handleLoadSampleData} title="Load Sample Data"
              className={`p-1.5 rounded-full transition-all ${darkMode ? "bg-blue-500/10 text-blue-400 hover:bg-blue-500/20" : "bg-blue-50 text-blue-600 hover:bg-blue-100"}`}>
              <Database className="h-3.5 w-3.5" />
            </button>
            <button onClick={toggleDarkMode}
              className={`p-1.5 rounded-full transition-all ${darkMode ? "bg-slate-800 text-slate-400 hover:bg-slate-700" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>
              {darkMode ? <Sun className="h-3.5 w-3.5" /> : <Moon className="h-3.5 w-3.5" />}
            </button>
          </div>
        </div>
      </div>

      {/* ══ SPLIT PANELS ══ */}
      <div className="flex-1 flex overflow-hidden">

        {/* ── LEFT: 30% — Form ── */}
        <div
          className={`flex flex-col border-r overflow-hidden ${darkMode ? "border-slate-800 bg-slate-950" : "border-slate-200 bg-slate-50"}`}
          style={{ width: "40%", minWidth: 320 }}
        >
          {/* Panel label */}
          <div className={`flex-none px-3 py-1.5 border-b flex items-center gap-2 ${darkMode ? "border-slate-800 bg-slate-900/60" : "border-slate-100 bg-white"}`}>
            <span className="w-1.5 h-1.5 rounded-full bg-blue-500" />
            <span className={`text-[9px] font-black uppercase tracking-widest ${darkMode ? "text-slate-500" : "text-slate-400"}`}>Fill in answers</span>
          </div>

          {/* Scrollable form */}
          <div className="flex-1 overflow-y-auto px-3 py-3 space-y-3">

            {/* Chassis selector */}
            {formSessionId && chassisNumbers.length > 0 && currentSectionIndex === 0 && (
              <div className={`p-3 rounded-xl border ${darkMode ? "bg-purple-500/5 border-purple-500/20" : "bg-purple-50 border-purple-100"}`}>
                <h2 className={`text-xs font-bold mb-2 flex items-center gap-1.5 ${darkMode ? "text-purple-300" : "text-purple-900"}`}>
                  <Users className="w-3.5 h-3.5" /> Select Chassis Number *
                </h2>
                <div className="flex flex-col gap-1.5">
                  {chassisNumbers.map((cn) => {
                    const visible = !user?.tenantId || !chassisTenantAssignments[cn.chassisNumber] || !chassisTenantAssignments[cn.chassisNumber].length || chassisTenantAssignments[cn.chassisNumber].includes(user.tenantId);
                    if (!visible) return null;
                    return (
                      <button key={cn.chassisNumber} type="button"
                        onClick={() => handleResponseChange("chassis_number", cn.chassisNumber)}
                        className={`p-2.5 rounded-lg text-left border-2 text-[10px] transition-all ${answers["chassis_number"] === cn.chassisNumber ? `border-purple-600 ${darkMode ? "bg-slate-800" : "bg-white"} shadow-md` : `border-transparent ${darkMode ? "bg-slate-800/60 hover:border-purple-500/50" : "bg-white/60 hover:border-purple-300"}`}`}>
                        <div className="flex items-center justify-between">
                          <div>
                            <div className={`font-bold ${answers["chassis_number"] === cn.chassisNumber ? "text-purple-600" : ""}`}>{cn.chassisNumber}</div>
                            {cn.partDescription && <div className="opacity-60 text-[9px]">{cn.partDescription}</div>}
                          </div>
                          {answers["chassis_number"] === cn.chassisNumber && <Send className="w-3 h-3 text-purple-600" />}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Section form */}
            <form id="customer-form" onSubmit={handleSubmitClick}>
              <div className={`rounded-xl border overflow-hidden ${darkMode ? "border-slate-800 bg-slate-900/40" : "border-slate-200 bg-white shadow-sm"}`}>
                {/* Section header */}
                <div className={`border-b px-4 py-3 ${darkMode ? "border-slate-800 bg-slate-900/40" : "border-slate-100 bg-slate-50/50"}`}>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="flex items-center justify-center w-5 h-5 rounded-md bg-blue-600 text-white font-black text-[9px]">{currentSectionIndex + 1}</div>
                    <span className={`text-[8px] font-black uppercase tracking-widest ${darkMode ? "text-blue-400" : "text-blue-600"}`}>
                      {effView === "question-wise" && currentSection?.isVirtual
                        ? `Q${currentSection.questionIndex + 1}/${currentSection.totalQuestionsInSection}`
                        : `${currentSectionIndex + 1} of ${mainSections.length}`}
                    </span>
                  </div>
                  {currentSection?.title && currentSection.title !== form.title && (
                    <h2 className={`text-sm font-black ${darkMode ? "text-white" : "text-slate-900"}`}>{currentSection.title}</h2>
                  )}
                  {currentSection?.description && currentSection.description !== form.description && (
                    <p className={`text-[10px] mt-0.5 ${darkMode ? "text-slate-500" : "text-slate-500"}`}>{currentSection.description}</p>
                  )}
                </div>

                {/* Questions */}
                <div className="px-4 py-4 space-y-4">
                  {allSectionsToDisplay.map((section) => (
                    <div key={section.id}>
                      {section.isSubsection && (
                        <div className={`mb-3 pb-1.5 border-b ${darkMode ? "border-emerald-500/10" : "border-emerald-100"}`}>
                          <h3 className={`text-xs font-bold ${darkMode ? "text-emerald-400" : "text-emerald-700"}`}>{section.title}</h3>
                        </div>
                      )}
                      <SectionContent
                        section={section}
                        formTitle={form.title}
                        answers={answers}
                        onAnswerChange={handleResponseChange}
                        validationErrors={validationErrors}
                        formId={formId}
                        tenantSlug={tenantSlug}
                        suggestedAnswers={null}
                        lastSuggestionSource={null}
                        onApplyFullSuggestion={() => { }}
                        fetchingSuggestionsForId={null}
                        rankMatchedAnswers={null}
                        currentRank={null}
                        onPreviousAnswersChange={() => { }}
                      />
                    </div>
                  ))}
                </div>
              </div>
            </form>
          </div>

          {/* Bottom nav */}
          <div className={`flex-none border-t px-3 py-2 ${darkMode ? "border-slate-800 bg-slate-950/90" : "border-slate-200 bg-white/90"}`}>
            <div className="flex items-center justify-between gap-2">
              <button type="button" onClick={() => navigate("/forms/analytics")}
                className={`text-[9px] font-bold uppercase tracking-wider px-2.5 py-1.5 rounded-lg transition-all ${darkMode ? "bg-slate-800/50 text-slate-500 hover:text-slate-200" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>
                ← Portal
              </button>
              <div className="flex items-center gap-1.5">
                {currentSectionIndex > 0 && (
                  <button type="button" onClick={handlePrevSection}
                    className={`inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-[9px] font-bold uppercase transition-all ${darkMode ? "bg-slate-800/50 text-slate-500 hover:text-slate-200" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>
                    <ChevronUp className="h-2.5 w-2.5" /> Prev
                  </button>
                )}
                {!isLastSection ? (
                  <button type="button" onClick={handleSectionSubmit} disabled={sectionSubmitting}
                    className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg bg-blue-600 text-white text-[9px] font-bold uppercase hover:bg-blue-700 disabled:opacity-50 transition-all">
                    {sectionSubmitting ? <><Loader2 className="h-2.5 w-2.5 animate-spin" /> Saving...</> : <><CheckCircle2 className="h-2.5 w-2.5" /> Next</>}
                  </button>
                ) : (
                  <button type="submit" form="customer-form" onClick={handleSubmitClick} disabled={submitting}
                    className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg bg-emerald-600 text-white text-[9px] font-bold uppercase hover:bg-emerald-700 disabled:opacity-50 transition-all">
                    {submitting ? <><Loader2 className="h-2.5 w-2.5 animate-spin" /> Submitting...</> : <><CheckCircle2 className="h-2.5 w-2.5" /> Submit</>}
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* ── RIGHT: 70% — OPS Live Template ── */}
        <div className="flex flex-col overflow-hidden bg-gray-100" style={{ width: "80%" }}>
          {/* Panel label + print button */}
          <div className="flex-none px-4 py-1.5 border-b border-gray-200 bg-white flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
              <span className="text-[9px] font-black uppercase tracking-widest text-slate-400">Live OPS Template</span>
              <span className="text-[8px] text-slate-300">— updates as you type</span>
            </div>
            <button onClick={handlePrintOPS}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[9px] font-bold uppercase tracking-wider transition-all">
              <Printer className="h-3 w-3" /> Save as PDFs (A3)
            </button>
          </div>



          {/* Scrollable OPS output */}
          <div className="flex-1 overflow-y-auto overflow-x-auto p-2">
            <div ref={opsContainerRef} className="bg-white shadow-sm" style={{ width: "1550px", minWidth: 860 }}>
              {showDebugView ? (
                // Debug view - show all questions and answers
                <div style={{ padding: "20px", fontFamily: "Arial", fontSize: "12px" }}>
                  <h2 style={{ fontSize: "18px", fontWeight: "bold", marginBottom: "20px" }}>Form Questions & Answers (Debug View)</h2>
                  <p style={{ marginBottom: "20px", color: "#666" }}>
                    No OPS section mapping found. Please configure opsSectionMapping prop or ensure your form has sections with matching titles.
                    Current sections: {form.sections?.map(s => s.title).join(", ") || "none"}
                  </p>

                  {form.sections?.map((section, sIdx) => (
                    <div key={sIdx} style={{ marginBottom: "30px", border: "1px solid #ddd", borderRadius: "8px", overflow: "hidden" }}>
                      <div style={{ background: "#f0f0f0", padding: "10px", fontWeight: "bold", borderBottom: "1px solid #ddd" }}>
                        Section: {section.title || `Section ${sIdx + 1}`} (ID: {section.id || section._id})
                      </div>
                      <div style={{ padding: "15px" }}>
                        {section.questions?.map((q, qIdx) => {
                          const qId = q.id || q._id;
                          const answer = answers[qId];
                          return (
                            <div key={qIdx} style={{ marginBottom: "15px", paddingBottom: "10px", borderBottom: "1px solid #eee" }}>
                              <div style={{ fontWeight: "bold", marginBottom: "5px" }}>
                                {qIdx + 1}. {q.text || q.label || `Question ${qIdx + 1}`}
                                <span style={{ fontSize: "11px", color: "#999", marginLeft: "10px" }}>({q.type})</span>
                              </div>
                              <div style={{ color: "#0066cc", marginTop: "5px" }}>
                                Answer: {answer !== undefined && answer !== null && answer !== "" ? (
                                  typeof answer === "object" ? JSON.stringify(answer) : answer
                                ) : (
                                  <span style={{ color: "#999", fontStyle: "italic" }}>Not answered yet</span>
                                )}
                              </div>
                              <div style={{ fontSize: "11px", color: "#999", marginTop: "3px" }}>
                                Question ID: {qId}
                              </div>
                            </div>
                          );
                        })}
                        {(!section.questions || section.questions.length === 0) && (
                          <div style={{ color: "#999", fontStyle: "italic" }}>No questions in this section</div>
                        )}
                      </div>
                    </div>
                  ))}

                  {/* Also show all answers collected so far */}
                  <div style={{ marginTop: "30px", border: "1px solid #ddd", borderRadius: "8px", overflow: "hidden" }}>
                    <div style={{ background: "#f0f0f0", padding: "10px", fontWeight: "bold", borderBottom: "1px solid #ddd" }}>
                      All Answers (Raw Data)
                    </div>
                    <div style={{ padding: "15px" }}>
                      <pre style={{ background: "#f9f9f9", padding: "10px", borderRadius: "4px", overflow: "auto", fontSize: "11px" }}>
                        {JSON.stringify(answers, null, 2)}
                      </pre>
                    </div>
                  </div>
                </div>
              ) : (
                <OPSTemplate
                  form={form}
                  answers={answers}
                  opsSectionMapping={effectiveOpsMapping}
                  onPrint={handlePrintOPS}
                  highlightedField={highlightedField}
                  historicalAnswers={historicalAnswers}
                  isLoadingHistory={isLoadingHistory}
                  submissionHistory={submissionHistory}
                />
              )}
            </div>
          </div>
        </div>

      </div>

      {/* ── Submit Confirmation Popup ── */}
      {showSubmitPopup && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className={`bg-white dark:bg-gray-900 rounded-xl p-6 w-full max-w-[400px] shadow-2xl border ${darkMode ? 'border-gray-700' : 'border-gray-200'} relative`}>
            <button
              onClick={() => setShowSubmitPopup(false)}
              className="absolute top-4 right-4 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 transition-colors"
            >
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 className={`text-lg font-bold mb-2 ${darkMode ? 'text-white' : 'text-gray-900'}`}>Confirm Submission</h3>
            <p className={`text-sm mb-4 ${darkMode ? 'text-gray-400' : 'text-gray-600'}`}>Are you sure you want to submit this OPS form?</p>

            <div className="space-y-4 mb-6">
              <div>
                <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>Date *</label>
                <input
                  type="date"
                  value={popupDate}
                  onChange={(e) => setPopupDate(e.target.value)}
                  className={`w-full px-3 py-2 rounded-lg border ${darkMode ? 'bg-gray-800 border-gray-700 text-white' : 'bg-white border-gray-300 text-gray-900'} focus:outline-none focus:ring-2 focus:ring-blue-500`}
                  required
                />
              </div>
              <div>
                <label className={`block text-sm font-medium mb-1 ${darkMode ? 'text-gray-300' : 'text-gray-700'}`}>Issuance / Revision Details *</label>
                <textarea
                  value={popupIssuanceDetails}
                  onChange={(e) => setPopupIssuanceDetails(e.target.value)}
                  rows={3}
                  className={`w-full px-3 py-2 rounded-lg border ${darkMode ? 'bg-gray-800 border-gray-700 text-white' : 'bg-white border-gray-300 text-gray-900'} focus:outline-none focus:ring-2 focus:ring-blue-500`}
                  required
                />
              </div>
            </div>

            <div className="flex gap-3 justify-end">
              <button
                type="button"
                onClick={() => setShowSubmitPopup(false)}
                className={`px-4 py-2 rounded-lg font-medium text-sm transition-colors ${darkMode ? 'bg-gray-800 text-gray-300 hover:bg-gray-700' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'}`}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleConfirmSubmit}
                disabled={!popupDate || !popupIssuanceDetails || submitting}
                className="px-4 py-2 rounded-lg font-medium text-sm text-white bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2 transition-colors"
              >
                {submitting ? <><Loader2 className="w-4 h-4 animate-spin" /> Submitting...</> : "Confirm Submit"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
