import React, { useState, useEffect, useMemo, useRef } from "react";
import { useParams, useNavigate, useLocation, useSearchParams } from "react-router-dom";
import { useAuth } from "../../context/AuthContext";
import CameraCapture from "../forms/CameraCapture";
import { exportResponsesToOPSExcel, setToastFunction } from "../../utils/opsExcelExporter";
import { exportResponsesToOPSPDF } from "../../utils/opsPDFgenerator";
import {
  exportDashboardToPDF,
  exportFormAnalyticsToPDF,
} from "../../utils/formanalyticsexport";
import {
  OpsFormImages,
  downloadFormImportTemplate
} from "../../utils/exportUtils";
import {
  Users,
  CheckCircle,
  Clock,
  XCircle,
  BarChart3,
  Calendar,
  FileText,
  ArrowLeft,
  TrendingUp,
  PieChart,
  Download,
  Table,
  Edit,
  Trash2,
  Eye,
  MoreHorizontal,
  Save,
  X,
  Share2,
  Mail,
  Send,
  MessageCircle,
  Info,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  Filter,
  Reply,
  Upload,
  Camera,
  Loader2,
  Maximize,
  UsersIcon
} from "lucide-react";
import { createPortal } from "react-dom";
import { Pie, Doughnut, Radar } from "react-chartjs-2";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  ArcElement,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
  RadialLinearScale,
} from "chart.js";
import { Line, Bar } from "react-chartjs-2";
import ChartDataLabels from "chartjs-plugin-datalabels";
import { apiClient } from "../../api/client";
import ResponseQuestion from "./ResponseQuestion";
import SectionAnalytics from "./SectionAnalytics";
import CascadingFilterModal from "./CascadingFilterModal";
import * as XLSX from "xlsx-js-style";
import { isImageUrl } from "../../utils/answerTemplateUtils";
import ImageLink from "../ImageLink";
import FilePreview from "../FilePreview";
import TableColumnFilter from "./TableColumnFilter";
import ShareAnalyticsModal from "./ShareAnalyticsModal";
import AutoSendModal from "../forms/AutoSendModal";
import { useTheme } from "../../context/ThemeContext";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  ArcElement,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
  RadialLinearScale,
  ChartDataLabels,
);

interface Response {
  _id?: string;
  id: string;
  questionId: string;
  answers: Record<string, any>;
  timestamp?: string;
  createdAt?: string; // MongoDB timestamp field
  parentResponseId?: string;
  assignedTo?: string;
  assignedAt?: string;
  verifiedBy?: string;
  verifiedAt?: string;
  status?: "pending" | "verified" | "rejected";
  notes?: string;
  submissionMetadata?: {
    location?: {
      city?: string;
      country?: string;
      region?: string;
      latitude?: number;
      longitude?: number;
    };
    capturedLocation?: {
      latitude?: number;
      longitude?: number;
    };
  };
  questionTimings?: Array<{ questionId: string; timeSpent: number }>;
  totalTimeSpent?: number;
  responseRanks?: Record<string, number>;
  createdBy?: string;
  isDispatched?: boolean;
  dispatchedAt?: string;
  tenantId?: string;
  submittedBy?: string;
  timeSpent?: number;
  submitterContact?: {
    email?: string;
  };
}

// Helper function to get the timestamp from response (handles both timestamp and createdAt)
const getResponseTimestamp = (response: Response): string | undefined => {
  return response.timestamp || response.createdAt;
};

interface Section {
  weightage(weightage: any): unknown;
  id: string;
  title: string;
  description?: string;
  questions: FollowUpQuestion[];
}

interface FollowUpQuestion {
  id: string;
  text: string;
  type: string;
  required?: boolean;
  options?: string[];
  description?: string;
  followUpQuestions?: FollowUpQuestion[];
  correctAnswer?: any;
  trackResponseRank?: boolean;
  trackResponseQuestion?: boolean;
}

interface Form {
  _id: string;
  id?: string;
  title: string;
  description?: string;
  createdAt?: string;
  isVisible?: boolean;
  logoUrl?: string;
  imageUrl?: string;
  sections?: Section[];
  followUpQuestions?: FollowUpQuestion[];
  parentFormId?: string;
  parentFormTitle?: string;
}

type SectionPerformanceStat = {
  id: string;
  title: string;
  yes: number;
  no: number;
  na: number;
  accepted?: number;
  rejected?: number;
  rework?: number;
  total: number;
  weightage: number;
};

type QuestionPerformanceStat = {
  id: string;
  text: string;
  sectionTitle: string;
  accepted: number;
  rejected: number;
  rework: number;
  total: number;
};

// Add this interface
export interface SectionAnalyticsData {
  sectionId: string;
  sectionTitle: string;
  description?: string;
  stats: {
    mainQuestionCount: number;
    totalFollowUpCount: number;
    answeredMainQuestions: number;
    answeredFollowUpQuestions: number;
    totalAnswered: number;
    totalResponses: number;
    completionRate: string;
    avgResponsesPerQuestion: string;
    questionsDetail: Array<{
      id: string;
      text: string;
      followUpCount: number;
      responses: number;
      followUpDetails?: Array<{
        id: string;
        text: string;
        responses: number;
      }>;
    }>;
  };
  qualityBreakdown: Array<{
    parameterName: string;
    yes: number;
    no: number;
    na: number;
    total: number;
  }>;
  overallQuality: {
    totalYes: number;
    totalNo: number;
    totalNA: number;
    totalResponses: number;
    percentages: {
      yes: string;
      no: string;
      na: string;
    };
  };
}

export interface ZoneAnalyticsData {
  inspectionStatus: {
    accepted: number;
    rework: number;
    rejected: number;
    total: number;
  };
  zoneBreakdown: Array<{
    zone: string;
    categories: Array<{
      category: string;
      count: number;
      defects: Array<{
        name: string;
        count: number;
        reworkCount: number;
        rejectedCount: number;
      }>;
    }>;
  }>;
}

// Add helper functions
const getSectionQualityBreakdown = (
  section: Section,
  responses: Response[],
): Array<{
  parameterName: string;
  yes: number;
  no: number;
  na: number;
  total: number;
}> => {
  const qualityData: Array<{
    parameterName: string;
    yes: number;
    no: number;
    na: number;
    total: number;
  }> = [];

  // Group questions by parameter/subParam1
  const parameterGroups = new Map<
    string,
    {
      parameterName: string;
      yes: number;
      no: number;
      na: number;
      total: number;
      questions: FollowUpQuestion[];
      isRealParameter: boolean;
    }
  >();

  // Process all main questions in the section
  section.questions.forEach((q: any) => {
    // Only process main questions (not follow-ups)
    if (!q.parentId && !q.showWhen?.questionId) {
      // Check if this has a real parameter name
      const hasRealParameter = !!q.subParam1 || !!q.parameter;

      // Get parameter name (prefer subParam1 or parameter over question text)
      const paramName =
        q.subParam1 ||
        q.parameter ||
        (hasRealParameter
          ? null
          : q.text?.substring(0, 30) + (q.text?.length > 30 ? "..." : "")) ||
        null;

      // Skip if no parameter name can be extracted
      if (!paramName) return;

      if (!parameterGroups.has(paramName)) {
        parameterGroups.set(paramName, {
          parameterName: paramName,
          yes: 0,
          no: 0,
          na: 0,
          total: 0,
          questions: [],
          isRealParameter: hasRealParameter,
        });
      }

      const group = parameterGroups.get(paramName)!;
      group.questions.push(q);

      // Count responses for this question
      responses.forEach((response) => {
        const answer = response.answers?.[q.id];
        if (answer !== null && answer !== undefined && answer !== "") {
          group.total++;

          // Check if it's an inspection status object (like from ChassisWithZone)
          if (typeof answer === "object" && answer.status) {
            const status = String(answer.status).toLowerCase().trim();
            if (
              status === "accepted" ||
              status === "rework completed" ||
              status === "verified"
            ) {
              group.yes++;
            } else if (status === "rejected") {
              group.no++;
            } else if (
              status === "rework" ||
              status === "reworked" ||
              status.includes("re-rework")
            ) {
              group.na++;
            }
            return;
          }

          const answerStr = String(answer).toLowerCase().trim();
          if (
            answerStr === "accepted" ||
            answerStr === "rework completed" ||
            answerStr === "verified"
          ) {
            group.yes++;
          } else if (answerStr === "rejected") {
            group.no++;
          } else if (
            answerStr === "rework" ||
            answerStr === "reworked" ||
            answerStr.includes("re-rework")
          ) {
            group.na++;
          } else if (answerStr.includes("yes") || answerStr === "y") {
            group.yes++;
          } else if (answerStr.includes("no") || answerStr === "n") {
            group.no++;
          } else if (
            answerStr.includes("na") ||
            answerStr.includes("n/a") ||
            answerStr.includes("not applicable")
          ) {
            group.na++;
          }
        }
      });
    }
  });

  // Convert map to array and filter out groups that don't have real parameters
  parameterGroups.forEach((group) => {
    if (group.total > 0 && group.isRealParameter) {
      qualityData.push({
        parameterName: group.parameterName,
        yes: group.yes,
        no: group.no,
        na: group.na,
        total: group.total,
      });
    }
  });

  return qualityData;
};

// Add calculateOverallQuality function
const calculateOverallQuality = (qualityBreakdown: any[]) => {
  let totalYes = 0;
  let totalNo = 0;
  let totalNA = 0;
  let totalResponses = 0;

  qualityBreakdown.forEach((item) => {
    totalYes += item.yes;
    totalNo += item.no;
    totalNA += item.na;
    totalResponses += item.total;
  });

  const total = totalYes + totalNo + totalNA;

  return {
    totalYes,
    totalNo,
    totalNA,
    totalResponses,
    percentages: {
      yes: total > 0 ? ((totalYes / total) * 100).toFixed(1) : "0.0",
      no: total > 0 ? ((totalNo / total) * 100).toFixed(1) : "0.0",
      na: total > 0 ? ((totalNA / total) * 100).toFixed(1) : "0.0",
    },
  };
};

const getZoneAnalytics = (responses: Response[]): ZoneAnalyticsData => {
  const stats: ZoneAnalyticsData = {
    inspectionStatus: { accepted: 0, rework: 0, rejected: 0, total: 0 },
    zoneBreakdown: [],
  };

  const zoneMap = new Map<string, Map<string, { count: number; defects: Map<string, { total: number; rework: number; rejected: number }> }>>();

  responses.forEach((response) => {
    if (!response.answers) return;

    Object.values(response.answers).forEach((answer) => {
      if (typeof answer === "object" && answer !== null && (answer.chassisNumber || answer.status || answer.zonesData || answer.zones)) {
        // This looks like a ChassisWithZone or ZoneIn/Out answer
        const statusVal = (answer.status || "").toLowerCase().trim();
        const isAccepted = statusVal === "accepted" || statusVal === "rework completed" || statusVal === "verified";
        const isRework = statusVal === "rework" || statusVal === "reworked" || statusVal.includes("re-rework");
        const isRejected = statusVal === "rejected";

        if (isAccepted) stats.inspectionStatus.accepted++;
        else if (isRework) stats.inspectionStatus.rework++;
        else if (isRejected) stats.inspectionStatus.rejected++;

        if (statusVal) stats.inspectionStatus.total++;

        // Handle hierarchical zonesData (ChassisWithZone)
        if (answer.zonesData) {
          Object.entries(answer.zonesData).forEach(([zoneName, zoneContent]: [string, any]) => {
            if (!zoneMap.has(zoneName)) {
              zoneMap.set(zoneName, new Map());
            }
            const catMap = zoneMap.get(zoneName)!;

            if (zoneContent.categories && Array.isArray(zoneContent.categories)) {
              zoneContent.categories.forEach((cat: any) => {
                if (!catMap.has(cat.name)) {
                  catMap.set(cat.name, { count: 0, defects: new Map() });
                }
                const catData = catMap.get(cat.name)!;
                catData.count++;

                if (cat.defects && Array.isArray(cat.defects)) {
                  cat.defects.forEach((defect: any) => {
                    const defectStats = catData.defects.get(defect.name) || { total: 0, rework: 0, rejected: 0 };
                    defectStats.total++;
                    if (isRework) defectStats.rework++;
                    else if (isRejected) defectStats.rejected++;
                    catData.defects.set(defect.name, defectStats);
                  });
                }
              });
            }
          });
        }

        // Handle simple zones array (ZoneIn/ZoneOut)
        if (answer.zones && Array.isArray(answer.zones)) {
          answer.zones.forEach((zoneName: string) => {
            if (!zoneMap.has(zoneName)) {
              zoneMap.set(zoneName, new Map());
            }
            const catMap = zoneMap.get(zoneName)!;

            // For simple zones, we might use a generic "Defect" category if it's a rework/rejected
            if (isRework || isRejected) {
              const catName = "Unspecified Defects";
              if (!catMap.has(catName)) {
                catMap.set(catName, { count: 0, defects: new Map() });
              }
              const catData = catMap.get(catName)!;
              catData.count++;

              const defectName = answer.remark || "Generic Defect";
              const defectStats = catData.defects.get(defectName) || { total: 0, rework: 0, rejected: 0 };
              defectStats.total++;
              if (isRework) defectStats.rework++;
              else if (isRejected) defectStats.rejected++;
              catData.defects.set(defectName, defectStats);
            }
          });
        }
      }
    });
  });

  // Convert map to array structure
  zoneMap.forEach((catMap, zoneName) => {
    const categories: any[] = [];
    catMap.forEach((catData, catName) => {
      const defects: any[] = [];
      catData.defects.forEach((dStats, defectName) => {
        defects.push({ name: defectName, count: dStats.total, reworkCount: dStats.rework, rejectedCount: dStats.rejected });
      });
      categories.push({ category: catName, count: catData.count, defects });
    });
    stats.zoneBreakdown.push({ zone: zoneName, categories });
  });

  return stats;
};

// Add getSectionStats function
const getSectionStats = (section: Section, responses: Response[]) => {
  // Filter for main questions only (not follow-ups)
  const mainQuestionsOnly = section.questions.filter(
    (q: any) => !q.parentId && !q.showWhen?.questionId,
  );

  console.log("Main questions found:", mainQuestionsOnly.length);
  console.log("All questions in section:", section.questions.length);

  const mainQuestionCount = mainQuestionsOnly.length;
  let totalFollowUpCount = 0;
  let answeredMainQuestions = 0;
  let answeredFollowUpQuestions = 0;
  let mainQuestionResponses = 0;
  let followUpResponses = 0;

  // Count follow-up questions
  const followUpQuestionsInSection = section.questions.filter(
    (q: any) => q.parentId || q.showWhen?.questionId,
  );
  totalFollowUpCount = followUpQuestionsInSection.length;

  // Process follow-up questions
  followUpQuestionsInSection.forEach((followUp: any) => {
    const followUpResponders = responses.filter(
      (r) => r.answers && r.answers[followUp.id],
    ).length;
    if (followUpResponders > 0) {
      answeredFollowUpQuestions++;
      followUpResponses += followUpResponders;
    }
  });

  // Process main questions
  const questionsDetail = mainQuestionsOnly.map((q: any) => {
    const mainQuestionResponders = responses.filter(
      (r) => r.answers && r.answers[q.id],
    ).length;

    if (mainQuestionResponders > 0) {
      answeredMainQuestions++;
      mainQuestionResponses += mainQuestionResponders;
    }

    const relatedFollowUps = section.questions.filter(
      (fq: any) => fq.parentId === q.id || fq.showWhen?.questionId === q.id,
    );

    return {
      id: q.id,
      text: q.text || "Unnamed Question",
      followUpCount: relatedFollowUps.length,
      responses: mainQuestionResponders,
      followUpDetails: relatedFollowUps.map((fq: any) => ({
        id: fq.id,
        text: fq.text || "Unnamed Follow-up",
        responses: responses.filter((r) => r.answers && r.answers[fq.id])
          .length,
      })),
    };
  });

  const totalAnswered = answeredMainQuestions + answeredFollowUpQuestions;
  const totalQuestions = mainQuestionCount + totalFollowUpCount;
  const totalResponses = mainQuestionResponses + followUpResponses;

  const completionRate =
    totalQuestions > 0
      ? ((totalAnswered / totalQuestions) * 100).toFixed(1)
      : "0.0";

  const avgResponsesPerQuestion =
    totalQuestions > 0 ? (totalResponses / totalQuestions).toFixed(1) : "0.0";

  console.log("Processed questionsDetail:", questionsDetail);

  return {
    mainQuestionCount,
    totalFollowUpCount,
    answeredMainQuestions,
    answeredFollowUpQuestions,
    totalAnswered,
    totalResponses,
    completionRate,
    avgResponsesPerQuestion,
    questionsDetail, // Make sure this is returned
  };
};


const formatSectionLabel = (label: string, maxLength = 20): string => {
  if (!label) {
    return "";
  }
  const parts = label.match(/[A-Za-z0-9]+/g) || [];
  if (!parts.length) {
    return "";
  }
  const camel = parts
    .map((part, index) => {
      const lower = part.toLowerCase();
      if (index === 0) {
        return lower;
      }
      return lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join("");
  if (!camel) {
    return "";
  }
  const formatted = camel.charAt(0).toUpperCase() + camel.slice(1);
  return formatted.length > maxLength
    ? `${formatted.slice(0, maxLength - 3)}...`
    : formatted;
};

const extractYesNoValues = (value: any): string[] => {
  if (value === null || value === undefined) {
    return [];
  }
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return normalized ? [normalized] : [];
  }
  if (typeof value === "boolean") {
    return [value ? "yes" : "no"];
  }
  if (Array.isArray(value)) {
    return value.flatMap((item) => extractYesNoValues(item));
  }
  if (typeof value === "object") {
    return Object.values(value).flatMap((item) => extractYesNoValues(item));
  }
  return [];
};

const recognizedYesNoValues = [
  "yes",
  "no",
  "n/a",
  "na",
  "not applicable",
  "accepted",
  "rejected",
  "rework",
  "reworked",
  "verified",
  "rework completed",
];

const getRankStyle = (answer: any, darkMode: boolean = false) => {
  if (answer === null || answer === undefined) return "";
  // Ensure we stringify object/array answers for consistent hashing
  const str =
    typeof answer === "object"
      ? JSON.stringify(answer)
      : String(answer).trim().toLowerCase();
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  const colors = [
    {
      l: "bg-blue-50 text-blue-700 border-blue-200",
      d: "bg-blue-900/30 text-blue-300 border-blue-800",
    },
    {
      l: "bg-emerald-50 text-emerald-700 border-emerald-200",
      d: "bg-emerald-900/30 text-emerald-300 border-emerald-800",
    },
    {
      l: "bg-amber-50 text-amber-700 border-amber-200",
      d: "bg-amber-900/30 text-amber-300 border-amber-800",
    },
    {
      l: "bg-orange-50 text-orange-700 border-orange-200",
      d: "bg-orange-900/30 text-orange-300 border-orange-800",
    },
    {
      l: "bg-rose-50 text-rose-700 border-rose-200",
      d: "bg-rose-900/30 text-rose-300 border-rose-800",
    },
    {
      l: "bg-purple-50 text-purple-700 border-purple-200",
      d: "bg-purple-900/30 text-purple-300 border-purple-800",
    },
    {
      l: "bg-pink-50 text-pink-700 border-pink-200",
      d: "bg-pink-900/30 text-pink-300 border-pink-800",
    },
    {
      l: "bg-indigo-50 text-indigo-700 border-indigo-200",
      d: "bg-indigo-900/30 text-indigo-300 border-indigo-800",
    },
    {
      l: "bg-teal-50 text-teal-700 border-teal-200",
      d: "bg-teal-900/30 text-teal-300 border-teal-800",
    },
    {
      l: "bg-cyan-50 text-cyan-700 border-cyan-200",
      d: "bg-cyan-900/30 text-cyan-300 border-cyan-800",
    },
  ];
  const color = colors[Math.abs(hash) % colors.length];
  return darkMode ? color.d : color.l;
};

const computeSectionPerformanceStats = (
  form: Form | null,
  responses: Response[],
): SectionPerformanceStat[] => {
  if (!form?.sections || !responses.length) {
    return [];
  }

  const stats =
    form.sections?.map((section) => {
      const counts = {
        yes: 0,
        no: 0,
        na: 0,
        accepted: 0,
        rejected: 0,
        rework: 0,
        total: 0,
      };

      const processQuestion = (question: any) => {
        if (!question) {
          return;
        }

        // Process based on response answers
        responses.forEach((response) => {
          const answer = response.answers?.[question.id];
          if (answer === null || answer === undefined || answer === "") {
            return;
          }

          // Check if it's an inspection status object (like from ChassisWithZone)
          if (typeof answer === "object" && answer.status) {
            const status = String(answer.status).toLowerCase().trim();
            if (
              status === "accepted" ||
              status === "rework completed" ||
              status === "verified"
            ) {
              counts.accepted += 1;
              counts.total += 1;
            } else if (status === "rejected") {
              counts.rejected += 1;
              counts.total += 1;
            } else if (
              status === "rework" ||
              status === "reworked" ||
              status.includes("re-rework")
            ) {
              counts.rework += 1;
              counts.total += 1;
            }
          } else {
            // Handle string answers that might be inspection statuses even for non-yesNoNA types
            const answerStr = String(answer).toLowerCase().trim();
            const normalizedValues = extractYesNoValues(answer);

            if (
              answerStr === "accepted" ||
              answerStr === "rework completed" ||
              answerStr === "verified"
            ) {
              counts.accepted += 1;
              counts.total += 1;
            } else if (answerStr === "rejected") {
              counts.rejected += 1;
              counts.total += 1;
            } else if (
              answerStr === "rework" ||
              answerStr === "reworked" ||
              answerStr.includes("re-rework")
            ) {
              counts.rework += 1;
              counts.total += 1;
            } else if (
              question.type === "yesNoNA" ||
              question.type === "chassisWithZone" ||
              question.type === "chassisWithoutZone" ||
              question.type === "chassis" ||
              question.type === "zone-in" ||
              question.type === "zone-out" ||
              question.text?.toLowerCase().includes("chassis")
            ) {
              const options = question.options || [];

              if (options.length >= 3) {
                const yesOption = String(options[0]).toLowerCase().trim();
                const noOption = String(options[1]).toLowerCase().trim();
                const naOption = String(options[2]).toLowerCase().trim();

                normalizedValues.forEach((val) => {
                  if (val === yesOption) {
                    counts.yes += 1;
                    counts.total += 1;
                  } else if (val === noOption) {
                    counts.no += 1;
                    counts.total += 1;
                  } else if (val === naOption) {
                    counts.na += 1;
                    counts.total += 1;
                  }
                });
              } else {
                // Fallback to recognized values
                const hasRecognizedValue = normalizedValues.some((value) =>
                  recognizedYesNoValues.includes(value),
                );
                if (hasRecognizedValue) {
                  counts.total += 1;
                  if (normalizedValues.includes("yes") || normalizedValues.includes("accepted") || normalizedValues.includes("verified")) {
                    counts.yes += 1;
                  }
                  if (normalizedValues.includes("no") || normalizedValues.includes("rejected")) {
                    counts.no += 1;
                  }
                  if (
                    normalizedValues.includes("n/a") ||
                    normalizedValues.includes("na") ||
                    normalizedValues.includes("not applicable") ||
                    normalizedValues.includes("rework") ||
                    normalizedValues.includes("reworked") ||
                    normalizedValues.some(v => v.includes("re-rework"))
                  ) {
                    counts.na += 1;
                  }
                }
              }
            }
          }
        });

        question.followUpQuestions?.forEach(processQuestion);
      };

      section.questions?.forEach(processQuestion);

      if (!counts.total) {
        return null;
      }

      return {
        id: section.id,
        title: section.title || "Untitled Section",
        yes: counts.yes,
        no: counts.no,
        na: counts.na,
        accepted: counts.accepted,
        rejected: counts.rejected,
        rework: counts.rework,
        total: counts.total,
      };
    }) ?? [];

  return stats.filter((stat): stat is SectionPerformanceStat => Boolean(stat));
};



const computeQuestionPerformanceStats = (
  form: Form | null,
  responses: Response[],
): QuestionPerformanceStat[] => {
  if (!form?.sections || !responses.length) {
    return [];
  }

  const allQuestionStats: QuestionPerformanceStat[] = [];

  form.sections.forEach((section) => {
    section.questions.forEach((question: any) => {
      const counts = {
        accepted: 0,
        rejected: 0,
        rework: 0,
        total: 0,
      };

      responses.forEach((response) => {
        const answer = response.answers?.[question.id];
        if (answer === null || answer === undefined || answer === "") {
          return;
        }

        if (typeof answer === "object" && answer.status) {
          const status = String(answer.status).toLowerCase().trim();
          if (
            status === "accepted" ||
            status === "rework completed" ||
            status === "verified"
          ) {
            counts.accepted += 1;
            counts.total += 1;
          } else if (status === "rejected") {
            counts.rejected += 1;
            counts.total += 1;
          } else if (
            status === "rework" ||
            status === "reworked" ||
            status.includes("re-rework")
          ) {
            counts.rework += 1;
            counts.total += 1;
          }
        } else {
          const answerStr = String(answer).toLowerCase().trim();
          const normalizedValues = extractYesNoValues(answer);

          if (
            answerStr === "accepted" ||
            answerStr === "rework completed" ||
            answerStr === "verified"
          ) {
            counts.accepted += 1;
            counts.total += 1;
          } else if (answerStr === "rejected") {
            counts.rejected += 1;
            counts.total += 1;
          } else if (
            answerStr === "rework" ||
            answerStr === "reworked" ||
            answerStr.includes("re-rework")
          ) {
            counts.rework += 1;
            counts.total += 1;
          } else if (
            question.type === "yesNoNA" ||
            question.type === "chassisWithZone" ||
            question.type === "chassisWithoutZone" ||
            question.type === "chassis" ||
            question.type === "zone-in" ||
            question.type === "zone-out" ||
            question.text?.toLowerCase().includes("chassis")
          ) {
            const options = question.options || [];
            if (options.length >= 3) {
              const yesOption = String(options[0]).toLowerCase().trim();
              const noOption = String(options[1]).toLowerCase().trim();
              const naOption = String(options[2]).toLowerCase().trim();

              normalizedValues.forEach((val) => {
                if (val === yesOption) {
                  counts.accepted += 1;
                  counts.total += 1;
                } else if (val === noOption) {
                  counts.rejected += 1;
                  counts.total += 1;
                } else if (val === naOption) {
                  counts.rework += 1;
                  counts.total += 1;
                }
              });
            } else {
              const hasRecognizedValue = normalizedValues.some((value) =>
                recognizedYesNoValues.includes(value),
              );
              if (hasRecognizedValue) {
                counts.total += 1;
                if (normalizedValues.includes("yes") || normalizedValues.includes("accepted") || normalizedValues.includes("verified")) {
                  counts.accepted += 1;
                }
                if (normalizedValues.includes("no") || normalizedValues.includes("rejected")) {
                  counts.rejected += 1;
                }
                if (
                  normalizedValues.includes("n/a") ||
                  normalizedValues.includes("na") ||
                  normalizedValues.includes("not applicable") ||
                  normalizedValues.includes("rework") ||
                  normalizedValues.includes("reworked") ||
                  normalizedValues.some(v => v.includes("re-rework"))
                ) {
                  counts.rework += 1;
                }
              }
            }
          }
        }
      });

      if (counts.total > 0) {
        allQuestionStats.push({
          id: question.id,
          text: question.text,
          sectionTitle: section.title,
          accepted: counts.accepted,
          rejected: counts.rejected,
          rework: counts.rework,
          total: counts.total,
        });
      }
    });
  });

  return allQuestionStats;
};

type DailyPerformanceStat = {
  date: string;
  dateKey: string;
  totalResponses: number;
  reworkCount: number;
  acceptedCount: number;
};

const computeDailyPerformanceStats = (
  responses: Response[],
  statuses: Record<string, string>,
  startDate?: string,
  endDate?: string,
): DailyPerformanceStat[] => {
  const dailyMap = new Map<string, { total: number; rework: number; accepted: number }>();

  let start: Date | null = null;
  let end: Date | null = null;

  if (startDate) {
    start = new Date(startDate);
  } else {
    // Default to start of current month if no start date provided
    start = new Date();
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
  }

  if (endDate) {
    end = new Date(endDate);
  } else {
    // Default to today if no end date provided
    end = new Date();
    end.setHours(23, 59, 59, 999);
  }

  if (responses.length > 0 && !startDate && !endDate) {
    const timestamps = responses.map((r) =>
      new Date(getResponseTimestamp(r) || 0).getTime(),
    );
    const minTS = Math.min(...timestamps);
    const maxTS = Math.max(...timestamps);

    // Expand range to include responses if they are outside current month
    if (minTS < start.getTime()) start = new Date(minTS);
    if (maxTS > end.getTime()) end = new Date(maxTS);
  }

  if (start && end) {
    const curr = new Date(start);
    curr.setHours(0, 0, 0, 0);
    const last = new Date(end);
    last.setHours(0, 0, 0, 0);

    while (curr <= last) {
      const dKey = curr.toISOString().split("T")[0];
      dailyMap.set(dKey, { total: 0, rework: 0, accepted: 0 });
      curr.setDate(curr.getDate() + 1);
    }
  }

  responses.forEach((response) => {
    const timestamp = getResponseTimestamp(response);
    if (!timestamp) return;

    const dateKey = new Date(timestamp).toISOString().split("T")[0];
    if (!dailyMap.has(dateKey)) {
      dailyMap.set(dateKey, { total: 0, rework: 0, accepted: 0 });
    }

    const dayStats = dailyMap.get(dateKey)!;
    dayStats.total += 1;

    const status = statuses[response.id];
    if (status) {
      if (status.startsWith("Rework") || status === "Rework Accepted") {
        dayStats.rework += 1;
      } else if (status === "Direct Ok" || status === "Accepted") {
        dayStats.accepted += 1;
      }
    }
  });

  return Array.from(dailyMap.entries())
    .map(([dateKey, stats]) => {
      const dateObj = new Date(dateKey);
      const formattedDate = dateObj.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
      });
      return {
        date: formattedDate,
        dateKey,
        totalResponses: stats.total,
        reworkCount: stats.rework,
        acceptedCount: stats.accepted,
      };
    })
    .sort((a, b) => (a.dateKey > b.dateKey ? 1 : -1));
};

const computeMonthlyPerformanceStats = (
  responses: Response[],
  statuses: Record<string, string>,
  startDate?: string,
  endDate?: string,
): DailyPerformanceStat[] => {
  const monthMap = new Map<string, { total: number; rework: number; accepted: number }>();

  let start: Date | null = null;
  let end: Date | null = null;

  if (startDate) {
    start = new Date(startDate);
  } else {
    // Default to start of current month if no start date provided
    start = new Date();
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
  }

  if (endDate) {
    end = new Date(endDate);
  } else {
    // Default to today if no end date provided
    end = new Date();
    end.setHours(23, 59, 59, 999);
  }

  if (responses.length > 0 && !startDate && !endDate) {
    const timestamps = responses.map((r) =>
      new Date(getResponseTimestamp(r) || 0).getTime(),
    );
    const minTS = Math.min(...timestamps);
    const maxTS = Math.max(...timestamps);

    // Expand range to include responses if they are outside current month
    if (minTS < start.getTime()) start = new Date(minTS);
    if (maxTS > end.getTime()) end = new Date(maxTS);
  }

  if (start && end) {
    const curr = new Date(start);
    curr.setDate(1);
    const last = new Date(end);
    last.setDate(1);

    while (curr <= last) {
      const monthKey = `${curr.getFullYear()}-${String(
        curr.getMonth() + 1,
      ).padStart(2, "0")}`;
      monthMap.set(monthKey, { total: 0, rework: 0, accepted: 0 });
      curr.setMonth(curr.getMonth() + 1);
    }
  }

  responses.forEach((response) => {
    const timestamp = getResponseTimestamp(response);
    if (!timestamp) return;

    const dateObj = new Date(timestamp);
    const monthKey = `${dateObj.getFullYear()}-${String(
      dateObj.getMonth() + 1,
    ).padStart(2, "0")}`;

    if (!monthMap.has(monthKey)) {
      monthMap.set(monthKey, { total: 0, rework: 0, accepted: 0 });
    }

    const monthStats = monthMap.get(monthKey)!;
    monthStats.total += 1;

    const status = statuses[response.id];
    if (status) {
      if (status.startsWith("Rework") || status === "Rework Accepted") {
        monthStats.rework += 1;
      } else if (status === "Direct Ok" || status === "Accepted") {
        monthStats.accepted += 1;
      }
    }
  });

  return Array.from(monthMap.entries())
    .map(([monthKey, stats]) => {
      const [year, month] = monthKey.split("-");
      const dateObj = new Date(parseInt(year), parseInt(month) - 1);
      const formattedMonth = dateObj.toLocaleDateString("en-US", {
        month: "short",
        year: "numeric",
      });
      return {
        date: formattedMonth,
        dateKey: monthKey,
        totalResponses: stats.total,
        reworkCount: stats.rework,
        acceptedCount: stats.accepted,
      };
    })
    .sort((a, b) => (a.dateKey > b.dateKey ? 1 : -1));
};

const computeDirectAcceptedDailyStats = (
  responses: Response[],
  statuses: Record<string, string>,
  startDate?: string,
  endDate?: string,
): {
  date: string;
  dateKey: string;
  directCount: number;
  reworkCount: number;
  rejectedCount: number;
  total: number;
  questionReworkCount: number;
  reworkCompletedCount: number;
}[] => {
  const dailyMap = new Map<string, { total: number; direct: number; rework: number; rejected: number; questionRework: number; reworkCompleted: number }>();

  let start: Date | null = null;
  let end: Date | null = null;

  if (startDate) {
    start = new Date(startDate);
  } else {
    start = new Date();
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
  }

  if (endDate) {
    end = new Date(endDate);
  } else {
    end = new Date();
    end.setHours(23, 59, 59, 999);
  }

  if (start && end) {
    const curr = new Date(start);
    curr.setHours(0, 0, 0, 0);
    const last = new Date(end);
    last.setHours(0, 0, 0, 0);

    while (curr <= last) {
      const dKey = curr.toISOString().split("T")[0];
      dailyMap.set(dKey, { total: 0, direct: 0, rework: 0, rejected: 0, questionRework: 0, reworkCompleted: 0 });
      curr.setDate(curr.getDate() + 1);
    }
  }

  responses.forEach((response) => {
    const timestamp = getResponseTimestamp(response);
    if (!timestamp) return;

    const dateKey = new Date(timestamp).toISOString().split("T")[0];
    if (!dailyMap.has(dateKey)) {
      dailyMap.set(dateKey, { total: 0, direct: 0, rework: 0, rejected: 0, questionRework: 0, reworkCompleted: 0 });
    }

    const dayStats = dailyMap.get(dateKey)!;
    dayStats.total += 1;

    // Calculate rework count based on individual questions
    let formReworkQuestionsCount = 0;
    if (response.answers) {
      Object.values(response.answers).forEach((ans) => {
        if (typeof ans === "object" && ans !== null && (ans as any).status) {
          const s = String((ans as any).status).toLowerCase().trim();
          if (s === "rework" || s === "reworked" || s.includes("re-rework")) {
            formReworkQuestionsCount++;
          }
        } else if (typeof ans === "string") {
          const s = ans.toLowerCase().trim();
          if (s === "rework" || s === "reworked" || s.includes("re-rework")) {
            formReworkQuestionsCount++;
          }
        }
      });
    }
    dayStats.questionRework += formReworkQuestionsCount;

    const status = statuses[response.id];
    if (status === "Direct Ok" || status === "Accepted") {
      dayStats.direct += 1;
    } else if (status && (status.startsWith("Rework Accepted") || status.startsWith("Rework Completed"))) {
      dayStats.reworkCompleted += 1;
    } else if (status && status.startsWith("Rework")) {
      dayStats.rework += 1;
    } else if (status === "Rejected") {
      dayStats.rejected += 1;
    }
  });

  return Array.from(dailyMap.entries())
    .map(([dateKey, stats]) => {
      const dateObj = new Date(dateKey);
      const formattedDate = dateObj.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
      });
      return {
        date: formattedDate,
        dateKey,
        directCount: stats.direct,
        reworkCount: stats.rework,
        rejectedCount: stats.rejected,
        total: stats.total,
        questionReworkCount: stats.questionRework,
        reworkCompletedCount: stats.reworkCompleted,
      };
    })
    .sort((a, b) => (a.dateKey > b.dateKey ? 1 : -1));
};

const computeDailyReworkVolumeStats = (
  form: Form | null,
  responses: Response[],
  startDate?: string,
  endDate?: string,
): { date: string; dateKey: string; reworkCount: number }[] => {
  const dailyMap = new Map<string, { rework: number }>();

  let start: Date | null = null;
  let end: Date | null = null;

  if (startDate) {
    start = new Date(startDate);
  } else {
    start = new Date();
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
  }

  if (endDate) {
    end = new Date(endDate);
  } else {
    end = new Date();
    end.setHours(23, 59, 59, 999);
  }

  if (start && end) {
    const curr = new Date(start);
    curr.setHours(0, 0, 0, 0);
    const last = new Date(end);
    last.setHours(0, 0, 0, 0);

    while (curr <= last) {
      const dKey = curr.toISOString().split("T")[0];
      dailyMap.set(dKey, { rework: 0 });
      curr.setDate(curr.getDate() + 1);
    }
  }

  if (!form?.sections) return [];

  responses.forEach((response) => {
    const timestamp = getResponseTimestamp(response);
    if (!timestamp) return;

    const dateKey = new Date(timestamp).toISOString().split("T")[0];
    if (!dailyMap.has(dateKey)) {
      dailyMap.set(dateKey, { rework: 0 });
    }

    const dayStats = dailyMap.get(dateKey)!;

    form.sections.forEach((section) => {
      section.questions.forEach((question: any) => {
        const answer = response.answers?.[question.id];
        if (answer === null || answer === undefined || answer === "") return;

        if (typeof answer === "object" && answer.status) {
          const status = String(answer.status).toLowerCase().trim();
          if (
            status === "rework" ||
            status === "reworked" ||
            status.includes("re-rework")
          ) {
            dayStats.rework += 1;
          }
        } else {
          const answerStr = String(answer).toLowerCase().trim();
          if (
            answerStr === "rework" ||
            answerStr === "reworked" ||
            answerStr.includes("re-rework")
          ) {
            dayStats.rework += 1;
          } else if (
            question.type === "yesNoNA" ||
            question.type === "chassisWithZone" ||
            question.type === "chassisWithoutZone" ||
            question.type === "chassis" ||
            question.type === "zone-in" ||
            question.type === "zone-out" ||
            question.text?.toLowerCase().includes("chassis")
          ) {
            const options = question.options || [];
            if (options.length >= 3) {
              const naOption = String(options[2]).toLowerCase().trim();
              const normalizedValues = extractYesNoValues(answer);
              normalizedValues.forEach((val) => {
                if (val === naOption) {
                  dayStats.rework += 1;
                }
              });
            }
          }
        }
      });
    });
  });

  return Array.from(dailyMap.entries())
    .map(([dateKey, stats]) => {
      const dateObj = new Date(dateKey);
      const formattedDate = dateObj.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
      });
      return {
        date: formattedDate,
        dateKey,
        reworkCount: stats.rework,
      };
    })
    .sort((a, b) => (a.dateKey > b.dateKey ? 1 : -1));
};

interface SectionStat {
  id: string;
  title: string;
  yes: number;
  no: number;
  na: number;
  accepted?: number;
  rejected?: number;
  rework?: number;
  total: number;
}

const getSectionYesNoStats = (
  form: any,
  answers: Record<string, any>,
): SectionStat[] => {
  const stats =
    form.sections?.map((section: any) => {
      const counts = {
        yes: 0,
        no: 0,
        na: 0,
        accepted: 0,
        rejected: 0,
        rework: 0,
        total: 0,
      };

      const processQuestion = (question: any) => {
        if (!question) {
          return;
        }

        const answer = answers?.[question.id];
        if (answer !== null && answer !== undefined && answer !== "") {
          // Check if it's an inspection status object
          if (typeof answer === "object" && answer.status) {
            const status = String(answer.status).toLowerCase().trim();
            if (
              status === "accepted" ||
              status === "rework completed" ||
              status === "verified"
            ) {
              counts.accepted += 1;
              counts.total += 1;
            } else if (status === "rejected") {
              counts.rejected += 1;
              counts.total += 1;
            } else if (
              status === "rework" ||
              status === "reworked" ||
              status.includes("re-rework")
            ) {
              counts.rework += 1;
              counts.total += 1;
            }
          } else {
            // Handle string answers that might be inspection statuses
            const answerStr = String(answer).toLowerCase().trim();
            const normalizedValues = extractYesNoValues(answer);

            if (
              answerStr === "accepted" ||
              answerStr === "rework completed" ||
              answerStr === "verified"
            ) {
              counts.accepted += 1;
              counts.total += 1;
            } else if (answerStr === "rejected") {
              counts.rejected += 1;
              counts.total += 1;
            } else if (
              answerStr === "rework" ||
              answerStr === "reworked" ||
              answerStr.includes("re-rework")
            ) {
              counts.rework += 1;
              counts.total += 1;
            } else if (
              question.type === "yesNoNA" ||
              question.type === "chassisWithZone" ||
              question.type === "chassisWithoutZone" ||
              question.type === "chassis" ||
              question.type === "zone-in" ||
              question.type === "zone-out" ||
              question.text?.toLowerCase().includes("chassis")
            ) {
              const options = question.options || [];

              if (options.length >= 3) {
                const yesOption = String(options[0]).toLowerCase().trim();
                const noOption = String(options[1]).toLowerCase().trim();
                const naOption = String(options[2]).toLowerCase().trim();

                normalizedValues.forEach((val) => {
                  if (val === yesOption) {
                    counts.yes += 1;
                    counts.total += 1;
                  } else if (val === noOption) {
                    counts.no += 1;
                    counts.total += 1;
                  } else if (val === naOption) {
                    counts.na += 1;
                    counts.total += 1;
                  }
                });
              } else {
                const hasRecognizedValue = normalizedValues.some((value) =>
                  recognizedYesNoValues.includes(value),
                );
                if (hasRecognizedValue) {
                  counts.total += 1;
                  if (normalizedValues.includes("yes") || normalizedValues.includes("accepted") || normalizedValues.includes("verified")) {
                    counts.yes += 1;
                  }
                  if (normalizedValues.includes("no") || normalizedValues.includes("rejected")) {
                    counts.no += 1;
                  }
                  if (
                    normalizedValues.includes("n/a") ||
                    normalizedValues.includes("na") ||
                    normalizedValues.includes("not applicable") ||
                    normalizedValues.includes("rework") ||
                    normalizedValues.includes("reworked") ||
                    normalizedValues.some(v => v.includes("re-rework"))
                  ) {
                    counts.na += 1;
                  }
                }
              }
            }
          }
        }

        question.followUpQuestions?.forEach(processQuestion);
      };

      section.questions?.forEach(processQuestion);

      if (!counts.total) {
        return null;
      }

      return {
        id: section.id,
        title: section.title || "Untitled Section",
        yes: counts.yes,
        no: counts.no,
        na: counts.na,
        accepted: counts.accepted,
        rejected: counts.rejected,
        rework: counts.rework,
        total: counts.total,
      };
    }) ?? [];

  return stats.filter((stat): stat is SectionStat => Boolean(stat));
};

const QuestionSuggestionRenderer = ({
  question,
  value,
  onChange,
  currentAnswer
}: {
  question: any,
  value: any,
  onChange: (val: any) => void,
  currentAnswer?: any
}) => {
  const [uploading, setUploading] = useState(false);
  const [showCamera, setShowCamera] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // All values stored flat in the value object
  const safeValue = (typeof value === 'object' && value !== null) ? value : {};
  const status = safeValue.inspectionStatus || '';
  const remark = safeValue.inspectionRemark || '';
  const fileUrl = safeValue.inspectionFileUrl || '';

  const update = (patch: Record<string, any>) => {
    onChange({ ...safeValue, ...patch });
  };

  const handleFileUpload = async (file: File) => {
    try {
      setUploading(true);
      const result = await apiClient.uploadFile(file, 'form');
      const url = apiClient.resolveUploadedFileUrl(result);
      if (url) {
        update({ inspectionFileUrl: url });
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    } catch (err) {
      console.error('Upload failed:', err);
      alert('Upload failed. Please try again.');
    } finally {
      setUploading(false);
    }
  };

  // Render the current answer as read-only context


  return (
    <div className="space-y-3 mt-2" onClick={e => e.stopPropagation()}>





      {/* Remark */}
      <div
        className="p-3 bg-amber-50 dark:bg-amber-900/20 rounded-xl border border-amber-200 dark:border-amber-800"
        onClick={e => e.stopPropagation()}
      >
        <label className="text-[10px] font-black text-amber-600 uppercase tracking-widest block mb-1">
          Remark
        </label>
        <textarea
          rows={2}
          value={remark}
          onChange={e => { e.stopPropagation(); update({ inspectionRemark: e.target.value }); }}
          onKeyDown={e => e.stopPropagation()}
          onClick={e => e.stopPropagation()}
          placeholder="Enter remark..."
          className="w-full p-2 text-xs bg-white dark:bg-gray-900 border border-amber-200 dark:border-amber-700 rounded-lg focus:ring-2 focus:ring-amber-400 outline-none resize-none"
        />
      </div>

      {/* File Upload + Camera */}
      <div
        className="p-3 bg-blue-50 dark:bg-blue-900/20 rounded-xl border border-blue-200 dark:border-blue-800"
        onClick={e => e.stopPropagation()}
      >
        <label className="text-[10px] font-black text-blue-500 uppercase tracking-widest block mb-1.5">
          Evidence Photo
        </label>

        {fileUrl ? (
          /* Uploaded — show thumbnail + change/camera */
          <div className="flex gap-2">
            <label className="flex-1 cursor-pointer group relative">
              <input
                ref={fileInputRef}
                type="file"
                className="hidden"
                accept="image/*"
                onClick={e => e.stopPropagation()}
                onChange={async e => {
                  e.stopPropagation();
                  const file = e.target.files?.[0];
                  if (file) await handleFileUpload(file);
                }}
              />
              <img
                src={fileUrl}
                alt="Evidence"
                className="w-full h-28 object-cover rounded-lg border-2 border-emerald-400"
                onClick={e => { e.stopPropagation(); window.open(fileUrl, '_blank'); }}
              />
              <div className="absolute inset-0 flex items-center justify-center bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity rounded-lg pointer-events-none">
                <span className="text-[10px] text-white font-bold">Change</span>
              </div>
            </label>
            <button
              type="button"
              onClick={e => { e.stopPropagation(); setShowCamera(true); }}
              className="flex-1 flex flex-col items-center justify-center gap-1 p-3 border-2 border-dashed border-purple-300 dark:border-purple-700 rounded-lg hover:border-purple-500 transition-colors"
            >
              <Camera className="w-4 h-4 text-purple-500" />
              <span className="text-[9px] text-purple-500 font-bold">Camera</span>
            </button>
          </div>
        ) : uploading ? (
          <div className="flex items-center gap-2 p-3 bg-blue-50 dark:bg-blue-900/10 rounded-lg">
            <Loader2 className="w-4 h-4 animate-spin text-blue-500" />
            <span className="text-xs text-blue-600 dark:text-blue-400">Uploading...</span>
          </div>
        ) : (
          /* Empty — show upload + camera side by side */
          <div className="flex gap-2">
            <label className="flex-1 cursor-pointer">
              <input
                ref={fileInputRef}
                type="file"
                className="hidden"
                accept="image/*"
                onClick={e => e.stopPropagation()}
                onChange={async e => {
                  e.stopPropagation();
                  const file = e.target.files?.[0];
                  if (file) await handleFileUpload(file);
                }}
              />
              <div className="flex flex-col items-center justify-center gap-1 p-3 border-2 border-dashed border-blue-200 dark:border-blue-700 rounded-lg hover:border-blue-400 hover:bg-blue-50/50 dark:hover:bg-blue-900/10 transition-all">
                <Upload className="w-4 h-4 text-blue-400" />
                <span className="text-[9px] text-blue-500 dark:text-blue-400 font-bold">Upload Photo</span>
              </div>
            </label>
            <button
              type="button"
              onClick={e => { e.stopPropagation(); setShowCamera(true); }}
              className="flex-1 flex flex-col items-center justify-center gap-1 p-3 border-2 border-dashed border-purple-200 dark:border-purple-700 rounded-lg hover:border-purple-400 hover:bg-purple-50/50 dark:hover:bg-purple-900/10 transition-colors"
            >
              <Camera className="w-4 h-4 text-purple-400" />
              <span className="text-[9px] text-purple-500 dark:text-purple-400 font-bold">Camera</span>
            </button>
          </div>
        )}
      </div>

      {/* Camera modal */}
      {showCamera && createPortal(
        <CameraCapture
          onCapture={async (file: File) => {
            await handleFileUpload(file);
            setShowCamera(false);
          }}
          onClose={() => setShowCamera(false)}
        />,
        document.body
      )}
    </div>
  );
};

// Function to render message with images
const renderMessageWithImages = (message: string) => {
  if (!message) return null;

  console.log('renderMessageWithImages called with:', message);

  // Split message by image markdown syntax
  const parts = message.split(/(!\[.*?\]\(.*?\))/g);
  console.log('Message parts:', parts);

  return (
    <div className="space-y-2">
      {parts.map((part, index) => {
        // Check if this part is an image markdown
        const imageMatch = part.match(/!\[.*?\]\((.*?)\)/);
        if (imageMatch) {
          const imageUrl = imageMatch[1];
          console.log('Found image URL:', imageUrl);
          return (
            <div key={index} className="inline-block">
              <img
                src={imageUrl}
                alt="Evidence"
                className="max-w-32 max-h-32 object-cover rounded-lg border border-gray-200 dark:border-gray-600 cursor-pointer hover:opacity-90 transition-opacity"
                onClick={() => window.open(imageUrl, '_blank')}
                onError={(e) => {
                  console.error('Image failed to load:', imageUrl);
                  e.currentTarget.style.display = 'none';
                  // Show fallback text
                  const fallback = document.createElement('div');
                  fallback.className = 'text-xs text-red-500 mt-1';
                  fallback.textContent = 'Image failed to load';
                  e.currentTarget.parentNode?.appendChild(fallback);
                }}
                onLoad={() => console.log('Image loaded successfully:', imageUrl)}
              />
            </div>
          );
        }
        // Regular text part
        return part.trim() ? (
          <span key={index} className="whitespace-pre-wrap">{part}</span>
        ) : null;
      })}
    </div>
  );
};
const OPSMasterListTable = ({
  form,
  filteredResponses,
  responseStatuses,
  getResponseTimestamp,
}: {
  form: Form | null;
  filteredResponses: Response[];
  responseStatuses: Record<string, string>;
  getResponseTimestamp: (r: Response) => string | undefined;
}) => {
  const [headerTitle, setHeaderTitle] = useState(form?.title || "Master List of OPS");
  const [division, setDivision] = useState("");
  const [department, setDepartment] = useState("");
  const [headerDate, setHeaderDate] = useState(
    new Date().toLocaleDateString("en-GB", {
      day: "2-digit", month: "2-digit", year: "2-digit",
    }).replace(/\//g, ".")
  );
  const [docCode, setDocCode] = useState("03010-QMHO-F0-002");

  // Find key question IDs
  const chassisQId = useMemo(() => {
    if (!form?.sections) return null;
    for (const section of form.sections) {
      for (const q of section.questions || []) {
        if (
          q.type === "chassis" || q.type === "chassisWithZone" ||
          q.type === "chassisWithoutZone" || q.type === "zone-in" ||
          q.type === "zone-out" || q.text?.toLowerCase().includes("chassis") ||
          q.trackResponseRank
        ) return q.id;
      }
    }
    return null;
  }, [form]);

  const modelQId = useMemo(() => {
    if (!form?.sections) return null;
    for (const section of form.sections) {
      for (const q of section.questions || []) {
        if (q.text?.toLowerCase().includes("model")) return q.id;
      }
    }
    return null;
  }, [form]);

  const descQId = useMemo(() => {
    if (!form?.sections) return null;
    for (const section of form.sections) {
      for (const q of section.questions || []) {
        if (
          q.text?.toLowerCase().includes("description") ||
          q.text?.toLowerCase().includes("process") ||
          q.text?.toLowerCase().includes("station") ||
          q.text?.toLowerCase().includes("operation")
        ) return q.id;
      }
    }
    return null;
  }, [form]);

  const docQId = useMemo(() => {
    if (!form?.sections) return null;
    for (const section of form.sections) {
      for (const q of section.questions || []) {
        if (
          q.text?.toLowerCase().includes("document") ||
          q.text?.toLowerCase().includes("doc no") ||
          q.text?.toLowerCase().includes("control")
        ) return q.id;
      }
    }
    return null;
  }, [form]);

  const getStrVal = (response: Response, qId: string | null): string => {
    if (!qId) return "";
    const val = response.answers?.[qId];
    if (!val) return "";
    if (typeof val === "object") {
      return (val as any).chassisNumber || (val as any).status || "";
    }
    return String(val).trim();
  };

  // 

  // Each submission becomes a row with L-shape pattern
  // 
  const tableRows = useMemo(() => {
    // Sort all responses oldest → newest
    const sorted = [...filteredResponses].sort((a, b) => {
      const tA = new Date(getResponseTimestamp(a) || 0).getTime();
      const tB = new Date(getResponseTimestamp(b) || 0).getTime();
      return tA - tB;
    });

    // Group by MODEL ONLY
    const modelMap = new Map<string, {
      model: string;
      submissions: Response[];
    }>();

    sorted.forEach((response) => {
      // Determine model
      let model = "";
      if (chassisQId) {
        const val = response.answers?.[chassisQId];
        if (typeof val === "object" && val?.chassisNumber) model = val.chassisNumber;
        else if (typeof val === "string") model = val.trim();
      }
      if (!model && modelQId) model = getStrVal(response, modelQId);
      if (!model) model = "Unknown";

      if (model === "Unknown") return;

      if (!modelMap.has(model)) {
        modelMap.set(model, {
          model,
          submissions: [],
        });
      }
      modelMap.get(model)!.submissions.push(response);
    });

    // Build rows - EACH submission becomes a row with L-shape
    let sl = 1;
    const rows: Array<{
      sl: number;
      model: string;
      controlNo: string;
      description: string;
      revDates: (string | null)[];
      remarks: string;
    }> = [];

    modelMap.forEach(({ model, submissions }) => {
      // Sort submissions by date (oldest first) for this model
      const sortedSubs = [...submissions].sort((a, b) => {
        const tA = new Date(getResponseTimestamp(a) || 0).getTime();
        const tB = new Date(getResponseTimestamp(b) || 0).getTime();
        return tA - tB;
      });

      // Get ALL dates from ALL submissions for this model
      const allDates = sortedSubs.map((resp) => {
        const ts = getResponseTimestamp(resp);
        if (!ts) return null;
        const d = new Date(ts);
        return [
          String(d.getDate()).padStart(2, "0"),
          String(d.getMonth() + 1).padStart(2, "0"),
          String(d.getFullYear()).slice(2),
        ].join("-");
      });

      // For each submission, create a row with L-shape
      sortedSubs.forEach((resp, idx) => {
        const revDates: (string | null)[] = [null, null, null, null, null, null];

        // L-SHAPE: Fill columns 0 through idx with dates from submissions 0 through idx
        for (let j = 0; j <= Math.min(idx, 5); j++) {
          if (j < allDates.length && allDates[j]) {
            revDates[j] = allDates[j];
          }
        }

        // Get control number (Document No.) for this specific submission
        const controlNo = getStrVal(resp, docQId) || resp.id.substring(0, 8);

        // Get description for this specific submission
        const description = getStrVal(resp, descQId);

        // Get remark from submission history issuance details
        const history = resp.answers?.__submissionHistory || [];
        let remark = "";

        // Get the issuance details from submission history for this specific submission
        if (history.length > 0 && history[idx]?.issuanceDetails) {
          remark = history[idx].issuanceDetails;
        } else if (idx === 0) {
          remark = "Initial Preparation";
        } else {
          remark = `Revision ${idx}`;
        }

        rows.push({
          sl: sl++,
          model,
          controlNo,
          description,
          revDates,
          remarks: remark,
        });
      });
    });

    // Sort by model name
    rows.sort((a, b) => a.model.localeCompare(b.model));

    return rows;
  }, [filteredResponses, chassisQId, modelQId, descQId, docQId]);

  const b = "1px solid #2563eb";
  const cell: React.CSSProperties = {
    border: b,
    padding: "5px 8px",
    fontSize: "12px",
    color: "var(--color-text-primary)",
    verticalAlign: "middle",
    wordBreak: "break-word",
  };
  const hCell: React.CSSProperties = {
    ...cell,
    background: "var(--color-background-secondary)",
    fontWeight: 500,
    textAlign: "center" as const,
  };
  const inp: React.CSSProperties = {
    background: "transparent",
    border: "none",
    outline: "none",
    width: "100%",
    fontSize: "12px",
    color: "var(--color-text-primary)",
    fontFamily: "inherit",
  };

  const handlePrint = () => {
    const el = document.getElementById("ops-master-list-table");
    if (!el) return;
    const html = `<html><head><style>
      body{font-family:Arial,sans-serif;margin:12px}
      table{border-collapse:collapse;width:100%;font-size:11px}
      td,th{border:1px solid #2563eb;padding:4px 8px;vertical-align:middle}
      input{background:transparent;border:none;font-size:11px;font-family:Arial,sans-serif;width:100%}
    </style></head><body>${el.innerHTML}</body></html>`;
    const w = window.open("", "_blank");
    w?.document.write(html);
    w?.document.close();
    w?.print();
  };

  return (
    <div className="space-y-3">
      <div
        className="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-700 p-3 overflow-x-auto"
        id="ops-master-list-table"
      >
        <table style={{ borderCollapse: "collapse", width: "100%", minWidth: "880px", tableLayout: "fixed" }}>
          <colgroup>
            <col style={{ width: "60px" }} />   {/* S.L */}
            <col style={{ width: "180px" }} />   {/* Model */}
            <col style={{ width: "180px" }} />  {/* Description */}
            <col style={{ width: "240px" }} />  {/* Doc No */}
            <col style={{ width: "100px" }} />   {/* Rev 0 */}
            <col style={{ width: "100px" }} />   {/* Rev 1 */}
            <col style={{ width: "100px" }} />   {/* Rev 2 */}
            <col style={{ width: "100px" }} />   {/* Rev 3 */}
            <col style={{ width: "100px" }} />   {/* Rev 4 */}
            <col style={{ width: "100px" }} />   {/* Rev 5 */}
            <col style={{ width: "auto" }} />   {/* Remarks */}
          </colgroup>

          <tbody>
            {/* ── ROW 1: HMSI | Title | DocCode ── */}
            <tr>
              <td style={{ ...cell, textAlign: "center", fontWeight: 500, fontSize: "13px" }}>
                HMSI
              </td>
              <td style={cell} />
              <td colSpan={8} style={{ ...cell, textAlign: "center" }}>
                <input
                  value={headerTitle}
                  onChange={(e) => setHeaderTitle(e.target.value)}
                  style={{ ...inp, textAlign: "center", fontWeight: 500, fontSize: "13px" }}
                  placeholder="Enter title..."
                />
              </td>
              <td colSpan={0} style={{ ...cell, textAlign: "right", fontWeight: 500, fontSize: "11px" }}>
                <input
                  value={docCode}
                  onChange={(e) => setDocCode(e.target.value)}
                  style={{ ...inp, textAlign: "right", fontSize: "11px" }}
                  placeholder="Doc code..."
                />
              </td>
            </tr>

            {/* ── ROW 2: Division | Department | (empty) | DATE ── */}
            <tr>
              <td colSpan={4} style={cell}>
                <span style={{ fontWeight: 500 }}>Division : </span>
                <input
                  value={division}
                  onChange={(e) => setDivision(e.target.value)}
                  style={{ ...inp, display: "inline", width: "calc(100% - 68px)" }}
                  placeholder="Enter division..."
                />
              </td>
              <td colSpan={6} style={cell}>
                <span style={{ fontWeight: 500 }}>Department : </span>
                <input
                  value={department}
                  onChange={(e) => setDepartment(e.target.value)}
                  style={{ ...inp, display: "inline", width: "calc(100% - 90px)" }}
                  placeholder="Enter department..."
                />
              </td>
              <td colSpan={1} style={{ ...cell, textAlign: "right", fontWeight: 500, fontSize: "11px" }}>
                <span>DATE: </span>
                <input
                  value={headerDate}
                  onChange={(e) => setHeaderDate(e.target.value)}
                  style={{ ...inp, display: "inline", width: "70px", textAlign: "right" }}
                  placeholder="DD.MM.YY"
                />
              </td>
            </tr>

            {/* ── ROW 3: Column headers (rowspan) + "Revision Number" spanning cols 5-10 ── */}
            <tr>
              <td rowSpan={2} style={hCell}>S.L</td>
              <td rowSpan={2} style={hCell}>Model</td>
              <td rowSpan={2} style={hCell}>Description</td>
              <td rowSpan={2} style={hCell}>Document No.</td>
              <td colSpan={6} style={hCell}>Revision Number</td>
              <td rowSpan={2} style={hCell}>Remarks</td>
            </tr>

            {/* ── ROW 4: Sub-headers 0-5 ── */}
            <tr>
              {[0, 1, 2, 3, 4, 5].map((n) => (
                <td key={n} style={hCell}>{n}</td>
              ))}
            </tr>

            {/* ── DATA ROWS ── */}
            {tableRows.length === 0
              ? Array.from({ length: 5 }).map((_, i) => (
                <tr key={`empty-${i}`} style={{ height: "38px" }}>
                  <td style={{ ...cell, textAlign: "center" }}>{i + 1}</td>
                  {Array.from({ length: 10 }).map((__, j) => (
                    <td key={j} style={cell} />
                  ))}
                </tr>
              ))
              : tableRows.map((row) => (
                <tr
                  key={`${row.model}-${row.sl}`}
                  style={{
                    height: "38px",
                  }}
                >
                  {/* S.L */}
                  <td style={{ ...cell, textAlign: "center" }}>{row.sl}</td>

                  {/* Model */}
                  <td style={{ ...cell, textAlign: "center", color: "#2563eb", fontWeight: 500 }}>
                    {row.model}
                  </td>

                  {/* Description */}
                  <td style={cell}>{row.description}</td>

                  {/* Document / Control No */}
                  <td style={{ ...cell, textAlign: "center", fontFamily: "monospace", fontSize: "11px" }}>
                    {row.controlNo}
                  </td>

                  {/* Revision dates 0–5 - L-SHAPE pattern */}
                  {row.revDates.map((date, i) => (
                    <td
                      key={i}
                      style={{
                        ...cell,
                        textAlign: "center",
                        fontSize: "11px",
                        color: date ? "#2563eb" : "transparent",
                        background: date
                          ? "rgba(37,99,235,0.10)"
                          : "transparent",
                        fontWeight: date ? "600" : "normal",
                      }}
                    >
                      {date ?? ""}
                    </td>
                  ))}

                  {/* Remarks - shows issuance details from submission history */}
                  <td style={{ ...cell, fontSize: "11px" }}>
                    {row.remarks}
                  </td>
                </tr>
              ))}
          </tbody>
        </table>
      </div>

      <div className="flex items-center justify-between px-1">

        <button
          onClick={handlePrint}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors"
        >
          <Download className="w-3.5 h-3.5" />
          Print / Export
        </button>
      </div>
    </div>
  );
};

export default function FormAnalyticsDashboard() {
  const [selectedForm, setSelectedForm] = useState<Form | null>(null);
  // Add this with other state declarations
  const [uploadingFileId, setUploadingFileId] = useState<string | null>(null);
  const fileInputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  // Add this function near other handlers (around handleEditStart)
  const handleFileUpload = async (questionId: string, file: File) => {
    try {
      setUploadingFileId(questionId);
      const result = await apiClient.uploadFile(file, 'form');
      const url = apiClient.resolveUploadedFileUrl(result);
      if (url) {
        // Get current value for this question
        const currentVal = editFormData[questionId];
        let urls: string[] = [];

        if (currentVal) {
          if (Array.isArray(currentVal)) {
            urls = currentVal.map(item => typeof item === 'string' ? item : item?.url).filter(Boolean);
          } else if (typeof currentVal === 'string') {
            try {
              const parsed = JSON.parse(currentVal);
              if (Array.isArray(parsed)) {
                urls = parsed.map(item => typeof item === 'string' ? item : item?.url).filter(Boolean);
              } else if (parsed?.url) {
                urls = [parsed.url];
              }
            } catch {
              urls = [currentVal];
            }
          } else if (typeof currentVal === 'object' && currentVal?.url) {
            urls = [currentVal.url];
          }
        }

        // Add new URL
        urls.push(url);

        setEditFormData({
          ...editFormData,
          [questionId]: urls
        });

        if (fileInputRefs.current[questionId]) {
          fileInputRefs.current[questionId]!.value = '';
        }
      }
    } catch (err) {
      console.error('Upload failed:', err);
      alert('Upload failed. Please try again.');
    } finally {
      setUploadingFileId(null);
    }
  };

  const handleRemoveImage = (questionId: string, index: number) => {
    const currentVal = editFormData[questionId];
    let urls: string[] = [];

    if (currentVal) {
      if (Array.isArray(currentVal)) {
        urls = currentVal.map(item => typeof item === 'string' ? item : item?.url).filter(Boolean);
      } else if (typeof currentVal === 'string') {
        try {
          const parsed = JSON.parse(currentVal);
          if (Array.isArray(parsed)) {
            urls = parsed.map(item => typeof item === 'string' ? item : item?.url).filter(Boolean);
          } else if (parsed?.url) {
            urls = [parsed.url];
          }
        } catch {
          urls = [currentVal];
        }
      } else if (typeof currentVal === 'object' && currentVal?.url) {
        urls = [currentVal.url];
      }
    }

    const newUrls = urls.filter((_, i) => i !== index);
    setEditFormData({
      ...editFormData,
      [questionId]: newUrls.length > 0 ? newUrls : undefined
    });
  };

  const getFileUrls = (questionId: string): string[] => {
    const val = editFormData[questionId];
    if (!val) return [];
    if (Array.isArray(val)) {
      return val.map(item => typeof item === 'string' ? item : item?.url).filter(Boolean);
    }
    if (typeof val === 'string') {
      try {
        const parsed = JSON.parse(val);
        if (Array.isArray(parsed)) {
          return parsed.map(item => typeof item === 'string' ? item : item?.url).filter(Boolean);
        }
        if (parsed?.url) return [parsed.url];
      } catch {
        // Not JSON
      }
      return [val];
    }
    if (typeof val === 'object' && val?.url) {
      return [val.url];
    }
    return [];
  };

  const generateTableBarChart = (
    yesPercent: number,
    noPercent: number,
    naPercent: number,
  ) => {
    const totalWidth = 160;
    const yesWidth = (yesPercent / 100) * totalWidth;
    const noWidth = (noPercent / 100) * totalWidth;
    const naWidth = (naPercent / 100) * totalWidth;

    return (
      <div
        className="relative"
        style={{
          width: `${totalWidth}px`,
          height: "20px",
        }}
      >
        <div className="absolute inset-0 bg-gray-100 dark:bg-gray-700 rounded-sm border border-gray-300 dark:border-gray-600"></div>

        {yesPercent > 0 && (
          <div
            className="absolute left-0 h-full bg-green-500"
            style={{ width: `${yesWidth}px` }}
          >
            {yesPercent >= 10 && (
              <div className="absolute inset-0 flex items-center justify-center">
                <span
                  className="text-xs font-bold text-white"
                  style={{
                    textShadow: "0 0 2px rgba(0,0,0,0.5)",
                  }}
                >
                  {yesPercent.toFixed(0)}%
                </span>
              </div>
            )}
          </div>
        )}

        {noPercent > 0 && (
          <div
            className="absolute h-full bg-red-500"
            style={{
              left: `${yesWidth}px`,
              width: `${noWidth}px`,
            }}
          >
            {noPercent >= 10 && (
              <div className="absolute inset-0 flex items-center justify-center">
                <span
                  className="text-xs font-bold text-white"
                  style={{
                    textShadow: "0 0 2px rgba(0,0,0,0.5)",
                  }}
                >
                  {noPercent.toFixed(0)}%
                </span>
              </div>
            )}
          </div>
        )}

        {naPercent > 0 && (
          <div
            className="absolute h-full bg-gray-400"
            style={{
              left: `${yesWidth + noWidth}px`,
              width: `${naWidth}px`,
            }}
          >
            {naPercent >= 10 && (
              <div className="absolute inset-0 flex items-center justify-center">
                <span
                  className="text-xs font-bold text-white"
                  style={{
                    textShadow: "0 0 2px rgba(0,0,0,0.5)",
                  }}
                >
                  {naPercent.toFixed(0)}%
                </span>
              </div>
            )}
          </div>
        )}

        {yesPercent > 0 && yesPercent < 10 && (
          <div className="absolute" style={{ left: "2px", top: "1px" }}>
            <span className="text-[9px] font-bold text-green-700 bg-white/80 px-0.5 rounded">
              {yesPercent.toFixed(0)}%
            </span>
          </div>
        )}
        {noPercent > 0 && noPercent < 10 && (
          <div
            className="absolute"
            style={{
              left: `${yesWidth + 2}px`,
              top: "1px",
            }}
          >
            <span className="text-[9px] font-bold text-red-700 bg-white/80 px-0.5 rounded">
              {noPercent.toFixed(0)}%
            </span>
          </div>
        )}
        {naPercent > 0 && naPercent < 10 && (
          <div
            className="absolute"
            style={{
              left: `${yesWidth + noWidth + 2}px`,
              top: "1px",
            }}
          >
            <span className="text-[9px] font-bold text-gray-700 bg-white/80 px-0.5 rounded">
              {naPercent.toFixed(0)}%
            </span>
          </div>
        )}
      </div>
    );
  };
  const [selectedQuestionIndex, setSelectedQuestionIndex] = useState<number | null>(null);
  const { darkMode } = useTheme();
  const { user } = useAuth();
  const isInspector = user?.role === "inspector";

  const renderSuggestion = (suggestion: any) => {
    if (!suggestion || Object.keys(suggestion).length === 0) return null;

    // Check if this is a chassis-with-zone or chassis-without-zone structure
    const isChassisStructure = suggestion.status !== undefined ||
      suggestion.chassisNumber !== undefined ||
      suggestion.zonesData !== undefined ||
      suggestion.zone !== undefined;

    if (isChassisStructure) {
      // Format chassis structure with proper headings
      const sections: JSX.Element[] = [];

      // Status
      if (suggestion.status && suggestion.status.trim()) {
        const statusColor =
          suggestion.status.toLowerCase() === 'accepted' ? 'text-green-600 bg-green-50 border-green-200' :
            suggestion.status.toLowerCase() === 'rejected' ? 'text-red-600 bg-red-50 border-red-200' :
              'text-amber-600 bg-amber-50 border-amber-200';

        sections.push(
          <div key="status" className="flex items-center gap-2 p-2 rounded-lg border" style={{ backgroundColor: 'rgba(var(--status-bg), 0.1)' }}>
            <span className={`px-2 py-1 rounded-md text-[11px] font-black uppercase ${statusColor} border shadow-sm`}>
              Status: {suggestion.status}
            </span>
          </div>
        );
      }

      // Chassis Number
      if (suggestion.chassisNumber && suggestion.chassisNumber.trim()) {
        sections.push(
          <div key="chassis" className="p-2 rounded-lg border border-gray-200 dark:border-gray-700">
            <span className="text-[9px] font-bold text-gray-400 uppercase tracking-wider block mb-1">Chassis Number</span>
            <span className="text-[11px] font-mono font-bold text-gray-700 dark:text-gray-300">{suggestion.chassisNumber}</span>
          </div>
        );
      }

      // Handle zonesData (with zones)
      if (suggestion.zonesData && typeof suggestion.zonesData === 'object') {
        Object.entries(suggestion.zonesData).forEach(([zoneName, zoneData]: [string, any]) => {
          if (zoneData?.categories && Array.isArray(zoneData.categories)) {
            zoneData.categories.forEach((category: any) => {
              const categoryName = category.name;
              const defects = category.defects || [];

              defects.forEach((defect: any) => {
                const defectName = defect.name;
                const remark = defect.details?.remark || defect.remark || '';
                const fileUrl = defect.details?.fileUrl || defect.fileUrl || '';

                sections.push(
                  <div key={`${zoneName}-${categoryName}-${defectName}`} className="p-3 rounded-lg border-l-4 border-indigo-400 bg-indigo-50/30 dark:bg-indigo-900/20 space-y-2">
                    <div className="grid grid-cols-1 gap-2">
                      {zoneName && (
                        <div>
                          <span className="text-[9px] font-bold text-blue-500 uppercase tracking-wider">Zone</span>
                          <p className="text-[11px] font-bold text-gray-700 dark:text-gray-300">{zoneName}</p>
                        </div>
                      )}
                      {categoryName && (
                        <div>
                          <span className="text-[9px] font-bold text-purple-500 uppercase tracking-wider">Category</span>
                          <p className="text-[11px] font-bold text-gray-700 dark:text-gray-300">{categoryName}</p>
                        </div>
                      )}
                      {defectName && (
                        <div>
                          <span className="text-[9px] font-bold text-amber-500 uppercase tracking-wider">Defect</span>
                          <p className="text-[11px] font-bold text-gray-700 dark:text-gray-300">{defectName}</p>
                        </div>
                      )}
                      {remark && remark.trim() && (
                        <div>
                          <span className="text-[9px] font-bold text-gray-400 uppercase tracking-wider">Remark</span>
                          <p className="text-[10px] italic text-gray-600 dark:text-gray-400">"{remark}"</p>
                        </div>
                      )}
                      {fileUrl && fileUrl.trim() && (
                        <div>
                          <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Evidence :</span>
                          <a
                            href={fileUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex items-center gap-1 text-[10px] font-bold text-indigo-600 hover:underline mt-0 "
                          >
                            <Eye className="w-4 h-2 mt-1 " />
                            View Evidence
                          </a>
                        </div>
                      )}
                    </div>
                  </div>
                );
              });
            });
          }
        });
      }

      return (
        <div className="space-y-2">
          {sections}
        </div>
      );
    }

    // Default: for simple fields like text, select, radio
    return (
      <div className="p-2.5 bg-indigo-50 dark:bg-indigo-900/20 rounded-xl border border-indigo-100 dark:border-indigo-800">
        <span className="text-[9px] font-black text-indigo-500 uppercase tracking-widest block mb-1">Revised Answer</span>
        <div className="text-[11px] font-bold text-gray-800 dark:text-gray-200">
          {renderAnswerDisplay(suggestion, { type: 'text' } as any)}
        </div>
      </div>
    );
  };
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const location = useLocation();

  // Guest mode detection
  const isGuest = useMemo(() => {
    const searchParams = new URLSearchParams(location.search);
    return searchParams.get("guest") === "true" || !!localStorage.getItem("guest_auth_token");
  }, [location.search]);

  const handleLogout = () => {
    if (isGuest) {
      localStorage.removeItem("guest_auth_token");
      localStorage.removeItem("guest_email");
      localStorage.removeItem("guest_form_id");
      localStorage.removeItem("guest_expires_at");
      navigate(`/forms/${id}/analytics/login`);
    }
  };

  const handleShareAnalytics = () => {
    if (id) {
      setShareAnalyticsModal({
        open: true,
        formId: id,
        formTitle: form?.title || "Form Analytics"
      });
    }
  };

  const handleAutoSendSetup = () => {
    if (id) {
      setAutoSendModal({
        open: true,
        formId: id,
        formTitle: form?.title || "Form Analytics"
      });
    }
  };

  const [responses, setResponses] = useState<Response[]>([]);

  const [form, setForm] = useState<Form | null>(null);

  const [loading, setLoading] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [autoOpenSectionId, setAutoOpenSectionId] = useState<string | null>(
    null,
  );
  const [analyticsView, setAnalyticsView] = useState<
    "question" | "section" | "table" | "responses" | "dashboard" | "comparison" | "opsTable"
  >(isGuest ? "dashboard" : user?.role === "inspector" ? "responses" : "dashboard");
  const [tableViewType, setTableViewType] = useState<"question" | "section">(
    "question",
  );
  const [selectedSectionIds, setSelectedSectionIds] = useState<string[]>([]);

  const [selectedQuestionId, setSelectedQuestionId] = useState<string>("");
  const [selectedQuestion, setSelectedQuestion] = useState<any>(null);
  const [filterValues, setFilterValues] = useState<string[]>([]);
  const [selectedAnswers, setSelectedAnswers] = useState<string[]>([]);

  const [showFilterModal, setShowFilterModal] = useState(false);
  const [showSectionSelector, setShowSectionSelector] = useState(false);
  const [shareAnalyticsModal, setShareAnalyticsModal] = useState<{
    open: boolean;
    formId: string;
    formTitle: string;
  }>({ open: false, formId: "", formTitle: "" });
  const [autoSendModal, setAutoSendModal] = useState<{
    open: boolean;
    formId: string;
    formTitle: string;
  }>({ open: false, formId: "", formTitle: "" });
  const [appliedFilters, setAppliedFilters] = useState<
    Array<{ id: string; label: string; value: string }>
  >([]);
  const [cascadingFilters, setCascadingFilters] = useState<
    Record<string, string[]>
  >({});
  const [showChatModal, setShowChatModal] = useState(false);
  const [chatResponse, setChatResponse] = useState<Response | null>(null);
  const [chatFilters, setChatFilters] = useState(() => {
    // Load from localStorage to persist suggestedAnswers between modal opens
    try {
      const saved = localStorage.getItem('chatFilters');
      if (saved) {
        const parsed = JSON.parse(saved);
        return {
          chassisNumber: "",
          location: "",
          questions: [] as string[],
          selectedCategories: {} as Record<string, string[]>,
          suggestedAnswers: parsed.suggestedAnswers || {},
          zoneType: "both" as "with" | "without" | "both",
        };
      }
    } catch (error) {
      console.error('Failed to load chatFilters from localStorage:', error);
    }

    return {
      chassisNumber: "",
      location: "",
      questions: [] as string[],
      selectedCategories: {} as Record<string, string[]>,
      suggestedAnswers: {} as Record<string, any>,
      zoneType: "both" as "with" | "without" | "both",
    };
  });
  const [searchParams] = useSearchParams();
  const viewParam = searchParams.get("view");

  useEffect(() => {
    if (
      viewParam === "dashboard" ||
      viewParam === "question" ||
      viewParam === "section" ||
      viewParam === "responses" ||
      viewParam === "comparison" ||
      viewParam === "opsTable" ||
      viewParam === "table"
    ) {
      setAnalyticsView(viewParam);
    }
  }, [viewParam]);

  const [chatMessages, setChatMessages] = useState<any[]>([]);
  const [newMessage, setNewMessage] = useState("");
  const [isSendingMessage, setIsSendingMessage] = useState(false);
  const [selectedInspectorForTrend, setSelectedInspectorForTrend] = useState<string>("Overall");
  const [dateFilter, setDateFilter] = useState<{
    type: "all" | "single" | "range";
    startDate: string;
    endDate: string;
  }>({ type: "all", startDate: "", endDate: "" });
  const [locationFilter, setLocationFilter] = useState<string[]>([]);
  const [columnFilters, setColumnFilters] = useState<
    Record<string, string[] | null>
  >({});
  const [selectedResponsesSectionIds, setSelectedResponsesSectionIds] =
    useState<string[]>([]);
  const [showResponsesFilter, setShowResponsesFilter] = useState(false);
  const [editingResponseId, setEditingResponseId] = useState<string | null>(
    null,
  );
  const [editFormData, setEditFormData] = useState<Record<string, any>>({});
  const [editFormStatus, setEditFormStatus] = useState<string>("Accepted");
  const [editFormNotes, setEditFormNotes] = useState<string>("");
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [deletingResponseId, setDeletingResponseId] = useState<string | null>(
    null,
  );
  const [isSaving, setIsSaving] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [selectedResponseIds, setSelectedResponseIds] = useState<string[]>([]);
  const [showBulkDeleteConfirm, setShowBulkDeleteConfirm] = useState(false);
  const [showActionMenuModal, setShowActionMenuModal] = useState(false);
  const [actionResponse, setActionResponse] = useState<Response | null>(null);

  // Performance scoring system
  const [performanceScores, setPerformanceScores] = useState<Record<string, number>>({});

  const [performanceTableData, setPerformanceTableData] = useState<any[]>([]);
  const [performanceTableLoading, setPerformanceTableLoading] = useState(false);
  const [performancePage, setPerformancePage] = useState(1);
  const [performancePageSize, setPerformancePageSize] = useState(10);
  const [defectStartDate, setDefectStartDate] = useState<string>("");
  const [defectEndDate, setDefectEndDate] = useState<string>("");
  const [trendStartDate, setTrendStartDate] = useState<string>("");
  const [trendEndDate, setTrendEndDate] = useState<string>("");
  const [qualityStartDate, setQualityStartDate] = useState<string>("");
  const [qualityEndDate, setQualityEndDate] = useState<string>("");
  const [sectionStartDate, setSectionStartDate] = useState<string>("");
  const [sectionEndDate, setSectionEndDate] = useState<string>("");
  const [directAcceptedStartDate, setDirectAcceptedStartDate] = useState<string>("");
  const [directAcceptedEndDate, setDirectAcceptedEndDate] = useState<string>("");
  const allQuestionsWithSections = useMemo(() => {
    if (!form?.sections) return [];
    const questionsWithSections: Array<{
      id: string;
      text: string;
      type?: string;
      sectionId: string;
      sectionTitle: string;
    }> = [];
    form.sections.forEach((section) => {
      (section.questions || []).forEach((q: any) => {
        // Only include main questions (not follow-ups) for filtering
        if (!q.parentId && !q.showWhen?.questionId) {
          questionsWithSections.push({
            id: q.id,
            text: q.text || "Unnamed Question",
            type: q.type,
            sectionId: section.id,
            sectionTitle: section.title || "Untitled Section"
          });
        }
      });
    });

    return questionsWithSections;
  }, [form]);
  // Fetch performance table data
  useEffect(() => {
    const fetchPerformanceTable = async () => {
      if (!id || (user?.role !== 'admin' && user?.role !== 'superadmin')) return;

      setPerformanceTableLoading(true);
      try {
        const response = await apiClient.getPerformanceTable({
          startDate: dateFilter.startDate,
          endDate: dateFilter.endDate,
          formId: id
        });
        if (response.success) {
          setPerformanceTableData(response.data || []);
        }
      } catch (error) {
        console.error("Error fetching performance table:", error);
      } finally {
        setPerformanceTableLoading(false);
      }
    };

    fetchPerformanceTable();
  }, [id, dateFilter.startDate, dateFilter.endDate, user?.role]);

  // Load performance scores from API
  useEffect(() => {
    const loadScores = async () => {
      try {
        const response = await apiClient.getPerformanceScores();
        if (response && response.data) {
          setPerformanceScores(response.data);
        }
      } catch (error) {
        console.error('Failed to load performance scores:', error);
        // Fallback to empty scores
        setPerformanceScores({});
      }
    };
    loadScores();
  }, []);

  // Initialize performance score for current user if not exists
  useEffect(() => {
    if (user?._id && !performanceScores[user._id]) {
      setPerformanceScores(prev => ({
        ...prev,
        [user._id]: 100 // Start with 100%
      }));
    }
  }, [user?._id, performanceScores]);



  // Review system for peer evaluation
  const [selectedReviewOptions, setSelectedReviewOptions] = useState<Record<string, string>>(() => {
    try {
      const saved = localStorage.getItem('selectedReviewOptions');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [reviewedBy, setReviewedBy] = useState<Record<string, { id: string, name: string, email: string } | null>>(() => {
    try {
      const saved = localStorage.getItem('reviewedBy');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [reviewSubmitted, setReviewSubmitted] = useState<Record<string, boolean>>(() => {
    try {
      const saved = localStorage.getItem('reviewSubmitted');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [forceUpdate, setForceUpdate] = useState(0);
  const [pendingReviewOption, setPendingReviewOption] = useState<string | null>(null); // for Rejected/Rework question-selection flow

  // Save chatFilters suggestedAnswers to localStorage
  useEffect(() => {
    console.log('suggestedAnswers changed:', Object.keys(chatFilters.suggestedAnswers || {}));
    try {
      const filtersToSave = {
        suggestedAnswers: chatFilters.suggestedAnswers
      };
      localStorage.setItem('chatFilters', JSON.stringify(filtersToSave));
    } catch (error) {
      console.error('Failed to save chatFilters:', error);
    }
  }, [chatFilters.suggestedAnswers]);

  // Save review states to localStorage
  useEffect(() => {
    try {
      localStorage.setItem('selectedReviewOptions', JSON.stringify(selectedReviewOptions));
    } catch (error) {
      console.error('Failed to save selectedReviewOptions:', error);
    }
  }, [selectedReviewOptions]);

  useEffect(() => {
    try {
      localStorage.setItem('reviewedBy', JSON.stringify(reviewedBy));
    } catch (error) {
      console.error('Failed to save reviewedBy:', error);
    }
  }, [reviewedBy]);

  useEffect(() => {
    try {
      localStorage.setItem('reviewSubmitted', JSON.stringify(reviewSubmitted));
    } catch (error) {
      console.error('Failed to save reviewSubmitted:', error);
    }
  }, [reviewSubmitted]);

  const tenantId = useMemo(() => {
    // Get from form
    if (form?.tenantId) {
      return typeof form.tenantId === 'object' ? form.tenantId._id : form.tenantId;
    }
    // Or from user
    if (user?.tenantId) {
      return typeof user.tenantId === 'object' ? user.tenantId._id : user.tenantId;
    }
    return null;
  }, [form, user]);
  const handleReviewSubmit = async (responseId: string, reviewOption: string) => {
    console.log('=== HANDLE REVIEW SUBMIT START ===');
    console.log('responseId:', responseId);
    console.log('reviewOption:', reviewOption);

    // Get reviewerId properly
    let reviewerId = user?._id || user?.id;
    if (!reviewerId) {
      try {
        const token = localStorage.getItem('auth_token');
        if (token) {
          const payload = JSON.parse(atob(token.split('.')[1]));
          reviewerId = payload.userId || payload.id;
        }
      } catch (e) {
        console.error('Failed to parse token:', e);
      }
    }

    if (!reviewerId) {
      showToast("Cannot submit review. User ID not found.", "error");
      return;
    }

    if (!chatResponse) return;

    let submitterId = (chatResponse as any).submittedBy;

    if (!submitterId && chatResponse.createdBy) {
      if (typeof chatResponse.createdBy === 'object') {
        submitterId = (chatResponse.createdBy as any)._id || (chatResponse.createdBy as any).id;
      } else {
        submitterId = chatResponse.createdBy;
      }
    }

    if (!submitterId) {
      showToast("Cannot submit review. Missing submitter information.", "error");
      return;
    }

    // Don't allow self-review
    if (submitterId && reviewerId === submitterId) {
      showToast("You cannot review your own submissions", "error");
      return;
    }

    // Only allow reviews for valid response statuses
    const responseStatus = responseStatuses[responseId];
    const validStatusesForReview = ["Direct Ok", "Rework Accepted", "Accepted", "Pending Review", "Rework Completed"];
    if (!validStatusesForReview.includes(responseStatus)) {
      showToast("This response status cannot be reviewed", "error");
      return;
    }

    try {
      setPendingReviewOption(null);

      const reviewData = {
        responseId,
        reviewerId: reviewerId,
        submitterId: submitterId,
        reviewOption,
        tenantId: tenantId
      };

      console.log('📤 Submitting review with data:', reviewData);

      const result = await apiClient.submitReview(reviewData);

      console.log('📥 Review API response:', result);

      // ✅ Check if successful
      if (result && result.success) {
        console.log('✅ Review submitted successfully!');

        // Clear local state to force refresh from API
        setSelectedReviewOptions(prev => {
          const newState = { ...prev };
          delete newState[responseId];
          return newState;
        });
        setReviewedBy(prev => {
          const newState = { ...prev };
          delete newState[responseId];
          return newState;
        });
        setReviewSubmitted(prev => {
          const newState = { ...prev };
          delete newState[`${reviewerId}-${responseId}`];
          return newState;
        });

        // Refresh chat history to get the new review
        await fetchChatHistory(responseId);
        setForceUpdate(prev => prev + 1);

        showToast(result.message || `Review submitted: ${reviewOption}`, "success");
      } else {
        console.error('❌ Review submission failed:', result);
        showToast(result?.message || "Failed to submit review", "error");
      }

    } catch (error: any) {
      console.error('❌ Review submission error:', error);
      showToast(error.message || "Failed to submit review", "error");
    }
  };

  const [toast, setToast] = useState<{
    message: string;
    type: "success" | "error" | "info";
    id: string;
  } | null>(null);
  const [selectedResponse, setSelectedResponse] = useState<Response | null>(
    null,
  );
  const [selectedFormForModal, setSelectedFormForModal] = useState<Form | null>(
    null,
  );
  const [formLoading, setFormLoading] = useState(false);
  const [comparisonViewMode, setComparisonViewMode] = useState<
    "dashboard" | "responses"
  >("dashboard");
  const [isShareModalOpen, setIsShareModalOpen] = useState(false);
  const [inspectorSummary, setInspectorSummary] = useState<any[]>([]);
  const [expandedInspectorForms, setExpandedInspectorForms] = useState<Set<string>>(new Set());
  const [allInspectors, setAllInspectors] = useState<any[]>([]);
  const [summaryStatuses, setSummaryStatuses] = useState<string[]>([]);
  const [summaryLoading, setSummaryLoading] = useState(false);
  const [chartOrientation, setChartOrientation] = useState<"v" | "h">("v");
  const [timeSeriesView, setTimeSeriesView] = useState<"daily" | "monthly">("daily");
  const [chartSortOrder, setChartSortOrder] = useState<"default" | "percentage">("percentage");

  const activeGlobalFilterCount = useMemo(() => {
    let count = 0;
    Object.values(cascadingFilters).forEach((answers) => {
      if (answers && answers.length > 0) count++;
    });
    if (locationFilter && locationFilter.length > 0) count++;
    if (
      dateFilter.type !== "all" &&
      (dateFilter.startDate || dateFilter.endDate)
    )
      count++;
    Object.values(columnFilters).forEach((values) => {
      if (values && values.length > 0) count++;
    });
    return count;
  }, [cascadingFilters, locationFilter, dateFilter, columnFilters]);

  const complianceLabels = useMemo(() => {
    const defaultLabels = { yes: "Yes", no: "No", na: "N/A" };
    let labels = { ...defaultLabels };

    // Check if the form has any inspection status based questions
    let hasInspectionQuestions = false;

    // Check form title for inspection keywords
    if (
      form?.title?.toLowerCase().includes("inspection") ||
      form?.title?.toLowerCase().includes("chassis") ||
      form?.title?.toLowerCase().includes("pdi") ||
      form?.title?.toLowerCase().includes("rework") ||
      form?.title?.toLowerCase().includes("accepted") ||
      form?.title?.toLowerCase().includes("rejected") ||
      form?.title?.toLowerCase().includes("verified")
    ) {
      hasInspectionQuestions = true;
    }

    if (!hasInspectionQuestions && form?.sections) {
      for (const section of form.sections) {
        // Check section title
        const sectionTitle = section.title?.toLowerCase() || "";
        if (
          sectionTitle.includes("inspection") ||
          sectionTitle.includes("chassis") ||
          sectionTitle.includes("rework") ||
          sectionTitle.includes("accepted") ||
          sectionTitle.includes("rejected") ||
          sectionTitle.includes("verified")
        ) {
          hasInspectionQuestions = true;
          break;
        }

        if (section.questions) {
          for (const question of section.questions) {
            // Check for ChassisWithZone or similar that might not have a specific type but we handle as status objects
            if (
              question.type === "chassisWithZone" ||
              question.type === "chassisWithoutZone" ||
              question.type === "chassis" ||
              question.type === "zone-in" ||
              question.type === "zone-out" ||
              question.text?.toLowerCase().includes("chassis") ||
              question.text?.toLowerCase().includes("inspection") ||
              question.text?.toLowerCase().includes("accepted") ||
              question.text?.toLowerCase().includes("rework") ||
              question.options?.some(opt => {
                const o = String(opt).toLowerCase();
                return o.includes("accepted") || o.includes("rejected") || o.includes("rework") || o.includes("re-rework");
              })
            ) {
              hasInspectionQuestions = true;
              break;
            }
          }
        }
        if (hasInspectionQuestions) break;
      }
    }

    if (hasInspectionQuestions) {
      return { yes: "Accepted", no: "Rejected", na: "Rework" };
    }

    // Check if the responses contain any inspection status objects
    if (!hasInspectionQuestions && responses && responses.length > 0) {
      for (const response of responses) {
        if (response.answers) {
          for (const answer of Object.values(response.answers)) {
            if (
              typeof answer === "object" &&
              answer !== null &&
              (answer as any).status
            ) {
              const status = String((answer as any).status).toLowerCase().trim();
              if (
                [
                  "accepted",
                  "rejected",
                  "rework",
                  "reworked",
                  "verified",
                  "rework completed",
                ].includes(status) ||
                status.includes("re-rework")
              ) {
                hasInspectionQuestions = true;
                break;
              }
            }
          }
        }
        if (hasInspectionQuestions) break;
      }
    }

    if (hasInspectionQuestions) {
      return { yes: "Accepted", no: "Rejected", na: "Rework" };
    }

    if (form?.sections) {
      for (const section of form.sections) {
        if (section.questions) {
          for (const question of section.questions) {
            if (
              question.type === "yesNoNA" &&
              question.options &&
              question.options.length >= 2
            ) {
              const hasCustomLabels =
                question.options[0] !== "Yes" ||
                question.options[1] !== "No" ||
                (question.options[2] && question.options[2] !== "N/A");

              if (hasCustomLabels) {
                return {
                  yes: question.options[0] || "Yes",
                  no: question.options[1] || "No",
                  na: question.options[2] || "N/A",
                };
              }

              if (labels.yes === "Yes") {
                labels.yes = question.options[0] || "Yes";
                labels.no = question.options[1] || "No";
                labels.na = question.options[2] || "N/A";
              }
            }
          }
        }
      }
    }
    return labels;
  }, [form, responses]);

  useEffect(() => {
    const fetchSummary = async () => {
      setSummaryLoading(true);
      try {
        let url = "/analytics/inspector-summary";
        if (id) {
          url += `?formId=${id}`;
        }

        const [summaryRes, hierarchyRes] = await Promise.all([
          apiClient.get<any>(url),
          apiClient.getUsersHierarchy({ role: "Inspector" })
        ]);

        if (summaryRes.data) {
          setInspectorSummary(summaryRes.data.summary || []);
          setSummaryStatuses(summaryRes.data.allStatuses || []);
        }

        if (hierarchyRes.users) {
          setAllInspectors(hierarchyRes.users);
        }
      } catch (error) {
        console.error("Error fetching inspector data:", error);
      } finally {
        setSummaryLoading(false);
      }
    };

    if (user) {
      fetchSummary();
    }
  }, [user, id]);

  const groupedInspectorSummary = useMemo(() => {
    const groups: Record<string, any> = {};
    inspectorSummary.forEach(item => {
      const title = item.formTitle || "N/A";
      if (!groups[title]) {
        groups[title] = {
          formTitle: title,
          tenantName: item.tenantName,
          totalInspection: 0,
          statusCounts: {},
          subItems: []
        };
      }
      groups[title].totalInspection += item.totalInspection;
      Object.entries(item.statusCounts || {}).forEach(([status, count]) => {
        groups[title].statusCounts[status] = (groups[title].statusCounts[status] || 0) + (count as number);
      });
      groups[title].subItems.push(item);
    });
    return Object.values(groups);
  }, [inspectorSummary]);

  useEffect(() => {
    const fetchData = async () => {
      console.log("[ANALYTICS] fetchData called with ID:", id, "isGuest:", isGuest);
      if (!id) return;

      // Guest access check
      if (isGuest) {
        const guestToken = localStorage.getItem("guest_auth_token");
        const guestFormId = localStorage.getItem("guest_form_id");
        const guestExpiresAt = localStorage.getItem("guest_expires_at");

        const isExpired = guestExpiresAt ? new Date() > new Date(guestExpiresAt) : true;

        if (!guestToken || guestFormId !== id || isExpired) {
          // Clear expired or invalid guest session
          localStorage.removeItem("guest_auth_token");
          localStorage.removeItem("guest_email");
          localStorage.removeItem("guest_form_id");
          localStorage.removeItem("guest_expires_at");
          navigate(`/forms/${id}/analytics/login`);
          return;
        }
      }

      try {
        setLoading(true);
        setError(null);

        console.log("[ANALYTICS DEBUG] Fetching form:", id);

        // Fetch form details
        const formData = await apiClient.getForm(id);
        setForm(formData.form);

        console.log("[ANALYTICS DEBUG] Form fetched:", formData.form?.title);

        // Initialize selected sections for responses view - select all by default
        if (formData.form?.sections && formData.form.sections.length > 0) {
          setSelectedResponsesSectionIds(
            formData.form.sections.map((s: Section) => s.id),
          );
        }

        // Fetch responses for this form
        console.log("[ANALYTICS DEBUG] Fetching responses for form:", id);
        const responsesData = await apiClient.getFormResponses(id);
        console.log(
          "[ANALYTICS DEBUG] Responses fetched:",
          responsesData.responses?.length || 0,
        );
        setResponses(responsesData.responses || []);
      } catch (err) {
        console.error("Error fetching analytics data:", err);
        setError(
          err instanceof Error ? err.message : "Failed to load analytics",
        );
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, [id]);

  // Add this useEffect to update selectedQuestion
  useEffect(() => {
    if (!selectedQuestionId || !form?.sections?.[0]) {
      setSelectedQuestion(null);
      return;
    }

    // Find the selected question from the FIRST section only
    const firstSection = form.sections[0];
    const foundQuestion = firstSection.questions?.find(
      (q: any) => q.id === selectedQuestionId,
    );

    console.log("Found question:", foundQuestion); // For debugging
    console.log("Question options:", foundQuestion?.options); // For debugging

    setSelectedQuestion(foundQuestion || null);
  }, [selectedQuestionId, form]);

  const availableLocations = useMemo(() => {
    const locations = new Set<string>();
    responses.forEach((r) => {
      const meta = r.submissionMetadata?.location;
      if (meta) {
        const city = meta.city || "";
        const country = meta.country || "";
        const locationStr =
          city && country ? `${city}, ${country}` : country || "Unknown";
        if (locationStr !== "Unknown") {
          locations.add(locationStr);
        }
      }
    });
    return Array.from(locations).sort();
  }, [responses]);

  const quizQuestions = useMemo(() => {
    if (!form?.sections) return [];
    const allQs: any[] = [];
    form.sections.forEach((section) => {
      if (section.questions) {
        section.questions.forEach((q) => {
          if (q.correctAnswer !== undefined) {
            allQs.push(q);
          }
          if (q.followUpQuestions) {
            q.followUpQuestions.forEach((fq) => {
              if (fq.correctAnswer !== undefined) {
                allQs.push(fq);
              }
            });
          }
        });
      }
    });
    return allQs;
  }, [form]);

  const calculateScores = (response: Response) => {
    let correct = 0;
    let wrong = 0;

    quizQuestions.forEach((q) => {
      const answer = response.answers?.[q.id];
      if (answer !== undefined && answer !== null && answer !== "") {
        const answerStr = Array.isArray(answer)
          ? answer.join(", ").toLowerCase()
          : String(answer).toLowerCase();
        const correctStr = Array.isArray(q.correctAnswer)
          ? q.correctAnswer.join(", ").toLowerCase()
          : String(q.correctAnswer).toLowerCase();

        if (answerStr === correctStr) {
          correct++;
        } else {
          wrong++;
        }
      }
    });

    return { correct, wrong };
  };

  const baseFilteredResponses = useMemo(() => {
    let result = responses;

    // 0. Role-based Filter (LIFTED: Inspectors can now view and review other tenant/user responses)

    // 2. Location Filter
    if (locationFilter.length > 0) {
      result = result.filter((response) => {
        const meta = response.submissionMetadata?.location;
        if (!meta) return false;
        const city = meta.city || "";
        const country = meta.country || "";
        const locationStr =
          city && country ? `${city}, ${country}` : country || "Unknown";
        return locationFilter.includes(locationStr);
      });
    }

    // 3. Cascading Question Filters
    const cascadingFiltersArray = Object.entries(cascadingFilters).filter(
      ([_, answers]) => answers.length > 0,
    );

    if (cascadingFiltersArray.length > 0) {
      result = result.filter((response) => {
        return cascadingFiltersArray.every(([questionId, selectedAnswers]) => {
          const answer = response.answers?.[questionId];
          if (answer === null || answer === undefined) return false;

          if (Array.isArray(answer)) {
            return answer.some((a) => selectedAnswers.includes(String(a)));
          }
          if (typeof answer === "object" && answer.status) {
            return selectedAnswers.includes(String(answer.status));
          }
          return selectedAnswers.includes(String(answer));
        });
      });
    }

    // 4. Column Filters
    const activeColumnFilters = Object.entries(columnFilters).filter(
      ([_, values]) => values && values.length > 0,
    );

    if (activeColumnFilters.length > 0) {
      result = result.filter((response) => {
        return activeColumnFilters.every(([columnId, allowedValues]) => {
          if (!allowedValues) return true;
          const answer = response.answers?.[columnId];
          const val = answer === null || answer === undefined ? "No Response" : String(answer);
          return allowedValues.includes(val);
        });
      });
    }

    return result;
  }, [responses, user, locationFilter, cascadingFilters, columnFilters]);

  const fetchChatHistory = async (responseId: string) => {
    try {
      console.log('[ChatModal] Fetching chat history for response:', responseId);

      // Fetch messages
      const response = await apiClient.get<any[]>(`/messages/response/${responseId}`);
      const messages = Array.isArray(response.data) ? response.data : [];
      setChatMessages(messages);

      // Fetch reviews from API
      try {
        const reviewsResponse = await apiClient.getReviewsForResponse(responseId);
        console.log('[ChatModal] Reviews API response:', reviewsResponse);

        if (reviewsResponse && reviewsResponse.reviews && reviewsResponse.reviews.length > 0) {
          const latestReview = reviewsResponse.reviews[0];
          console.log('[ChatModal] Latest review from API:', latestReview);

          // Update state from API
          setSelectedReviewOptions(prev => ({
            ...prev,
            [responseId]: latestReview.option
          }));

          setReviewedBy(prev => ({
            ...prev,
            [responseId]: latestReview.reviewer ? {
              id: latestReview.reviewer.id,
              name: latestReview.reviewer.name || 'Reviewer',
              email: latestReview.reviewer.email || ''
            } : null
          }));

          // Also update chatResponse
          setChatResponse(prev => prev ? {
            ...prev,
            review: latestReview
          } : null);

          console.log('[ChatModal] Review state updated from API');
        } else {
          console.log('[ChatModal] No reviews found');
        }
      } catch (reviewError) {
        console.error("[ChatModal] Error fetching reviews:", reviewError);
      }
    } catch (err) {
      console.error("[ChatModal] Error fetching chat history:", err);
    }
  };
  useEffect(() => {
    if (showChatModal && chatResponse) {
      fetchChatHistory(chatResponse.id);
    }
  }, [showChatModal, chatResponse]);

  // Auto-open chat modal if responseId is in URL
  useEffect(() => {
    const responseId = searchParams.get('responseId');
    if (responseId && responses.length > 0) {
      const response = responses.find(r => r.id === responseId || r._id === responseId);
      if (response && !showChatModal) {
        setChatResponse(response);
        setShowChatModal(true);
      }
    }
  }, [searchParams, responses, showChatModal]);

  const handleSendMessage = async (messageOverride?: string) => {
    const messageToSend = messageOverride ?? newMessage;
    if (!messageToSend.trim() || !chatResponse) return;

    setIsSendingMessage(true);
    try {
      const questionContexts: any[] = [];
      const selectedQuestionTitles: string[] = [];
      chatFilters.questions.forEach(qid => {
        form?.sections?.forEach(section => {
          const q = section.questions?.find(q => q.id === qid);
          if (q) {
            const rawAnswer = chatResponse.answers?.[qid];
            let filteredAnswer = rawAnswer;

            // Apply category filtering if selections exist
            if (rawAnswer && typeof rawAnswer === 'object' && rawAnswer.categories && chatFilters.selectedCategories[qid]) {
              const selectedNames = chatFilters.selectedCategories[qid];
              if (selectedNames.length > 0) {
                filteredAnswer = {
                  ...rawAnswer,
                  categories: rawAnswer.categories.filter((cat: any) => selectedNames.includes(cat.name))
                };
              }
            }

            // Parse flat array into structured object for chassis-with-zone questions
            let structuredAnswer = filteredAnswer;
            if (q.type === 'chassis-with-zone' || q.type === 'chassis-without-zone') {
              // Always use chatFilters.suggestedAnswers for current values (they have the latest changes including uploads)
              const currentValue = chatFilters.suggestedAnswers?.[qid];
              if (currentValue) {
                // Use the suggestedAnswers which has the latest data
                structuredAnswer = {
                  status: currentValue.status || '',
                  chassisNumber: currentValue.chassisNumber || '',
                  zones: Array.isArray(currentValue.zone) ? currentValue.zone.join(', ') : (currentValue.zone || ''),
                  zonesData: currentValue.zonesData,
                  evidenceUrl: currentValue.evidenceUrl || ''
                };
                // Add categories from zonesData
                if (currentValue.zonesData) {
                  const cats: any[] = [];
                  Object.entries(currentValue.zonesData).forEach(([zoneName, zoneData]: [string, any]) => {
                    if (zoneData?.categories) {
                      zoneData.categories.forEach((cat: any) => {
                        const defectsArr = (cat.defects || []).map((d: any) => ({
                          name: d.name,
                          details: d.details || { remark: '', fileUrl: '' }
                        }));
                        cats.push({ name: cat.name, defects: defectsArr });
                      });
                    }
                  });
                  if (cats.length > 0) structuredAnswer.categories = cats;
                }
              } else if (Array.isArray(filteredAnswer) && filteredAnswer.length >= 7) {
                // Fallback to array parsing
                structuredAnswer = {
                  status: filteredAnswer[1] || '',
                  chassisNumber: filteredAnswer[0] || '',
                  zones: filteredAnswer[2] || '',
                  categories: filteredAnswer[3] ? [{ name: filteredAnswer[3], defects: filteredAnswer[4] ? [{ name: filteredAnswer[4], details: { remark: filteredAnswer[5] || '', fileUrl: filteredAnswer[6] || '' } }] : [] }] : []
                };
              } else if (filteredAnswer?.zonesData) {
                structuredAnswer = filteredAnswer;
              }
            }

            questionContexts.push({
              questionId: qid,
              title: q.text || 'Question',
              answer: structuredAnswer,
              suggestion: chatFilters.suggestedAnswers[qid]
            });
            selectedQuestionTitles.push(q.text || 'Question');
          }
        });
      });

      // Extract email from createdBy object or string
      const createdByObj = chatResponse.createdBy;
      const createdByEmail = (createdByObj && typeof createdByObj === 'object')
        ? ((createdByObj as any).email || (createdByObj as any)._id?.toString())
        : (typeof createdByObj === 'string' ? createdByObj : 'inspector@focus.com');

      console.log("[handleSendMessage] createdBy:", createdByObj);
      console.log("[handleSendMessage] toEmail:", createdByEmail);

      console.log('Sending message with data:', {
        toEmail: createdByEmail,
        message: messageToSend,
        responseId: chatResponse.id,
        formId: id,
        questionIds: chatFilters.questions,
        questionTitles: selectedQuestionTitles,
        questionContexts: questionContexts,
        tenantId: form?.tenantId || (user?.tenantId as any)?._id || user?.tenantId
      });

      await apiClient.post("/messages/send", {
        toEmail: createdByEmail,
        message: messageToSend,
        responseId: chatResponse.id,
        formId: id,
        questionIds: chatFilters.questions,
        questionTitles: selectedQuestionTitles,
        questionContexts: questionContexts,
        tenantId: form?.tenantId || (user?.tenantId as any)?._id || user?.tenantId
      });

      setNewMessage("");
      // Clear selected questions after sending
      setChatFilters(prev => ({
        ...prev,
        questions: [],
        selectedCategories: {},
        suggestedAnswers: {}
      }));
      fetchChatHistory(chatResponse.id);
    } catch (err) {
      console.error("Error sending message:", err);
    } finally {
      setIsSendingMessage(false);
    }
  };


  // Add these state variables for the confirm update popup
  const [showEditConfirmPopup, setShowEditConfirmPopup] = useState(false);
  const [popupDate, setPopupDate] = useState("");
  const [popupIssuanceDetails, setPopupIssuanceDetails] = useState("");
  const [pendingEditAnswers, setPendingEditAnswers] = useState<Record<string, any> | null>(null);



  // New function to handle the actual save after confirmation
  const handleConfirmEditSave = async () => {
    if (!editingResponseId || !pendingEditAnswers) {
      setShowEditConfirmPopup(false);
      return;
    }

    try {
      setIsSaving(true);

      const existingResponse = responses.find(r => r.id === editingResponseId);
      const existingHistory = existingResponse?.answers?.__submissionHistory || [];

      let newHistory;

      if (existingHistory.length > 0) {
        // ✅ EDITING: Update the LAST entry
        newHistory = [...existingHistory];
        const lastIndex = newHistory.length - 1;
        newHistory[lastIndex] = {
          ...newHistory[lastIndex],
          date: popupDate || newHistory[lastIndex].date,
          issuanceDetails: popupIssuanceDetails || newHistory[lastIndex].issuanceDetails,
          isUpdated: true
        };
        console.log('✅ Updated history entry with isUpdated=true:', newHistory[lastIndex]);
      } else {
        // New entry (first time)
        const newEntry = {
          no: 1,
          date: popupDate,
          issuanceDetails: popupIssuanceDetails,
          isUpdated: false
        };
        newHistory = [newEntry];
      }

      // ✅ IMPORTANT: Remove __submissionHistory from pendingEditAnswers
      const { __submissionHistory, ...cleanAnswers } = pendingEditAnswers;

      const updatedAnswers = {
        ...cleanAnswers,        // ← Clean answers without history
        __submissionHistory: newHistory  // ← Add the updated history
      };

      await apiClient.updateResponse(editingResponseId, {
        answers: updatedAnswers,
        status: editFormStatus,
        notes: editFormNotes,
      });

      // Update local state
      setResponses(
        responses.map((r) =>
          r.id === editingResponseId
            ? {
              ...r,
              answers: updatedAnswers,
              status: editFormStatus,
              notes: editFormNotes,
            }
            : r,
        ),
      );

      setEditingResponseId(null);
      setEditFormData({});
      setEditFormStatus("Accepted");
      setEditFormNotes("");
      setShowEditConfirmPopup(false);
      setPendingEditAnswers(null);
      showToast("Response updated successfully with revision history!", "success");
    } catch (err) {
      console.error("Error updating response:", err);
      showToast("Failed to update response. Please try again.", "error");
    } finally {
      setIsSaving(false);
    }
  };
  const filteredResponses = useMemo(() => {
    let result = baseFilteredResponses;

    // 1. Global Date Filter
    if (dateFilter.type !== "all") {
      result = result.filter((response) => {
        const timestamp = getResponseTimestamp(response);
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];

        if (dateFilter.type === "single" && dateFilter.startDate) {
          return responseDate === dateFilter.startDate;
        } else if (
          dateFilter.type === "range" &&
          dateFilter.startDate &&
          dateFilter.endDate
        ) {
          return (
            responseDate >= dateFilter.startDate &&
            responseDate <= dateFilter.endDate
          );
        }
        return true;
      });
    }

    return result;
  }, [baseFilteredResponses, dateFilter]);

  // Find the primary chassis question to identify unique items/vehicles
  const chassisQuestionId = useMemo(() => {
    if (!form?.sections) {
      return null;
    }
    for (const section of form.sections) {
      if (section.questions) {
        for (const q of section.questions) {
          if (
            q.type === "chassis" ||
            q.type === "chassisWithZone" ||
            q.type === "chassisWithoutZone" ||
            q.type === "zone-in" ||
            q.type === "zone-out" ||
            q.text?.toLowerCase().includes("chassis") ||
            q.trackResponseRank === true ||
            q.trackResponseRank === "true" ||
            q.trackResponseQuestion === true ||
            q.trackResponseQuestion === "true"
          ) {
            return q.id;
          }
        }
      }
    }
    return null;
  }, [form]);

  // Calculate sequential status (Direct Ok, Rework 1, Rework 2, etc.)
  const responseStatuses = useMemo(() => {
    if (!baseFilteredResponses.length) {
      return {};
    }

    // Group responses by unique item (e.g., chassis number)
    const itemGroups: Record<string, Response[]> = {};

    // Sort responses by timestamp ascending to determine sequential order
    const sortedResponses = [...baseFilteredResponses].sort((a, b) => {
      const tA = new Date(getResponseTimestamp(a) || 0).getTime();
      const tB = new Date(getResponseTimestamp(b) || 0).getTime();
      return tA - tB;
    });

    sortedResponses.forEach((r) => {
      let itemId = "unknown";
      if (chassisQuestionId) {
        const answer = r.answers[chassisQuestionId];
        if (answer) {
          if (typeof answer === "object") {
            itemId = answer.chassisNumber || JSON.stringify(answer);
          } else {
            itemId = String(answer);
          }
        } else {
          // If tracking question is present but not answered, treat as unique to avoid mixing un-tracked items
          itemId = `untracked-${r.id}`;
        }
      } else {
        // NO Tracking ID means no grouping for reworks - treat each as unique
        itemId = `response-${r.id}`;
      }

      if (!itemGroups[itemId]) {
        itemGroups[itemId] = [];
      }
      itemGroups[itemId].push(r);
    });

    const statuses: Record<string, string> = {};

    Object.entries(itemGroups).forEach(([groupId, group]) => {
      let reworkCount = 0;
      let hasBeenReworked = false;

      group.forEach((r, index) => {
        let isRework = false;
        let isAccepted = false;
        let isRejected = false;

        // Check individual answers for inspection status
        if (r.answers) {
          Object.values(r.answers).forEach((ans) => {
            if (typeof ans === "object" && ans !== null && (ans as any).status) {
              const s = String((ans as any).status).toLowerCase().trim();
              if (
                s === "rework" ||
                s === "reworked" ||
                s.includes("re-rework")
              ) {
                isRework = true;
              } else if (
                s === "accepted" ||
                s === "rework completed" ||
                s === "verified" ||
                s === "yes" ||
                s === "y"
              ) {
                isAccepted = true;
              } else if (s === "rejected" || s === "no" || s === "n") {
                isRejected = true;
              }
            } else if (typeof ans === "string") {
              const s = ans.toLowerCase().trim();
              if (
                s === "rework" ||
                s === "reworked" ||
                s.includes("re-rework")
              ) {
                isRework = true;
              } else if (
                s === "accepted" ||
                s === "rework completed" ||
                s === "verified" ||
                s === "yes" ||
                s === "y"
              ) {
                isAccepted = true;
              } else if (s === "rejected" || s === "no" || s === "n") {
                isRejected = true;
              }
            }
          });
        }

        const rank = chassisQuestionId ? r.responseRanks?.[chassisQuestionId] : null;

        if (isRejected) {
          statuses[r.id] = "Rejected";
        } else if (isRework) {
          if (chassisQuestionId && groupId !== `untracked-${r.id}`) {
            reworkCount++;
            hasBeenReworked = true;
            statuses[r.id] = `Rework ${reworkCount}`;
          } else {
            statuses[r.id] = "Rework";
          }
        } else if (isAccepted) {
          // If rank is 1, it's definitely the first time this item is seen
          // If no rank but index 0, assume it's the first time in the current view
          if (rank === 1 || (index === 0 && !hasBeenReworked)) {
            statuses[r.id] = "Direct Ok";
          } else if ((rank && rank > 1) || hasBeenReworked) {
            statuses[r.id] = "Rework Accepted";
          } else {
            statuses[r.id] = "Accepted";
          }
        } else {
          statuses[r.id] = "-";
        }
      });
    });

    return statuses;
  }, [baseFilteredResponses, chassisQuestionId]);

  const analytics = useMemo(() => {
    const total = filteredResponses.length;
    const pending = filteredResponses.filter(
      (r) => r.status === "pending" || !r.status,
    ).length;
    const verified = filteredResponses.filter(
      (r) => r.status === "verified",
    ).length;
    const rejected = filteredResponses.filter(
      (r) => r.status === "rejected",
    ).length;

    const recentResponses = filteredResponses
      .filter((r) => getResponseTimestamp(r))
      .sort((a, b) => {
        const timestampA = getResponseTimestamp(a);
        const timestampB = getResponseTimestamp(b);
        const dateA = timestampA ? new Date(timestampA).getTime() : 0;
        const dateB = timestampB ? new Date(timestampB).getTime() : 0;
        if (isNaN(dateA) && isNaN(dateB)) return 0;
        if (isNaN(dateA)) return 1;
        if (isNaN(dateB)) return -1;
        return dateB - dateA;
      })
      .slice(0, 5);

    const responseTrend = filteredResponses.reduce(
      (acc: Record<string, number>, response) => {
        const timestamp = getResponseTimestamp(response);
        if (timestamp) {
          const dateObj = new Date(timestamp);
          if (!isNaN(dateObj.getTime())) {
            const date = dateObj.toISOString().split("T")[0];
            acc[date] = (acc[date] || 0) + 1;
          }
        }
        return acc;
      },
      {},
    );

    const last7Days = Array.from({ length: 7 }, (_, i) => {
      const date = new Date();
      date.setDate(date.getDate() - i);
      return date.toISOString().split("T")[0];
    }).reverse();

    const maxCount = Math.max(
      ...last7Days.map((date) => responseTrend[date] || 0),
      1,
    );
    const percentageData = last7Days.map((date) =>
      Math.round(((responseTrend[date] || 0) / maxCount) * 100),
    );

    return {
      total,
      pending,
      verified,
      rejected,
      recentResponses,
      responseTrend,
      last7Days,
      percentageData,
    };
  }, [filteredResponses]);

  const qualityChartResponses = useMemo(() => {
    let result = [...filteredResponses];
    if (dateFilter.startDate || dateFilter.endDate) {
      result = result.filter((response) => {
        const timestamp = getResponseTimestamp(response);
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];
        if (dateFilter.startDate && dateFilter.endDate) {
          return responseDate >= dateFilter.startDate && responseDate <= dateFilter.endDate;
        } else if (dateFilter.startDate) {
          return responseDate >= dateFilter.startDate;
        } else if (dateFilter.endDate) {
          return responseDate <= dateFilter.endDate;
        }
        return true;
      });
    }
    return result;
  }, [filteredResponses, dateFilter.startDate, dateFilter.endDate]);

  const sectionChartResponses = useMemo(() => {
    let result = [...filteredResponses];
    if (dateFilter.startDate || dateFilter.endDate) {
      result = result.filter((response) => {
        const timestamp = getResponseTimestamp(response);
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];
        if (dateFilter.startDate && dateFilter.endDate) {
          return responseDate >= dateFilter.startDate && responseDate <= dateFilter.endDate;
        } else if (dateFilter.startDate) {
          return responseDate >= dateFilter.startDate;
        } else if (dateFilter.endDate) {
          return responseDate <= dateFilter.endDate;
        }
        return true;
      });
    }
    return result;
  }, [filteredResponses, dateFilter.startDate, dateFilter.endDate]);

  const qualitySectionPerformanceStats = useMemo(
    () => computeSectionPerformanceStats(form, qualityChartResponses),
    [form, qualityChartResponses],
  );

  const dashboardSectionPerformanceStats = useMemo(
    () => computeSectionPerformanceStats(form, sectionChartResponses),
    [form, sectionChartResponses],
  );

  const sectionPerformanceStats = useMemo(
    () => computeSectionPerformanceStats(form, filteredResponses),
    [form, filteredResponses],
  );

  const filteredSectionStats = useMemo(
    () =>
      sectionPerformanceStats.filter(
        (stat) =>
          stat.yes > 0 ||
          stat.no > 0 ||
          stat.na > 0 ||
          (stat.accepted && stat.accepted > 0) ||
          (stat.rejected && stat.rejected > 0) ||
          (stat.rework && stat.rework > 0),
      ),
    [sectionPerformanceStats],
  );

  useEffect(() => {
    const availableIds = filteredSectionStats.map((stat) => stat.id);
    setSelectedSectionIds((prev) => {
      if (!availableIds.length) {
        return [];
      }
      if (!prev.length) {
        return availableIds;
      }
      const next = prev.filter((id) => availableIds.includes(id));
      return next.length ? next : availableIds;
    });
  }, [filteredSectionStats]);

  const visibleSectionStats = useMemo(
    () =>
      filteredSectionStats.filter((stat) =>
        selectedSectionIds.includes(stat.id),
      ),
    [filteredSectionStats, selectedSectionIds],
  );

  const zoneAnalytics = useMemo(
    () => getZoneAnalytics(filteredResponses),
    [filteredResponses],
  );

  const top20Issues = useMemo(() => {
    const allDefects: Array<{
      name: string;
      zone: string;
      category: string;
      reworkCount: number;
      rejectedCount: number;
      total: number;
    }> = [];

    zoneAnalytics.zoneBreakdown.forEach((zone) => {
      zone.categories.forEach((cat) => {
        cat.defects.forEach((defect) => {
          allDefects.push({
            name: defect.name,
            zone: zone.zone,
            category: cat.category,
            reworkCount: defect.reworkCount,
            rejectedCount: defect.rejectedCount,
            total: defect.reworkCount + defect.rejectedCount,
          });
        });
      });
    });

    return allDefects
      .sort((a, b) => b.total - a.total)
      .slice(0, 20);
  }, [zoneAnalytics]);

  const getUniqueColumnValues = (
    questionId: string,
    responses: Response[],
  ): string[] => {
    const values = new Set<string>();
    responses.forEach((response) => {
      const answer = response.answers?.[questionId];
      if (answer !== null && answer !== undefined) {
        if (Array.isArray(answer)) {
          answer.forEach((item) => {
            const strValue = String(item).trim();
            if (strValue) values.add(strValue);
          });
        } else {
          const strValue = String(answer).trim();
          if (strValue) values.add(strValue);
        }
      } else {
        values.add("");
      }
    });
    return Array.from(values).sort((a, b) => {
      if (a === "") return 1;
      if (b === "") return -1;
      return a.localeCompare(b);
    });
  };

  const hasAnswerValue = (value: any) => {
    if (value === null || value === undefined) {
      return false;
    }
    if (typeof value === "string") {
      return value.trim() !== "";
    }
    if (Array.isArray(value)) {
      return value.length > 0;
    }
    if (typeof value === "object") {
      return Object.keys(value).length > 0;
    }
    return true;
  };

  const renderAnswerDisplay = (value: any, question?: any): React.ReactNode => {
    const ensureAbsoluteFileSource = (input: string) => {
      if (!input) {
        return "";
      }
      if (input.startsWith("data:")) {
        return input;
      }
      if (input.startsWith("http://") || input.startsWith("https://")) {
        return input;
      }
      if (input.startsWith("//")) {
        if (typeof window !== "undefined" && window.location) {
          return `${window.location.protocol}${input}`;
        }
        return `https:${input}`;
      }
      const normalized = input.startsWith("/") ? input : `/${input}`;
      if (typeof window !== "undefined" && window.location) {
        return `${window.location.origin}${normalized}`;
      }
      return normalized;
    };

    const extractFileName = (input: string | undefined) => {
      if (!input) {
        return undefined;
      }
      try {
        const sanitized = input.split("?")[0];
        const parts = sanitized.split("/");
        const name = parts[parts.length - 1] || undefined;
        return name ? decodeURIComponent(name) : undefined;
      } catch {
        return undefined;
      }
    };

    const resolveFileData = (input: any) => {
      if (!input) {
        return null;
      }
      const candidate =
        Array.isArray(input) && input.length === 1 ? input[0] : input;
      if (typeof candidate === "string") {
        if (candidate.startsWith("data:")) {
          return {
            data: candidate,
            fileName: question?.fileName || question?.name,
          };
        }
        if (
          candidate.startsWith("http") ||
          candidate.startsWith("//") ||
          candidate.startsWith("/") ||
          candidate.startsWith("uploads/")
        ) {
          const absolute = ensureAbsoluteFileSource(candidate);
          return {
            url: absolute,
            fileName:
              question?.fileName ||
              question?.name ||
              extractFileName(candidate),
          };
        }
        return null;
      }
      if (typeof candidate === "object") {
        const dataValue =
          candidate.data ||
          candidate.value ||
          candidate.file ||
          candidate.base64 ||
          candidate.url ||
          candidate.answer ||
          candidate.path;
        const nameValue =
          candidate.fileName ||
          candidate.filename ||
          candidate.name ||
          question?.fileName ||
          question?.name;
        if (typeof dataValue === "string" && dataValue.startsWith("data:")) {
          return { data: dataValue, fileName: nameValue };
        }
        if (typeof dataValue === "string") {
          const absolute = ensureAbsoluteFileSource(dataValue);
          return {
            url: absolute,
            fileName: nameValue || extractFileName(dataValue),
          };
        }
        if (typeof candidate.url === "string") {
          const absolute = ensureAbsoluteFileSource(candidate.url);
          return {
            url: absolute,
            fileName: nameValue || extractFileName(candidate.url),
          };
        }
      }
      return null;
    };

    if (value === null || value === undefined || value === "") {
      return <span className="text-gray-400">No response</span>;
    }

    if (typeof value === "string") {
      if (value.startsWith("data:")) {
        return (
          <FilePreview
            data={value}
            fileName={question?.fileName || question?.name}
          />
        );
      }

      if (isImageUrl(value)) {
        return <ImageLink text={value} />;
      }

      if (
        value.startsWith("http") ||
        value.startsWith("//") ||
        value.startsWith("/") ||
        value.startsWith("uploads/")
      ) {
        const absolute = ensureAbsoluteFileSource(value);
        if (isImageUrl(absolute)) {
          return <ImageLink text={absolute} />;
        }
        return (
          <a
            href={absolute}
            target="_blank"
            rel="noopener noreferrer"
            className="text-blue-600 hover:text-blue-800"
          >
            {value}
          </a>
        );
      }

      const trimmed = value.trim();
      return trimmed ? (
        trimmed
      ) : (
        <span className="text-gray-400">No response</span>
      );
    }

    if (Array.isArray(value)) {
      if (value.length === 0) {
        return <span className="text-gray-400">No response</span>;
      }

      const previews = value
        .map((entry: any, index: number) => {
          const fileData = resolveFileData(entry);
          if (!fileData) {
            if (typeof entry === "string" && isImageUrl(entry)) {
              return <ImageLink key={index} text={entry} />;
            }
            return (
              <span key={index} className="text-sm">
                {String(entry)}
              </span>
            );
          }
          if (isImageUrl(fileData.url || fileData.data || "")) {
            return (
              <ImageLink
                key={index}
                text={fileData.url || fileData.data || ""}
              />
            );
          }
          return (
            <FilePreview
              key={`${question?.id ?? "file-array"}-${index}`}
              data={fileData.data}
              url={fileData.url}
              fileName={fileData.fileName}
            />
          );
        })
        .filter(Boolean);

      if (previews.length) {
        return <div className="flex flex-wrap gap-2">{previews}</div>;
      }
    }

    if (typeof value === "object") {
      const fileData = resolveFileData(value);
      if (fileData?.url || fileData?.data) {
        const finalUrl = fileData.url || fileData.data;
        if (finalUrl && isImageUrl(finalUrl)) {
          return <ImageLink text={finalUrl} />;
        }
        if (fileData.data) {
          return (
            <FilePreview data={fileData.data} fileName={fileData.fileName} />
          );
        }
        if (fileData.url) {
          return (
            <FilePreview url={fileData.url} fileName={fileData.fileName} />
          );
        }
      }

      if (!Object.keys(value).length) {
        return <span className="text-gray-400">No response</span>;
      }

      const isChassisType =
        value.chassisNumber !== undefined ||
        value.status !== undefined ||
        value.zone !== undefined ||
        value.zones !== undefined ||
        value.categories !== undefined;

      if (isChassisType) {
        const parts: {
          label: string;
          value: string;
          zoneColor?: string;
          isImage?: boolean;
        }[] = [];

        // Get color for zone
        const getZoneColor = (zoneName: string): string => {
          const z = zoneName.toLowerCase().trim();
          if (z.includes("zone a") || z === "a") return "blue";
          if (z.includes("zone b") || z === "b") return "green";
          if (z.includes("zone c") || z === "c") return "purple";
          if (z.includes("zone d") || z === "d") return "orange";
          if (z.includes("zone e") || z === "e") return "pink";
          if (z.includes("zone f") || z === "f") return "cyan";
          return "indigo";
        };

        if (
          value.chassisNumber &&
          String(value.chassisNumber).trim() &&
          String(value.chassisNumber).toLowerCase() !== "no response"
        ) {
          parts.push({
            label: "Chassis",
            value: String(value.chassisNumber),
            zoneColor: "blue",
          });
        }
        if (
          value.status &&
          String(value.status).trim() &&
          String(value.status).toLowerCase() !== "no response"
        ) {
          parts.push({
            label: "Status",
            value: String(value.status),
            zoneColor: "red",
          });
        }
        if (
          (value.remark || value.remarks) &&
          String(value.remark || value.remarks).trim() &&
          String(value.remark || value.remarks).toLowerCase() !== "no response"
        ) {
          parts.push({
            label: "Remark",
            value: String(value.remark || value.remarks),
            zoneColor: "amber",
          });
        }
        const zoneRaw = value.zone || value.zones;
        if (zoneRaw) {
          const zoneVal = Array.isArray(zoneRaw)
            ? zoneRaw.join(", ")
            : String(zoneRaw);
          if (zoneVal.trim()) {
            // If multiple zones, use a mixed color
            if (zoneVal.includes(",")) {
              parts.push({
                label: "Zone",
                value: zoneVal,
                zoneColor: "indigo",
              });
            } else {
              parts.push({
                label: "Zone",
                value: zoneVal,
                zoneColor: getZoneColor(zoneVal),
              });
            }
          }
        }

        // Handle zonesData (categories, defects, remarks) - with zone colors
        if (value.zonesData && typeof value.zonesData === "object") {
          const zoneEntries = Object.entries(value.zonesData);
          for (const [zoneName, zoneVal] of zoneEntries) {
            const zoneColor = getZoneColor(zoneName);
            const colorMap: Record<string, string> = {
              blue: "bg-blue-50 dark:bg-blue-900/30 text-blue-800 dark:text-blue-200",
              green:
                "bg-green-50 dark:bg-green-900/30 text-green-800 dark:text-green-200",
              purple:
                "bg-purple-50 dark:bg-purple-900/30 text-purple-800 dark:text-purple-200",
              orange:
                "bg-orange-50 dark:bg-orange-900/30 text-orange-800 dark:text-orange-200",
              pink: "bg-pink-50 dark:bg-pink-900/30 text-pink-800 dark:text-pink-200",
              cyan: "bg-cyan-50 dark:bg-cyan-900/30 text-cyan-800 dark:text-cyan-200",
              red: "bg-red-50 dark:bg-red-900/30 text-red-800 dark:text-red-200",
              amber:
                "bg-amber-50 dark:bg-amber-900/30 text-amber-800 dark:text-amber-200",
              indigo:
                "bg-indigo-50 dark:bg-indigo-900/30 text-indigo-800 dark:text-indigo-200",
            };
            const colorClass = colorMap[zoneColor] || colorMap.indigo;

            // Add zone header
            parts.push({ label: "Zone", value: zoneName, zoneColor });

            const categories = (zoneVal as any)?.categories;
            if (categories && Array.isArray(categories)) {
              for (const cat of categories) {
                const catName =
                  typeof cat === "string"
                    ? cat
                    : cat?.name || cat?.category || cat?.categoryName || "-";
                parts.push({
                  label: "Category",
                  value: String(catName),
                  zoneColor,
                });

                const defects = cat?.defects;
                if (defects && Array.isArray(defects)) {
                  for (const defect of defects) {
                    const defectName =
                      typeof defect === "string"
                        ? defect
                        : defect?.name || defect?.defect || "-";
                    const defectDetails =
                      typeof defect === "object" ? defect?.details || {} : {};
                    const remark =
                      defectDetails?.remark || defectDetails?.remarks || "-";
                    parts.push({
                      label: "Defect",
                      value: String(defectName),
                      zoneColor,
                    });
                    if (
                      remark &&
                      String(remark).trim() &&
                      String(remark).toLowerCase() !== "-"
                    ) {
                      parts.push({
                        label: "Remark",
                        value: String(remark),
                        zoneColor,
                      });
                    }
                    const fileUrl =
                      defectDetails?.fileUrl ||
                      defectDetails?.file ||
                      defect?.fileUrl ||
                      defect?.file ||
                      defect?.imageUrl ||
                      "";
                    if (
                      fileUrl &&
                      String(fileUrl).toLowerCase() !== "no response" &&
                      String(fileUrl).trim()
                    ) {
                      parts.push({
                        label: "Evidence",
                        value: String(fileUrl),
                        zoneColor,
                        isImage: true,
                      });
                    }
                  }
                }
              }
            }
          }
        }

        // Handle categories (direct property) - both object and array formats
        if (value.categories) {
          if (Array.isArray(value.categories)) {
            // ChassisWithoutZone format: array of category objects
            for (const cat of value.categories) {
              const catName = cat?.name || cat?.category || "-";
              if (catName !== "-") {
                parts.push({
                  label: "Category",
                  value: String(catName),
                  zoneColor: "purple",
                });

                const defects = cat?.defects;
                if (defects && Array.isArray(defects)) {
                  for (const defect of defects) {
                    const defectName =
                      typeof defect === "string"
                        ? defect
                        : defect?.name || defect?.defect || "-";
                    const defectDetails =
                      typeof defect === "object" ? defect?.details || {} : {};
                    const remark =
                      defectDetails?.remark || defectDetails?.remarks || "-";
                    parts.push({
                      label: "Defect",
                      value: String(defectName),
                      zoneColor: "purple",
                    });
                    if (
                      remark &&
                      String(remark).trim() &&
                      String(remark).toLowerCase() !== "-"
                    ) {
                      parts.push({
                        label: "Remark",
                        value: String(remark),
                        zoneColor: "purple",
                      });
                    }
                    const fileUrl =
                      defectDetails?.fileUrl ||
                      defectDetails?.file ||
                      defect?.fileUrl ||
                      defect?.file ||
                      defect?.imageUrl ||
                      "";
                    if (
                      fileUrl &&
                      String(fileUrl).toLowerCase() !== "no response" &&
                      String(fileUrl).trim()
                    ) {
                      parts.push({
                        label: "Evidence",
                        value: String(fileUrl),
                        zoneColor: "purple",
                        isImage: true,
                      });
                    }
                  }
                }
              }
            }
          } else if (typeof value.categories === "object") {
            // Object format: key-value pairs
            const catEntries = Object.entries(value.categories);
            for (const [catKey, catVal] of catEntries) {
              parts.push({
                label: String(catKey),
                value: String(catVal),
                zoneColor: "amber",
              });
            }
          }
        }

        // Handle evidenceUrl
        if (
          value.evidenceUrl &&
          String(value.evidenceUrl).toLowerCase() !== "no response" &&
          String(value.evidenceUrl).trim()
        ) {
          parts.push({
            label: "Evidence",
            value: String(value.evidenceUrl),
            zoneColor: "indigo",
            isImage: true,
          });
        }

        if (parts.length > 0) {
          return (
            <div className="flex flex-col gap-2">
              {parts.map((part, idx) => {
                const colorMap: Record<string, string> = {
                  blue: "bg-blue-50 dark:bg-blue-900/30 text-blue-800 dark:text-blue-200",
                  green:
                    "bg-green-50 dark:bg-green-900/30 text-green-800 dark:text-green-200",
                  purple:
                    "bg-purple-50 dark:bg-purple-900/30 text-purple-800 dark:text-purple-200",
                  orange:
                    "bg-orange-50 dark:bg-orange-900/30 text-orange-800 dark:text-orange-200",
                  pink: "bg-pink-50 dark:bg-pink-900/30 text-pink-800 dark:text-pink-200",
                  cyan: "bg-cyan-50 dark:bg-cyan-900/30 text-cyan-800 dark:text-cyan-200",
                  red: "bg-red-50 dark:bg-red-900/30 text-red-800 dark:text-red-200",
                  amber:
                    "bg-amber-50 dark:bg-amber-900/30 text-amber-800 dark:text-amber-200",
                  indigo:
                    "bg-indigo-50 dark:bg-indigo-900/30 text-indigo-800 dark:text-indigo-200",
                };
                const colorClass =
                  colorMap[part.zoneColor || "indigo"] || colorMap.indigo;

                return (
                  <div key={idx} className="flex items-start gap-2">
                    <span
                      className={`px-2 py-1 ${colorClass} text-xs rounded font-medium min-w-[70px]`}
                    >
                      {part.label}
                    </span>
                    {part.isImage ? (
                      <ImageLink text={part.value} />
                    ) : (
                      <span
                        className={`px-2 py-1 ${colorClass} text-xs rounded font-medium`}
                      >
                        {part.value}
                      </span>
                    )}
                  </div>
                );
              })}
            </div>
          );
        }

        return <span className="text-gray-400">No response</span>;
      }

      const entries = Object.entries(value);
      return (
        <div className="flex flex-col gap-2">
          {entries.map(([k, v], i) => (
            <div
              key={i}
              className="flex flex-col gap-0.5 border-l-2 border-gray-100 dark:border-gray-800 pl-2"
            >
              <span className="text-[10px] font-bold opacity-70 uppercase tracking-tighter text-blue-800 dark:text-blue-300">
                {k}
              </span>
              {renderAnswerDisplay(v)}
            </div>
          ))}
        </div>
      );
    }

    return <pre className="text-xs whitespace-pre-wrap">{JSON.stringify(value, null, 2)}</pre>;
  };

  const handleSelectAllSections = () => {
    setSelectedSectionIds(filteredSectionStats.map((stat) => stat.id));
  };

  const toggleSectionSelection = (sectionId: string) => {
    setSelectedSectionIds((prev) => {
      if (prev.includes(sectionId)) {
        if (prev.length === 1) {
          return prev;
        }
        return prev.filter((id) => id !== sectionId);
      }
      return [...prev, sectionId];
    });
  };

  const sectionChartData = useMemo(() => {
    const calculatePercentage = (value: number, total: number) =>
      total ? parseFloat(((value / total) * 100).toFixed(1)) : 0;

    return {
      labels: visibleSectionStats.map((stat) => formatSectionLabel(stat.title)),
      datasets: [
        {
          label: complianceLabels.na,
          data: visibleSectionStats.map((stat) =>
            calculatePercentage(stat.na + (stat.rework || 0), stat.total),
          ),
          backgroundColor: "#93c5fd",
          borderRadius: 4,
          barThickness: 20,
          hoverBorderWidth: 2,
          hoverBorderColor: "#ffffff",
        },
        {
          label: complianceLabels.no,
          data: visibleSectionStats.map((stat) =>
            calculatePercentage(stat.no + (stat.rejected || 0), stat.total),
          ),
          backgroundColor: "#3b82f6",
          borderRadius: 4,
          barThickness: 20,
          hoverBorderWidth: 2,
          hoverBorderColor: "#ffffff",
        },
        {
          label: complianceLabels.yes,
          data: visibleSectionStats.map((stat) =>
            calculatePercentage(stat.yes + (stat.accepted || 0), stat.total),
          ),
          backgroundColor: "#1d4ed8",
          borderRadius: 4,
          barThickness: 20,
          hoverBorderWidth: 2,
          hoverBorderColor: "#ffffff",
        },
      ],
    };
  }, [filteredSectionStats]);

  const sectionChartOptions = useMemo(
    () => ({
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "point",
        intersect: false,
      },
      layout: {
        padding: { top: 16, right: 32, bottom: 16, left: 8 },
      },
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#374151",
            generateLabels: (chart: any) => {
              const labels =
                ChartJS.defaults.plugins.legend.labels.generateLabels(chart);
              labels.forEach((label: any) => {
                label.color = document.documentElement.classList.contains(
                  "dark",
                )
                  ? "#d1d5db"
                  : "#374151";
              });
              return labels;
            },
          },
        },
        tooltip: {
          enabled: true,
          mode: "index",
          intersect: false,
          anchor: "center",
          callbacks: {
            title: (items: any[]) => {
              const index = items?.[0]?.dataIndex;
              console.log("Tooltip title items:", items, "index:", index, "title:", visibleSectionStats[index]?.title);
              if (index === undefined) {
                return "";
              }
              return visibleSectionStats[index]?.title || "";
            },
            label: (context: any) => {
              console.log("Tooltip label context:", context, "raw:", context.raw, "dataset:", context.dataset.label);
              const value = context.raw ?? 0;
              return `${context.dataset.label}: ${value.toFixed(1)}%`;
            },
          },
        },
      },
      scales: {
        x: {
          beginAtZero: true,
          max: 100,
          stacked: true,
          ticks: {
            callback: (value: any) => `${value}%`,
            color: "#374151",
          },
          title: {
            display: true,
            text: "Percentage",
            color: "#374151",
          },
          grid: {
            color: "#e5e7eb",
          },
        },
        y: {
          stacked: true,
          ticks: {
            autoSkip: false,
            color: "#374151",
          },
          title: {
            display: true,
            text: "Sections",
            color: "#374151",
          },
          grid: {
            color: "#e5e7eb",
          },
        },
      },
    }),
    [visibleSectionStats],
  );

  const visibleDashboardSectionStats = useMemo(
    () =>
      dashboardSectionPerformanceStats.filter((stat) =>
        selectedSectionIds.includes(stat.id),
      ),
    [dashboardSectionPerformanceStats, selectedSectionIds],
  );

  const sectionSummaryRows = useMemo(
    () =>
      visibleDashboardSectionStats
        .map((stat) => {
          const rowYesCount = stat.yes + (stat.accepted || 0);
          const rowNoCount = stat.no + (stat.rejected || 0);
          const rowNaCount = stat.na + (stat.rework || 0);

          const yesPercent = stat.total ? (rowYesCount / stat.total) * 100 : 0;
          const noPercent = stat.total ? (rowNoCount / stat.total) * 100 : 0;
          const naPercent = stat.total ? (rowNaCount / stat.total) * 100 : 0;

          return {
            id: stat.id,
            title: stat.title,
            yesPercent,
            yesCount: rowYesCount,
            noPercent,
            noCount: rowNoCount,
            naPercent,
            naCount: rowNaCount,
            total: stat.total,
          };
        })
        // Sort by Yes percentage in descending order
        .sort((a, b) => b.yesPercent - a.yesPercent),
    [visibleDashboardSectionStats],
  );

  const summaryTotals = useMemo(() => {
    return sectionSummaryRows.reduce(
      (acc, row) => ({
        total: acc.total + row.total,
        yesCount: acc.yesCount + (row.yesCount || 0),
        noCount: acc.noCount + (row.noCount || 0),
        naCount: acc.naCount + (row.naCount || 0),
      }),
      {
        total: 0,
        yesCount: 0,
        noCount: 0,
        naCount: 0,
      },
    );
  }, [sectionSummaryRows]);

  const qualitySectionSummaryRows = useMemo(
    () =>
      qualitySectionPerformanceStats
        .map((stat) => {
          const rowYesCount = stat.yes + (stat.accepted || 0);
          const rowNoCount = stat.no + (stat.rejected || 0);
          const rowNaCount = stat.na + (stat.rework || 0);

          const yesPercent = stat.total ? (rowYesCount / stat.total) * 100 : 0;
          const noPercent = stat.total ? (rowNoCount / stat.total) * 100 : 0;
          const naPercent = stat.total ? (rowNaCount / stat.total) * 100 : 0;

          return {
            id: stat.id,
            title: stat.title,
            yesPercent,
            yesCount: rowYesCount,
            noPercent,
            noCount: rowNoCount,
            naPercent,
            naCount: rowNaCount,
            total: stat.total,
          };
        }),
    [qualitySectionPerformanceStats],
  );

  const uniqueInspectors = useMemo(() => {
    // Combine inspectors from responses, inspectorSummary, and allInspectors (Role-based)
    const inspectors = new Set<string>();

    // From Responses
    filteredResponses.forEach((r) => {
      if (r.submittedBy) inspectors.add(r.submittedBy);
    });

    // From Admin/System Summary
    inspectorSummary.forEach((s) => {
      if (s.qcInspector) inspectors.add(s.qcInspector);
    });

    // From Role-based fetch (All users with Inspector role)
    allInspectors.forEach((i) => {
      if (i.name) inspectors.add(i.name);
      if (i.email) inspectors.add(i.email);
    });

    return Array.from(inspectors).sort();
  }, [filteredResponses, inspectorSummary, allInspectors]);

  const inspectionStats = useMemo(() => {
    let accepted = 0;
    let rejected = 0;
    let reworked = 0;
    let reworkCompleted = 0;

    filteredResponses.forEach((response) => {
      // If an inspector is selected, ensure the response matches that inspector
      if (selectedInspectorForTrend !== "Overall") {
        const inspectorName = response.submittedBy;
        if (inspectorName !== selectedInspectorForTrend) {
          return;
        }
      }

      const status = responseStatuses[response.id];
      if (status === "Direct Ok" || status === "Accepted") {
        accepted++;
      } else if (status === "Rework Accepted") {
        reworkCompleted++;
      } else if (status && status.startsWith("Rework")) {
        reworked++;
      } else if (status === "Rejected") {
        rejected++;
      }
    });

    return { accepted, rejected, reworked, reworkCompleted };
  }, [filteredResponses, responseStatuses, selectedInspectorForTrend]);

  const totalPieChartData = useMemo(() => {
    const directOk = inspectionStats.accepted;
    const reworkCompleted = inspectionStats.reworkCompleted;
    const totalNo = inspectionStats.rejected;
    const totalNA = inspectionStats.reworked;

    const total = directOk + reworkCompleted + totalNo + totalNA;

    if (total === 0) {
      return {
        directOk: 0,
        reworkCompleted: 0,
        no: 0,
        na: 0,
        counts: { directOk: 0, reworkCompleted: 0, no: 0, na: 0, total: 0 },
      };
    }

    const directOkPercent = (directOk / total) * 100;
    const reworkCompletedPercent = (reworkCompleted / total) * 100;
    const noPercent = (totalNo / total) * 100;
    const naPercent = (totalNA / total) * 100;

    return {
      directOk: Number(directOkPercent.toFixed(1)),
      reworkCompleted: Number(reworkCompletedPercent.toFixed(1)),
      no: Number(noPercent.toFixed(1)),
      na: Number(naPercent.toFixed(1)),
      counts: {
        directOk: directOk,
        reworkCompleted: reworkCompleted,
        no: totalNo,
        na: totalNA,
        total: total,
      },
    };
  }, [inspectionStats]);

  const questionPerformanceStats = useMemo(
    () => computeQuestionPerformanceStats(form, filteredResponses),
    [form, filteredResponses],
  );

  const defectChartResponses = useMemo(() => {
    let result = [...filteredResponses];

    if (dateFilter.startDate || dateFilter.endDate) {
      result = result.filter((response) => {
        const timestamp = getResponseTimestamp(response);
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];

        if (dateFilter.startDate && dateFilter.endDate) {
          return responseDate >= dateFilter.startDate && responseDate <= dateFilter.endDate;
        } else if (dateFilter.startDate) {
          return responseDate >= dateFilter.startDate;
        } else if (dateFilter.endDate) {
          return responseDate <= dateFilter.endDate;
        }
        return true;
      });
    }

    result.sort((a, b) => {
      const dateA = new Date(getResponseTimestamp(a) || 0).getTime();
      const dateB = new Date(getResponseTimestamp(b) || 0).getTime();
      return dateB - dateA;
    });

    if (!dateFilter.startDate && !dateFilter.endDate) {
      return result.slice(0, 20);
    }

    return result;
  }, [filteredResponses, dateFilter.startDate, dateFilter.endDate]);

  const trendChartResponses = useMemo(() => {
    let result = [...filteredResponses];

    if (dateFilter.startDate || dateFilter.endDate) {
      result = result.filter((response) => {
        const timestamp = getResponseTimestamp(response);
        if (!timestamp) return false;
        const responseDate = new Date(timestamp).toISOString().split("T")[0];

        if (dateFilter.startDate && dateFilter.endDate) {
          return responseDate >= dateFilter.startDate && responseDate <= dateFilter.endDate;
        } else if (dateFilter.startDate) {
          return responseDate >= dateFilter.startDate;
        } else if (dateFilter.endDate) {
          return responseDate <= dateFilter.endDate;
        }
        return true;
      });
    }

    result.sort((a, b) => {
      const dateA = new Date(getResponseTimestamp(a) || 0).getTime();
      const dateB = new Date(getResponseTimestamp(b) || 0).getTime();
      return dateA - dateB; // Sort ascending for trend
    });

    return result;
  }, [filteredResponses, dateFilter.startDate, dateFilter.endDate]);

  const chartQuestionPerformanceStats = useMemo(() => {
    return computeQuestionPerformanceStats(form, defectChartResponses);
  }, [form, defectChartResponses]);

  const sectionChartHeight = Math.max(320, visibleSectionStats.length * 56);

  const sectionsStats = useMemo(() => {
    if (!form?.sections) return [];

    return form.sections.map((section) => ({
      section,
      stats: getSectionStats(section, responses),
    }));
  }, [form, responses]);

  const filteredSectionsStats = useMemo(() => {
    if (!form?.sections) return [];

    return form.sections.map((section) => ({
      section,
      stats: getSectionStats(section, filteredResponses),
    }));
  }, [form, filteredResponses]);

  const OverallQualityPieChart = () => {
    const data = {
      labels: [
        "Direct Ok",
        "Rework Completed",
        "Rejected",
        "Ongoing Rework",
      ],
      datasets: [
        {
          data: [
            totalPieChartData.directOk,
            totalPieChartData.reworkCompleted,
            totalPieChartData.no,
            totalPieChartData.na,
          ],
          backgroundColor: [
            "rgba(34, 197, 94, 0.85)", // Green for Direct Ok
            "rgba(59, 130, 246, 0.85)", // Blue for Rework Completed
            "rgba(239, 68, 68, 0.85)", // Red for Rejected
            "rgba(234, 179, 8, 0.85)", // Yellow for Ongoing Rework
          ],
          borderColor: [
            "rgb(34, 197, 94)",
            "rgb(59, 130, 246)",
            "rgb(239, 68, 68)",
            "rgb(234, 179, 8)",
          ],
          borderWidth: 2,
          hoverOffset: 15,
        },
      ],
    };

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        datalabels: {
          color: "white",
          font: { weight: "bold" as const, size: 10 },
          formatter: (value: number) => value > 0 ? `${value}%` : "",
        },
        legend: {
          position: "bottom" as const,
          labels: {
            color: document.documentElement.classList.contains("dark")
              ? "#e5e7eb"
              : "#374151",
            font: { size: 9, weight: "bold" as const },
            padding: 8,
            usePointStyle: true,
          },
        },
        tooltip: {
          callbacks: {
            label: function (context: any) {
              const label = context.label || "";
              const value = context.raw || 0;
              const index = context.dataIndex;

              let count = 0;
              if (index === 0) count = totalPieChartData.counts.directOk;
              else if (index === 1) count = totalPieChartData.counts.reworkCompleted;
              else if (index === 2) count = totalPieChartData.counts.no;
              else if (index === 3) count = totalPieChartData.counts.na;

              return `${label}: ${value}% (${count} responses)`;
            },
          },
        },
      },
      cutout: "60%",
      interaction: {
        mode: "nearest" as const,
        intersect: true,
      },
    };

    return (
      <div className="p-4 sm:p-6 bg-gradient-to-br from-white to-gray-50 dark:from-gray-800 dark:to-gray-900 flex flex-col h-full rounded-xl border border-gray-200 dark:border-gray-700">
        <div className="flex flex-col gap-4 mb-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center">
              <div className="p-2 bg-gradient-to-br from-purple-500 to-pink-600 rounded-lg mr-1.5">
                <PieChart className="w-4 h-4 text-white" />
              </div>
              <div>
                <h2 className="text-base sm:text-lg font-bold text-primary-900 dark:text-white">
                  Overall Inspection Trend
                </h2>
                <p className="text-[10px] sm:text-xs text-primary-500 dark:text-primary-400">
                  {selectedInspectorForTrend === "Overall" ? "Accepted/Rejected/Rework Distribution" : `Inspector: ${selectedInspectorForTrend}`}
                </p>
              </div>
            </div>

            {uniqueInspectors.length > 0 && (
              <select
                value={selectedInspectorForTrend}
                onChange={(e) => setSelectedInspectorForTrend(e.target.value)}
                className="text-[10px] font-bold bg-white dark:bg-gray-900 text-primary-700 dark:text-primary-300 border border-gray-200 dark:border-gray-700 rounded-lg px-2 py-1.5 outline-none focus:ring-2 focus:ring-purple-500/50 shadow-sm"
              >
                <option value="Overall">Overall View</option>
                <optgroup label="By Inspector">
                  {uniqueInspectors.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </optgroup>
              </select>
            )}
          </div>
        </div>

        <div className="flex-1 flex flex-col" id="overall-quality-chart">
          {totalPieChartData.counts.total === 0 ? (
            <div className="flex-1 flex items-center justify-center min-h-[200px]">
              <div className="text-center p-4">
                <PieChart className="w-10 h-10 text-gray-300 dark:text-gray-600 mx-auto mb-3" />
                <p className="text-primary-500 dark:text-primary-400 font-medium text-sm">
                  No inspection data available
                </p>
                <p className="text-[10px] text-primary-400 dark:text-primary-500 mt-1">
                  Will appear when inspection responses are recorded
                </p>
              </div>
            </div>
          ) : (
            <>
              <div style={{ height: "200px", position: "relative" }}>
                {/* Only change needed here - use Doughnut instead of Pie */}
                <Doughnut data={data} options={options} />
              </div>

              {/* Stats summary */}
              <div className="mt-4 grid grid-cols-4 gap-1 sm:gap-2">
                {/* Direct Ok */}
                <div className="text-center p-1 bg-green-50/50 dark:bg-green-900/10 rounded-lg">
                  <div className="text-[10px] sm:text-xs font-bold text-green-600 dark:text-green-400">
                    {totalPieChartData.directOk}%
                  </div>
                  <div className="text-[9px] font-medium text-gray-700 dark:text-gray-300 truncate">
                    Direct Ok
                  </div>
                  <div className="text-[8px] text-gray-600 dark:text-gray-500">
                    ({totalPieChartData.counts.directOk})
                  </div>
                </div>

                {/* Rework Completed */}
                <div className="text-center p-1 bg-blue-50/50 dark:bg-blue-900/10 rounded-lg">
                  <div className="text-[10px] sm:text-xs font-bold text-blue-600 dark:text-blue-400">
                    {totalPieChartData.reworkCompleted}%
                  </div>
                  <div className="text-[9px] font-medium text-gray-700 dark:text-gray-300 truncate">
                    Rework Comp
                  </div>
                  <div className="text-[8px] text-gray-600 dark:text-gray-500">
                    ({totalPieChartData.counts.reworkCompleted})
                  </div>
                </div>

                {/* Rejected */}
                <div className="text-center p-1 bg-red-50/50 dark:bg-red-900/10 rounded-lg">
                  <div className="text-[10px] sm:text-xs font-bold text-red-600 dark:text-red-400">
                    {totalPieChartData.no}%
                  </div>
                  <div className="text-[9px] font-medium text-gray-700 dark:text-gray-300 truncate">
                    Rejected
                  </div>
                  <div className="text-[8px] text-gray-600 dark:text-gray-500">
                    ({totalPieChartData.counts.no})
                  </div>
                </div>

                {/* Ongoing Rework */}
                <div className="text-center p-1 bg-amber-50/50 dark:bg-amber-900/10 rounded-lg">
                  <div className="text-[10px] sm:text-xs font-bold text-amber-600 dark:text-amber-400">
                    {totalPieChartData.na}%
                  </div>
                  <div className="text-[9px] font-medium text-gray-700 dark:text-gray-300 truncate">
                    Rework
                  </div>
                  <div className="text-[8px] text-gray-600 dark:text-gray-500">
                    ({totalPieChartData.counts.na})
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    );
  };

  const QuestionStatusDistributionChart = () => {
    // Filter and Sort questions based on issue volume
    const processedQuestions = useMemo(() => {
      let filtered = chartQuestionPerformanceStats.filter((q) => q.rejected > 0 || q.rework > 0);

      if (chartSortOrder === "percentage") {
        filtered = [...filtered].sort((a, b) => {
          const percentA = ((a.rejected + a.rework) / a.total) * 100;
          const percentB = ((b.rejected + b.rework) / b.total) * 100;
          return percentB - percentA;
        });
      }

      return filtered.slice(0, 20);
    }, [chartQuestionPerformanceStats, chartSortOrder]);

    const dateRangeLabel = useMemo(() => {
      if (defectChartResponses.length === 0) return "";
      const timestamps = defectChartResponses.map(r => new Date(getResponseTimestamp(r) || 0).getTime());
      const minDate = new Date(Math.min(...timestamps));
      const maxDate = new Date(Math.max(...timestamps));

      const format = (d: Date) => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

      if (format(minDate) === format(maxDate)) return format(minDate);
      return `${format(minDate)} - ${format(maxDate)}`;
    }, [defectChartResponses]);

    if (!processedQuestions.length && !dateFilter.startDate && !dateFilter.endDate) return null;

    const data = {
      labels: processedQuestions.map((q) =>
        q.text.length > 25 ? q.text.substring(0, 25) + "..." : q.text,
      ),
      datasets: [
        {
          label: complianceLabels.no,
          data: processedQuestions.map((q) => q.rejected),
          backgroundColor: "rgba(153, 27, 27, 0.85)", // Dark Red
          borderColor: "rgb(127, 29, 29)",
          borderWidth: 1,
          barPercentage: processedQuestions.length <= 2 ? 0.3 : 0.7,
          categoryPercentage: 0.8,
          datalabels: {
            color: "#ffffff",
            font: { weight: "bold" as const, size: 10 },
            formatter: (value: number) => value > 0 ? value : "",
            textAlign: "center" as const,
          },
        },
        {
          label: complianceLabels.na,
          data: processedQuestions.map((q) => q.rework),
          backgroundColor: "rgba(55, 65, 81, 0.85)", // Dark Gray
          borderColor: "rgb(31, 41, 55)",
          borderWidth: 1,
          barPercentage: processedQuestions.length <= 2 ? 0.3 : 0.7,
          categoryPercentage: 0.8,
          datalabels: {
            color: "#ffffff",
            font: { weight: "bold" as const, size: 10 },
            formatter: (value: number) => value > 0 ? value : "",
            textAlign: "center" as const,
          },
        },
      ],
    };

    const options = {
      indexAxis: chartOrientation === "h" ? "y" as const : "x" as const,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom" as const,
          labels: {
            color: document.documentElement.classList.contains("dark")
              ? "#e5e7eb"
              : "#374151",
            font: { size: 11, weight: "bold" as const },
            padding: 20,
            usePointStyle: true,
          },
        },
        datalabels: {
          display: (context: any) => {
            return context.dataset.data[context.dataIndex] > 0;
          },
        },
        tooltip: {
          backgroundColor: "rgba(0, 0, 0, 0.8)",
          padding: 12,
          cornerRadius: 8,
          callbacks: {
            label: function (context: any) {
              const value = context.raw;
              return `${context.dataset.label}: ${value}`;
            },
          },
        },
      },
      scales: {
        x: {
          stacked: true,
          ticks: {
            color: document.documentElement.classList.contains("dark")
              ? "#e5e7eb"
              : "#374151",
            font: { size: 10, weight: "600" as const },
            maxRotation: chartOrientation === "v" ? 45 : 0,
            minRotation: 0,
          },
          grid: {
            display: false,
          },
        },
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: {
            stepSize: 1,
            color: document.documentElement.classList.contains("dark")
              ? "#9ca3af"
              : "#6b7280",
            font: { size: 10 },
          },
          grid: {
            color: document.documentElement.classList.contains("dark")
              ? "rgba(255, 255, 255, 0.05)"
              : "rgba(0, 0, 0, 0.03)",
          },
        },
      },
      interaction: {
        mode: "nearest" as const,
        intersect: true,

      },
    };

    const containerStyle = chartOrientation === "h"
      ? { height: `${Math.max(450, processedQuestions.length * 40)}px`, position: "relative" as const }
      : { height: "450px", position: "relative" as const };

    return (
      <div id="defect-distribution-chart" className="p-4 sm:p-6 bg-gradient-to-br from-white to-slate-50 dark:from-gray-800 dark:to-gray-900 flex flex-col h-full rounded-xl border border-gray-200 dark:border-gray-700 shadow-lg hover:shadow-xl transition-shadow">
        <div data-pdf-hide="true" className="flex flex-col xl:flex-row xl:items-center justify-between gap-4 mb-6">
          <div className="flex items-center">
            <div className="p-2 bg-gradient-to-br from-red-600 to-slate-700 rounded-lg mr-2">
              <BarChart3 className="w-4 h-4 text-white" />
            </div>
            <div>
              <h3 className="text-base sm:text-lg font-bold text-slate-900 dark:text-white uppercase tracking-tight">
                Defect Distribution
              </h3>
              <p className="text-[10px] sm:text-xs text-slate-500 dark:text-slate-400">
                {complianceLabels.no} & {complianceLabels.na} volume ({dateRangeLabel})
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 sm:gap-3">
            {/* Sort Toggle */}
            <div className="flex items-center bg-slate-100 dark:bg-gray-700 p-1 rounded-lg">
              <button
                onClick={() => setChartSortOrder("default")}
                className={`px-2 py-1 text-[9px] sm:text-[10px] font-bold rounded transition-all ${chartSortOrder === "default"
                  ? "bg-white dark:bg-gray-600 text-blue-600 shadow-sm"
                  : "text-slate-500"
                  }`}
              >
                DEFAULT
              </button>
              <button
                onClick={() => setChartSortOrder("percentage")}
                className={`px-2 py-1 text-[9px] sm:text-[10px] font-bold rounded transition-all ${chartSortOrder === "percentage"
                  ? "bg-white dark:bg-gray-600 text-blue-600 shadow-sm"
                  : "text-slate-500"
                  }`}
              >
                ISSUE %
              </button>
            </div>

            {/* Orientation Toggle */}
            <div className="flex items-center bg-slate-100 dark:bg-gray-700 p-1 rounded-lg">
              <button
                onClick={() => setChartOrientation("v")}
                title="Vertical View"
                className={`p-1 rounded transition-all ${chartOrientation === "v"
                  ? "bg-white dark:bg-gray-600 text-blue-600 shadow-sm"
                  : "text-slate-500"
                  }`}
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 002 2h2a2 2 0 002-2" />
                </svg>
              </button>
              <button
                onClick={() => setChartOrientation("h")}
                title="Horizontal View"
                className={`p-1 rounded transition-all ${chartOrientation === "h"
                  ? "bg-white dark:bg-gray-600 text-blue-600 shadow-sm"
                  : "text-slate-500"
                  }`}
              >
                <svg className="w-4 h-4 rotate-90" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 002 2h2a2 2 0 002-2" />
                </svg>
              </button>
            </div>

            <div className="h-4 w-px bg-gray-200 dark:bg-gray-700 mx-1 hidden sm:block"></div>

            <button
              onClick={() => setShowFilterModal(true)}
              className={`p-1.5 rounded transition-colors relative ${activeGlobalFilterCount > 0
                ? "text-indigo-600 dark:text-indigo-400 hover:bg-indigo-100 dark:hover:bg-indigo-900/30 bg-indigo-50 dark:bg-indigo-900/20"
                : "text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700"
                }`}
              title="Advanced Filters"
            >
              <Filter className="w-4 h-4" />
              {activeGlobalFilterCount > 0 && (
                <span className="absolute top-0 right-0 flex items-center justify-center w-3.5 h-3.5 text-[8px] font-bold text-white bg-red-500 rounded-full -translate-y-1 translate-x-1">
                  {activeGlobalFilterCount}
                </span>
              )}
            </button>
          </div>
        </div>

        {processedQuestions.length === 0 ? (
          <div className="flex-1 flex flex-col items-center justify-center min-h-[300px] text-center p-8">
            <div className="p-4 bg-slate-50 dark:bg-gray-800/50 rounded-full mb-4">
              <CheckCircle className="w-12 h-12 text-green-500 opacity-50" />
            </div>
            <h4 className="text-slate-900 dark:text-white font-bold mb-1">No Defects Found</h4>
            <p className="text-sm text-slate-500 dark:text-slate-400 max-w-xs">
              No rejected or rework responses were found for your current selection.
            </p>
          </div>
        ) : (
          <div className={chartOrientation === "h" ? "overflow-y-auto" : "w-full"}>
            <div style={containerStyle} id="issue-percentage-chart">
              <Bar data={data} options={options} />
            </div>
          </div>
        )}
      </div>
    );
  };

  const TimeBasedPerformanceGraph = () => {
    const timeData = useMemo(() => {
      if (timeSeriesView === "monthly") {
        return computeMonthlyPerformanceStats(
          trendChartResponses,
          responseStatuses,
          dateFilter.startDate,
          dateFilter.endDate,
        );
      }
      return computeDailyPerformanceStats(
        trendChartResponses,
        responseStatuses,
        dateFilter.startDate,
        dateFilter.endDate,
      );
    }, [trendChartResponses, responseStatuses, timeSeriesView, dateFilter.startDate, dateFilter.endDate]);

    if (timeData.length === 0) return null;

    const data = {
      labels: timeData.map((s) => s.date),
      datasets: [
        {
          label: "Total Responses",
          data: timeData.map((s) => s.totalResponses),
          borderColor: "rgb(55, 65, 81)", // Dark Gray
          backgroundColor: "rgba(55, 65, 81, 0.1)",
          borderWidth: 3,
          tension: 0.3,
          pointRadius: 4,
          pointBackgroundColor: "rgb(55, 65, 81)",
          fill: true,
          datalabels: {
            color: darkMode ? "#e5e7eb" : "#374151",
            align: "top" as const,
            offset: 4,
            font: { weight: "bold" as const, size: 10 },
            formatter: (value: number) => (value > 0 ? value : ""),
          },
        },
        {
          label: "Rework Received",
          data: timeData.map((s) => s.reworkCount),
          borderColor: "rgb(153, 27, 27)", // Dark Red
          backgroundColor: "rgba(153, 27, 27, 0.1)",
          borderWidth: 3,
          tension: 0.3,
          pointRadius: 4,
          pointBackgroundColor: "rgb(153, 27, 27)",
          fill: true,
          datalabels: {
            color: "rgb(153, 27, 27)",
            align: "bottom" as const,
            offset: 4,
            font: { weight: "bold" as const, size: 10 },
            formatter: (value: number) => (value > 0 ? value : ""),
          },
        },
      ],
    };

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom" as const,
          labels: {
            color: darkMode ? "#e5e7eb" : "#374151",
            font: { size: 11, weight: "bold" as const },
            usePointStyle: true,
            padding: 20,
          },
        },
        tooltip: {
          mode: "nearest" as const,
          intersect: true,
          backgroundColor: darkMode ? "#1f2937" : "#ffffff",
          titleColor: darkMode ? "#ffffff" : "#111827",
          bodyColor: darkMode ? "#d1d5db" : "#374151",
          borderColor: darkMode ? "#374151" : "#e5e7eb",
          borderWidth: 1,
          padding: 12,
          cornerRadius: 8,
        },
        datalabels: {
          display: true,
        },
      },
      scales: {
        x: {
          grid: {
            display: false,
          },
          ticks: {
            color: darkMode ? "#e5e7eb" : "#374151",
            font: { size: 10, weight: "500" as const },
          },
        },
        y: {
          beginAtZero: true,
          ticks: {
            stepSize: 1,
            color: darkMode ? "#9ca3af" : "#6b7280",
            font: { size: 10 },
          },
          grid: {
            color: darkMode ? "rgba(255, 255, 255, 0.05)" : "rgba(0, 0, 0, 0.03)",
          },
        },
      },
      interaction: {
        mode: "nearest" as const,
        intersect: true,
        axis: "x" as const,
      },
    };

    return (
      <div id="performance-trend-chart" className="p-6 bg-gradient-to-br from-white to-slate-50 dark:from-gray-800 dark:to-gray-900 flex flex-col h-full rounded-xl border border-gray-200 dark:border-gray-700 shadow-lg hover:shadow-xl transition-shadow w-full mt-6">
        <div data-pdf-hide="true" className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div className="flex items-center">
            <div className="p-2 bg-gradient-to-br from-slate-700 to-red-600 rounded-lg mr-2">
              <TrendingUp className="w-4 h-4 text-white" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white uppercase tracking-tight">
                Performance Trend
              </h3>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Total responses received vs total rework received over time (Response-wise)
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center bg-slate-100 dark:bg-gray-700 p-1 rounded-lg">
              <button
                onClick={() => setTimeSeriesView("daily")}
                className={`px-3 py-1 text-[10px] font-bold rounded transition-all ${timeSeriesView === "daily"
                  ? "bg-white dark:bg-gray-600 text-blue-600 shadow-sm"
                  : "text-slate-500"
                  }`}
              >
                DAILY
              </button>
              <button
                onClick={() => setTimeSeriesView("monthly")}
                className={`px-3 py-1 text-[10px] font-bold rounded transition-all ${timeSeriesView === "monthly"
                  ? "bg-white dark:bg-gray-600 text-blue-600 shadow-sm"
                  : "text-slate-500"
                  }`}
              >
                MONTHLY
              </button>
            </div>
          </div>
        </div>
        <div style={{ height: "400px", position: "relative" }}>
          <Line data={data} options={options} />
        </div>
      </div>
    );
  };

  const InspectorPerformanceChart = () => {
    const currentUserScore = performanceScores[user?._id || ''] || 100;
    const circumference = 2 * Math.PI * 45; // radius = 45
    const strokeDasharray = circumference;
    const strokeDashoffset = circumference - (currentUserScore / 100) * circumference;

    return (
      <div className="card p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-bold text-gray-900 dark:text-white">
            Performance Score
          </h3>
          <div className="text-2xl font-bold text-indigo-600 dark:text-indigo-400">
            {currentUserScore}%
          </div>
        </div>

        <div className="flex items-center justify-center">
          <div className="relative">
            {/* Background circle */}
            <svg className="w-32 h-32 transform -rotate-90" viewBox="0 0 100 100">
              <circle
                cx="50"
                cy="50"
                r="45"
                stroke="currentColor"
                strokeWidth="8"
                fill="transparent"
                className="text-gray-200 dark:text-gray-700"
              />
              {/* Progress circle */}
              <circle
                cx="50"
                cy="50"
                r="45"
                stroke="currentColor"
                strokeWidth="8"
                fill="transparent"
                strokeDasharray={strokeDasharray}
                strokeDashoffset={strokeDashoffset}
                className={`transition-all duration-1000 ease-out ${currentUserScore >= 80 ? 'text-green-500' :
                  currentUserScore >= 60 ? 'text-yellow-500' :
                    currentUserScore >= 40 ? 'text-orange-500' : 'text-red-500'
                  }`}
                strokeLinecap="round"
              />
            </svg>

            {/* Center text */}
            <div className="absolute inset-0 flex items-center justify-center">
              <div className="text-center">
                <div className="text-2xl font-bold text-gray-900 dark:text-white">
                  {currentUserScore}
                </div>
                <div className="text-xs text-gray-500 dark:text-gray-400">
                  Score
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="mt-4 text-center">
          <p className="text-sm text-gray-600 dark:text-gray-400">
            Your performance score based on peer reviews
          </p>
          <div className="mt-2 flex justify-center gap-4 text-xs">
            <span className="text-green-600">+2% for Accepted</span>
            <span className="text-red-600">-2% for Rejected/Rework</span>
          </div>
        </div>
      </div>
    );
  };


  const DirectAcceptedPerformanceGraph = () => {
    const timeData = useMemo(() => {
      return computeDirectAcceptedDailyStats(
        baseFilteredResponses,
        responseStatuses,
        dateFilter.startDate,
        dateFilter.endDate,
      );
    }, [baseFilteredResponses, responseStatuses, dateFilter.startDate, dateFilter.endDate]);

    if (timeData.length === 0) return null;

    const data = {
      labels: timeData.map((s) => s.date),
      datasets: [
        {
          label: "Direct",
          data: timeData.map((s) => s.directCount),
          backgroundColor: "rgba(34, 197, 94, 0.7)", // Green-500
          borderColor: "rgb(21, 128, 61)", // Green-700
          borderWidth: 1,
          stack: "stack1",
        },
        {
          label: "Rework",
          data: timeData.map((s) => s.reworkCount),
          backgroundColor: "rgba(234, 179, 8, 0.7)", // Yellow-500
          borderColor: "rgb(161, 98, 7)", // Yellow-700
          borderWidth: 1,
          stack: "stack1",
        },
        {
          label: "Rework Completed",
          data: timeData.map((s) => s.reworkCompletedCount),
          backgroundColor: "rgba(59, 130, 246, 0.7)", // Blue-500
          borderColor: "rgb(29, 78, 216)", // Blue-700
          borderWidth: 1,
          stack: "stack1",
        },
        {
          label: "Rejected",
          data: timeData.map((s) => s.rejectedCount),
          backgroundColor: "rgba(239, 68, 68, 0.7)", // Red-500
          borderColor: "rgb(185, 28, 28)", // Red-700
          borderWidth: 1,
          stack: "stack1",
        },
      ],
    };

    const options = {
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: {
          display: true,
          position: "top" as const,
          labels: {
            color: darkMode ? "#e5e7eb" : "#374151",
            font: { size: 10, weight: "bold" as const },
            padding: 10,
            usePointStyle: true,
          },
        },

        tooltip: {
          mode: "nearest" as const,
          intersect: true,
          callbacks: {
            label: (context: any) => {
              const datasetLabel = context.dataset.label;
              const value = context.raw;
              return `${datasetLabel}: ${value}`;
            },
          },
        },

        datalabels: {
          display: (context: any) =>
            context.dataset.data[context.dataIndex] > 0,
          color: "#fff",
          font: { weight: "bold" as const, size: 9 },
          formatter: (value: number) => value,
        },
      },
      scales: {
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: {
            stepSize: 1,
            color: darkMode ? "#9ca3af" : "#6b7280",
          },
          grid: {
            color: darkMode
              ? "rgba(255, 255, 255, 0.05)"
              : "rgba(0, 0, 0, 0.03)",
          },
        },
        x: {
          stacked: true,
          ticks: {
            color: darkMode ? "#9ca3af" : "#6b7280",
          },
          grid: {
            display: false,
          },
        },
      },
    };

    return (
      <div id="inspection-status-distribution-chart" className="p-6 bg-gradient-to-br from-white to-slate-50 dark:from-gray-800 dark:to-gray-900 flex flex-col h-full rounded-xl border border-gray-200 dark:border-gray-700 shadow-lg hover:shadow-xl transition-shadow w-full mt-6">
        <div data-pdf-hide="true" className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div className="flex items-center">
            <div className={`p-2 bg-gradient-to-br from-blue-600 to-indigo-800 rounded-lg mr-2`}>
              <BarChart3 className="w-4 h-4 text-white" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white uppercase tracking-tight">
                Inspection Status Trend
              </h3>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Daily distribution of inspection outcomes (counts)
              </p>
            </div>
          </div>
        </div>
        <div style={{ height: "400px", position: "relative" }}>
          <Bar data={data} options={options} />
        </div>
      </div>
    );
  };

  const InspectionStatusLineChart = () => {
    const timeData = useMemo(() => {
      return computeDailyReworkVolumeStats(
        form,
        baseFilteredResponses,
        dateFilter.startDate,
        dateFilter.endDate,
      );
    }, [form, baseFilteredResponses, dateFilter.startDate, dateFilter.endDate]);

    if (timeData.length === 0) return null;

    const data = {
      labels: timeData.map((s) => s.date),
      datasets: [
        {
          label: "Rework",
          data: timeData.map((s) => s.reworkCount),
          borderColor: "rgb(234, 179, 8)", // Yellow-500
          backgroundColor: "rgba(234, 179, 8, 0.3)",
          borderWidth: 1,
          fill: true,
          tension: 0,
        },
      ],
    };

    const options = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          mode: "nearest" as const,
          intersect: true,
          callbacks: {
            label: (context: any) => {
              const label = context.dataset.label;
              const value = context.raw;
              return `${label}: ${value}`;
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            stepSize: 1,
            color: darkMode ? "#9ca3af" : "#6b7280",
          },
          grid: {
            color: darkMode ? "rgba(255, 255, 255, 0.05)" : "rgba(0, 0, 0, 0.03)",
          },
        },
        x: {
          ticks: {
            color: darkMode ? "#9ca3af" : "#6b7280",
          },
          grid: {
            display: false,
          },
        },
      },
    };

    return (
      <div id="status-trends-rework-chart" className="p-6 bg-gradient-to-br from-white to-slate-50 dark:from-gray-800 dark:to-gray-900 flex flex-col h-full rounded-xl border border-gray-200 dark:border-gray-700 shadow-lg hover:shadow-xl transition-shadow w-full mt-6">
        <div data-pdf-hide="true" className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div className="flex items-center">
            <div className={`p-2 bg-gradient-to-br from-purple-600 to-indigo-800 rounded-lg mr-2`}>
              <TrendingUp className="w-4 h-4 text-white" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white uppercase tracking-tight">
                Inspection Status Trends For REWORK
              </h3>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Daily Rework Volume (Sum of question-level reworks)
              </p>
            </div>
          </div>
        </div>
        <div style={{ height: "400px", position: "relative" }}>
          <Line data={data} options={options} />
        </div>
      </div>
    );
  };

  const getSectionAnalyticsData = (): SectionAnalyticsData[] => {
    if (!form?.sections || !form.sections.length) {
      return [];
    }

    return form.sections
      .map((section) => {
        const stats = getSectionStats(section, responses);
        const qualityBreakdown = getSectionQualityBreakdown(section, responses);
        const overallQuality = calculateOverallQuality(qualityBreakdown);

        // Debug log to see what we're getting
        console.log("Section Data for PDF:", {
          sectionId: section.id,
          sectionTitle: section.title,
          questionsCount: section.questions?.length || 0,
          questions: section.questions?.map((q) => ({
            id: q.id,
            text: q.text,
            type: q.type,
          })),
          statsQuestionsDetail: stats.questionsDetail?.length || 0,
        });

        return {
          sectionId: section.id,
          sectionTitle: section.title,
          description: section.description,
          stats: {
            mainQuestionCount: stats.mainQuestionCount,
            totalFollowUpCount: stats.totalFollowUpCount,
            answeredMainQuestions: stats.answeredMainQuestions,
            answeredFollowUpQuestions: stats.answeredFollowUpQuestions,
            totalAnswered: stats.totalAnswered,
            totalResponses: stats.totalResponses,
            completionRate: stats.completionRate,
            avgResponsesPerQuestion: stats.avgResponsesPerQuestion,
            questionsDetail: stats.questionsDetail || [], // Make sure this is not empty
          },
          qualityBreakdown,
          overallQuality,
        };
      })
      .filter((section) => section.stats.questionsDetail.length > 0); // Only include sections with questions
  };

  const fullAnalyticsData = useMemo(() => {
    return {
      total: analytics.total,
      pending: analytics.pending,
      verified: analytics.verified,
      rejected: analytics.rejected,
      inspectionStats: inspectionStats,
      sectionSummaryRows: sectionSummaryRows,
      totalPieChartData: totalPieChartData,
      sectionAnalyticsData: getSectionAnalyticsData(),
      inspectorSummary: inspectorSummary,
      summaryStatuses: summaryStatuses,
      performanceTableData: performanceTableData,
      defectStartDate: dateFilter.startDate,
      defectEndDate: dateFilter.endDate
    };
  }, [analytics, inspectionStats, sectionSummaryRows, totalPieChartData, inspectorSummary, summaryStatuses, defectStartDate, defectEndDate, form, responses]);

  const showToast = (
    message: string,
    type: "success" | "error" | "info" = "success",
  ) => {
    const id = Date.now().toString();
    setToast({ message, type, id });
    setTimeout(() => {
      setToast(null);
    }, 3000);
  };
  useEffect(() => {
    // Pass the toast function to opsExcelExporter
    setToastFunction(showToast);
  }, [showToast]);

  const handleExportToPDF_OPS = async () => {
    try {
      if (!form) {
        showToast("Form data not available.", "error");
        return;
      }

      if (filteredResponses.length === 0) {
        showToast("No responses to export.", "error");
        return;
      }

      setIsExporting(true);
      showToast("Generating PDF...", "info");

      const sectionMapping = {
        headerSectionId: form.sections?.[0]?.id || "",
        generalInstructionsSectionId: form.sections?.[1]?.id || "",
        pastProblemsSectionId: form.sections?.[2]?.id || "",
        processStepsSectionId: form.sections?.[3]?.id || "",
        associateSignSectionId: form.sections?.[4]?.id || "",
      };

      await exportResponsesToOPSPDF(
        form,
        filteredResponses,
        sectionMapping,
        form.title,
        (message) => showToast(message, "info")
      );

    } catch (error) {
      console.error("Error exporting OPS PDF:", error);
      showToast("Failed to export. Please try again.", "error");
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportOPSToExcel = async () => {
    try {
      if (!form) {
        showToast("Form data not available.", "error");
        return;
      }

      if (filteredResponses.length === 0) {
        showToast("No responses to export.", "error");
        return;
      }

      setIsExporting(true);
      showToast("Generating PDF...", "info");

      const sectionMapping = {
        headerSectionId: form.sections?.[0]?.id || "",
        generalInstructionsSectionId: form.sections?.[1]?.id || "",
        pastProblemsSectionId: form.sections?.[2]?.id || "",
        processStepsSectionId: form.sections?.[3]?.id || "",
        associateSignSectionId: form.sections?.[4]?.id || "",
      };

      await exportResponsesToOPSExcel(
        form,
        filteredResponses,
        sectionMapping,
        form.title,
        (message) => showToast(message, "info")
      );

    } catch (error) {
      console.error("Error exporting OPS PDF:", error);
      showToast("Failed to export. Please try again.", "error");
    } finally {
      setIsExporting(false);
    }
  };
  const handleExportToPDF = async () => {
    try {
      setIsExporting(true);
      showToast("Generating PDF report...", "info");

      const success = await exportDashboardToPDF(
        form?.title || "Form Analytics",
        fullAnalyticsData,
        true
      );

      if (success) {
        showToast("PDF report generated successfully!", "success");
      } else {
        showToast("Failed to generate PDF. Please try again.", "error");
      }
    } catch (error) {
      console.error("Error downloading PDF:", error);
      showToast("Failed to generate PDF. Please try again.", "error");
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportToExcel = () => {
    try {
      const headerRow: any[] = ["Timestamp", "Status", "Chassis Number"];
      const columnInfo: Array<{
        questionId: string;
        isFollowUp: boolean;
        correctAnswer?: any;
      }> = [];

      form?.sections?.forEach((section: Section) => {
        if (selectedResponsesSectionIds.includes(section.id)) {
          section.questions?.forEach((q: any) => {
            const isFollowUp = q.parentId || q.showWhen?.questionId;
            headerRow.push(q.text || "Question");
            columnInfo.push({
              questionId: q.id,
              isFollowUp: !!isFollowUp,
              correctAnswer: q.correctAnswer,
            });
          });
        }
      });

      const wsData: any[][] = [headerRow];

      responses.forEach((response: Response) => {
        const rowData: any[] = [
          getResponseTimestamp(response)
            ? new Date(getResponseTimestamp(response)!).toLocaleString()
            : "-",
          responseStatuses[response.id] || "-",
          response.answers?.chassis_number || "-",
        ];

        columnInfo.forEach(({ questionId }) => {
          const answer = response.answers?.[questionId];
          // For complex objects like chassis, stringify appropriately using JSON.stringify for now
          // or just standard string if it's simpler
          let answerStr = "-";
          if (answer !== undefined && answer !== null) {
            if (typeof answer === "object") {
              // Special handling for objects to make them readable in Excel
              if (answer.status) {
                answerStr = answer.status; // just show the status for inspection fields
              } else {
                answerStr = JSON.stringify(answer);
              }
            } else {
              answerStr = String(answer);
            }
          }
          rowData.push(answerStr);
        });

        wsData.push(rowData);
      });

      // Add Overall Inspection Statistics Summary Rows
      const statsHeaderRow: any[] = [
        "Overall Inspection Statistics",
        "",
        "",
        "",
      ];
      const statsDataRow: any[] = [
        `Total Accepted: ${inspectionStats.accepted}`,
        `Total Rejected: ${inspectionStats.rejected}`,
        `Total Reworked: ${inspectionStats.reworked}`,
        ``,
      ];

      wsData.push([]); // Empty spacing row
      const statsHeaderIdx = wsData.length;
      wsData.push(statsHeaderRow);
      const statsDataIdx = wsData.length;
      wsData.push(statsDataRow);

      const ws = XLSX.utils.aoa_to_sheet(wsData);

      const headerFill = { fgColor: { rgb: "FF4F46E5" } };
      const headerFont = { color: { rgb: "FFFFFFFF" }, bold: true };

      // Style Header Row
      for (let i = 0; i < headerRow.length; i++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c: i });
        ws[cellRef].s = {
          fill: headerFill,
          font: headerFont,
          alignment: {
            horizontal: "center",
            vertical: "center",
            wrapText: true,
          },
          border: {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          },
        };
      }

      // Style Common Answer Row
      for (let i = 0; i < headerRow.length; i++) {
        const cellRef = XLSX.utils.encode_cell({ r: 1, c: i });
        ws[cellRef].s = {
          fill: { fgColor: { rgb: "FFF3F4F6" } }, // Light gray background
          font: { italic: true, bold: i === 0 },
          alignment: {
            horizontal: i === 0 ? "left" : "center",
            vertical: "center",
            wrapText: true,
          },
          border: {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          },
        };
      }

      // Style response rows
      const lastResponseRowIdx = responses.length + 1;
      for (let rowIdx = 1; rowIdx < lastResponseRowIdx; rowIdx++) {
        const response = responses[rowIdx - 1];

        // Style Timestamp column
        const timeCellRef = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
        ws[timeCellRef].s = {
          fill: { fgColor: { rgb: "FFF9FAFB" } },
          font: { bold: false },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          },
        };

        // Style Status column
        const statusCellRef = XLSX.utils.encode_cell({ r: rowIdx, c: 1 });
        const currentStatus = responseStatuses[response.id] || "-";
        let statusBgColor = "FFF9FAFB"; // Default

        if (currentStatus === "Direct Ok" || currentStatus === "Rework Accepted" || currentStatus === "Accepted") {
          statusBgColor = "FFDCFCE7"; // green-100
        } else if (currentStatus.includes("Rework")) {
          statusBgColor = "FFFEF3C7"; // amber-100
        } else if (currentStatus === "Rejected") {
          statusBgColor = "FFFEE2E2"; // red-100
        }

        ws[statusCellRef].s = {
          fill: { fgColor: { rgb: statusBgColor } },
          font: { bold: true },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          },
        };

        // Style Chassis Number column
        const chassisCellRef = XLSX.utils.encode_cell({ r: rowIdx, c: 2 });
        ws[chassisCellRef].s = {
          fill: { fgColor: { rgb: "FFF9FAFB" } },
          font: { bold: false },
          alignment: { horizontal: "left", vertical: "center" },
          border: {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          },
        };

        // Style Question columns
        for (let colIdx = 0; colIdx < columnInfo.length; colIdx++) {
          const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx + 3 });
          const info = columnInfo[colIdx];
          const answer = response.answers?.[info.questionId];

          let bgColor = info.isFollowUp ? "FFE9D5FF" : "FFFFFFFF";

          ws[cellRef].s = {
            fill: { fgColor: { rgb: bgColor } },
            alignment: { vertical: "center", wrapText: true },
            border: {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            },
          };
        }
      }

      // Style Stats Header Row
      for (let i = 0; i < 4; i++) {
        const cellRef = XLSX.utils.encode_cell({ r: statsHeaderIdx, c: i });
        ws[cellRef].s = {
          fill: { fgColor: { rgb: "FF4F46E5" } },
          font: { color: { rgb: "FFFFFFFF" }, bold: true },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "medium" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          },
        };
      }

      // Style Stats Data Row
      for (let i = 0; i < 4; i++) {
        const cellRef = XLSX.utils.encode_cell({ r: statsDataIdx, c: i });
        ws[cellRef].s = {
          fill: { fgColor: { rgb: "FFE0E7FF" } }, // Indigo 100
          font: { bold: true, color: { rgb: "FF3730A3" } }, // Indigo 800
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "medium" },
            right: { style: "thin" },
          },
        };
      }

      ws["!cols"] = [
        { wch: 22 }, // Timestamp
        { wch: 15 }, // Status
        { wch: 18 }, // Chassis Number
        ...columnInfo.map(() => ({ wch: 35 })),
      ];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Responses");
      XLSX.writeFile(
        wb,
        `${form?.title || "responses"}-${new Date().toLocaleDateString('en-CA')}.xlsx`,
      );
      showToast("Excel report generated successfully!", "success");
    } catch (error) {
      console.error("Error exporting to Excel:", error);
      showToast("Failed to export to Excel. Please try again.", "error");
    }
  };

  const handleViewDetails = (response: Response) => {
    const responseId = response._id || response.id;
    console.log("Navigating to response:", responseId);
    navigate(`/responses/${responseId}`);
  };

  const handleOpenModal = async (response: Response) => {
    try {
      const formIdentifier = response.questionId;
      if (!formIdentifier) {
        throw new Error("Missing form identifier for response");
      }
      const formData = await apiClient.getForm(formIdentifier);
      const formDetails = formData.form;
      setSelectedResponse(response);
      setSelectedFormForModal(formDetails);
    } catch (err) {
      console.error("Failed to load form for modal:", err);
      showToast("Failed to load form. Please try again.", "error");
    }
  };

  const handleEditStart = (response: Response) => {
    setEditingResponseId(response.id);
    setEditFormData({ ...response.answers });
    setEditFormStatus(response.status || "Accepted");
    setEditFormNotes(response.notes || "");
  };

  const handleSaveEdit = async () => {
    if (!editingResponseId) return;

    // Show popup instead of saving directly
    setPopupDate(new Date().toLocaleDateString("en-GB"));
    setPopupIssuanceDetails(`Updated response`);
    setPendingEditAnswers(editFormData);
    setShowEditConfirmPopup(true);
  };

  const handleCancelEdit = () => {
    setEditingResponseId(null);
    setEditFormData({});
    setEditFormStatus("Accepted");
    setEditFormNotes("");
  };



  const handleDeleteResponse = async () => {
    if (!deletingResponseId) return;

    try {
      setIsDeleting(true);
      await apiClient.deleteResponse(deletingResponseId);

      setResponses(responses.filter((r) => r.id !== deletingResponseId));

      setShowDeleteConfirm(false);
      setDeletingResponseId(null);
      showToast("Response deleted successfully!", "success");
    } catch (err) {
      console.error("Error deleting response:", err);
      showToast("Failed to delete response. Please try again.", "error");
    } finally {
      setIsDeleting(false);
    }
  };

  const handleBulkDeleteResponses = async () => {
    if (selectedResponseIds.length === 0) return;

    try {
      setIsDeleting(true);

      for (const responseId of selectedResponseIds) {
        await apiClient.deleteResponse(responseId);
      }

      setResponses(
        responses.filter((r) => !selectedResponseIds.includes(r.id)),
      );
      setSelectedResponseIds([]);
      setShowBulkDeleteConfirm(false);
      showToast(
        `${selectedResponseIds.length} response(s) deleted successfully!`,
        "success",
      );
    } catch (err) {
      console.error("Error deleting responses:", err);
      showToast("Failed to delete some responses. Please try again.", "error");
    } finally {
      setIsDeleting(false);
    }
  };



  const toggleFormExpansion = (formTitle: string) => {
    setExpandedInspectorForms(prev => {
      const next = new Set(prev);
      if (next.has(formTitle)) {
        next.delete(formTitle);
      } else {
        next.add(formTitle);
      }
      return next;
    });
  };

  const renderSummaryTable = () => {
    if (summaryLoading) {
      return (
        <div className="text-center py-8">
          <div className="animate-spin rounded-full h-8 w-8 border-4 border-blue-600 border-t-transparent mx-auto mb-4"></div>
          <p className="text-gray-500 text-sm">Loading summary...</p>
        </div>
      );
    }

    if (groupedInspectorSummary.length === 0) {
      return null;
    }

    return (
      <div className="mt-8 pt-8 border-t border-gray-100 dark:border-gray-700">
        <div className="flex items-center gap-4 mb-6">
          <div className="w-2 h-8 bg-blue-600 rounded-full shadow-sm shadow-blue-500/20"></div>
          <div>
            <h3 className="text-xl font-black text-gray-900 dark:text-white leading-none mb-1">
              Inspection Summary
            </h3>
            <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">Real-time inspection data</p>
          </div>
        </div>

        <div className="relative bg-white dark:bg-gray-800 rounded-3xl border border-gray-100 dark:border-gray-700 shadow-xl shadow-gray-200/50 dark:shadow-none overflow-hidden">
          <div className="overflow-x-auto scrollbar-thin scrollbar-thumb-gray-200 dark:scrollbar-thumb-gray-700">
            <table className="w-full text-sm text-left border-collapse">
              <thead className="bg-gray-50/80 dark:bg-gray-700/80 backdrop-blur-md sticky top-0 z-10 text-gray-500 dark:text-gray-400 uppercase text-[10px] font-black tracking-[0.15em]">
                <tr>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap">Date</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap">Shift</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap">QC Inspector</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center">Total</th>
                  {/* Dynamic Status Columns */}
                  {summaryStatuses.map((status) => (
                    <th
                      key={status}
                      className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center"
                    >
                      {status}
                    </th>
                  ))}
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50 dark:divide-gray-700">
                {groupedInspectorSummary.map((group, groupIdx) => {
                  const isExpanded = expandedInspectorForms.has(group.formTitle);
                  // Collect all inspectors for this form
                  const inspectors = Array.from(new Set(group.subItems.map((i: any) => i.qcInspector)));

                  return (
                    <React.Fragment key={groupIdx}>
                      {/* Main Group Row */}
                      <tr
                        className={`transition-colors cursor-pointer ${isExpanded ? 'bg-blue-50/50 dark:bg-blue-900/20' : 'hover:bg-blue-50/30 dark:hover:bg-blue-900/10'}`}
                        onClick={() => toggleFormExpansion(group.formTitle)}
                      >
                        <td className="px-4 sm:px-6 py-5 text-gray-400 whitespace-nowrap text-xs">
                          {isExpanded ? '—' : (
                            group.subItems.length > 1
                              ? `${new Date(Math.min(...group.subItems.map((i: any) => new Date(i.date).getTime()))).toLocaleDateString()} - ...`
                              : new Date(group.subItems[0].date).toLocaleDateString()
                          )}
                        </td>
                        <td className="px-4 sm:px-6 py-5 whitespace-nowrap">
                          {!isExpanded && (
                            <span className="inline-flex items-center px-2.5 py-1 rounded-lg text-xs font-bold bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300">
                              {Array.from(new Set(group.subItems.map((i: any) => i.shift || "N/A"))).join(', ')}
                            </span>
                          )}
                        </td>
                        <td className="px-4 sm:px-6 py-5 whitespace-nowrap">
                          <div className="flex items-center -space-x-2">
                            {inspectors.slice(0, 3).map((inspector: any, i) => (
                              <div
                                key={i}
                                className="w-8 h-8 rounded-full bg-blue-100 dark:bg-blue-900/40 border-2 border-white dark:border-gray-800 flex items-center justify-center text-blue-700 dark:text-blue-300 font-black text-[10px] z-[i]"
                                title={inspector}
                              >
                                {inspector?.split(' ').map((n: string) => n[0]).join('')}
                              </div>
                            ))}
                            {inspectors.length > 3 && (
                              <div className="w-8 h-8 rounded-full bg-gray-100 dark:bg-gray-700 border-2 border-white dark:border-gray-800 flex items-center justify-center text-gray-600 dark:text-gray-400 font-bold text-[10px] z-10">
                                +{inspectors.length - 3}
                              </div>
                            )}
                            {inspectors.length <= 1 && inspectors[0] && (
                              <span className="ml-3 font-bold text-gray-700 dark:text-gray-200 text-xs">{inspectors[0]}</span>
                            )}
                          </div>
                        </td>
                        <td className="px-4 sm:px-6 py-5 text-center">
                          <span className="text-base font-black text-gray-900 dark:text-white tabular-nums">
                            {group.totalInspection}
                          </span>
                        </td>
                        {/* Dynamic Status Cells */}
                        {summaryStatuses.map((status) => {
                          const count = group.statusCounts?.[status] || 0;
                          const isZero = count === 0;
                          return (
                            <td
                              key={status}
                              className={`px-4 sm:px-6 py-5 text-center font-black tabular-nums transition-opacity ${isZero ? 'opacity-20 text-gray-400' :
                                status === 'Direct Ok' || status === 'Rework Accepted' ? 'text-emerald-600 dark:text-emerald-400' :
                                  status.startsWith('Rework') ? 'text-amber-600 dark:text-amber-400' :
                                    status === 'Rejected' ? 'text-rose-600 dark:text-rose-400' : 'text-blue-600 dark:text-blue-400'
                                }`}
                            >
                              {count}
                            </td>
                          );
                        })}
                        <td className="px-4 sm:px-6 py-5 text-center">
                          <button className="p-2 hover:bg-white/50 dark:hover:bg-white/10 rounded-full transition-colors">
                            {isExpanded ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
                          </button>
                        </td>
                      </tr>

                      {/* Sub Items (QC Inspectors for this Form) */}
                      {isExpanded && group.subItems.map((row: any, subIdx: number) => (
                        <tr key={`${groupIdx}-${subIdx}`} className="bg-gray-50/30 dark:bg-gray-900/20 border-l-4 border-l-blue-500">
                          <td className="px-4 sm:px-6 py-4 text-gray-600 dark:text-gray-400 whitespace-nowrap tabular-nums font-medium text-xs italic">
                            {new Date(row.date).toLocaleDateString()}
                          </td>
                          <td className="px-4 sm:px-6 py-4 whitespace-nowrap">
                            <span className="inline-flex items-center px-2 py-0.5 rounded-md text-[10px] font-bold bg-gray-200/50 dark:bg-gray-700/50 text-gray-600 dark:text-gray-400">
                              {row.shift || "N/A"}
                            </span>
                          </td>
                          <td className="px-4 sm:px-6 py-4 whitespace-nowrap pl-10">
                            <div className="flex items-center gap-3">
                              <div className="w-6 h-6 rounded-full bg-blue-50 dark:bg-blue-900/20 flex items-center justify-center text-blue-600 dark:text-blue-400 font-black text-[8px]">
                                {row.qcInspector?.split(' ').map((n: string) => n[0]).join('')}
                              </div>
                              <span className="font-bold text-gray-600 dark:text-gray-300 text-xs">{row.qcInspector}</span>
                            </div>
                          </td>
                          <td className="px-4 sm:px-6 py-4 text-center">
                            <span className="text-sm font-bold text-gray-700 dark:text-gray-200 tabular-nums">
                              {row.totalInspection}
                            </span>
                          </td>
                          {summaryStatuses.map((status) => {
                            const count = row.statusCounts?.[status] || 0;
                            const isZero = count === 0;
                            return (
                              <td
                                key={status}
                                className={`px-4 sm:px-6 py-4 text-center font-bold text-xs tabular-nums ${isZero ? 'opacity-10 text-gray-400' : 'opacity-70'}`}
                              >
                                {count}
                              </td>
                            );
                          })}
                          <td></td>
                        </tr>
                      ))}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const renderPerformanceTable = () => {
    if (user?.role !== 'admin' && user?.role !== 'superadmin') return null;

    if (performanceTableLoading) {
      return (
        <div className="mt-12 text-center py-8">
          <div className="animate-spin rounded-full h-8 w-8 border-4 border-blue-600 border-t-transparent mx-auto mb-4"></div>
          <p className="text-gray-500 text-sm">Loading performance data...</p>
        </div>
      );
    }

    if (performanceTableData.length === 0) return null;

    // Pagination logic
    const totalPerformanceItems = performanceTableData.length;
    const totalPerformancePages = Math.ceil(totalPerformanceItems / performancePageSize);
    const startIndex = (performancePage - 1) * performancePageSize;
    const endIndex = startIndex + performancePageSize;
    const paginatedPerformance = performanceTableData.slice(startIndex, endIndex);

    return (
      <div className="mt-12 border-t border-gray-100 dark:border-gray-700 pt-8">
        <div className="flex items-center gap-4 mb-6">
          <div className="w-2 h-8 bg-purple-600 rounded-full shadow-sm shadow-purple-500/20"></div>
          <div>
            <h3 className="text-xl font-black text-gray-900 dark:text-white leading-none mb-1">
              Performance Table
            </h3>
            <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">Form-specific inspector performance</p>
          </div>
        </div>

        <div className="relative bg-white dark:bg-gray-800 rounded-3xl border border-gray-100 dark:border-gray-700 shadow-xl shadow-gray-200/50 dark:shadow-none overflow-hidden">
          <div className="overflow-x-auto scrollbar-thin scrollbar-thumb-gray-200 dark:scrollbar-thumb-gray-700">
            <table className="w-full text-sm text-left border-collapse">
              <thead className="bg-gray-50/80 dark:bg-gray-700/80 backdrop-blur-md sticky top-0 z-10 text-gray-500 dark:text-gray-400 uppercase text-[10px] font-black tracking-[0.15em]">
                <tr>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap">User Name</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center">Total Submitted</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center text-blue-600">Dispatched</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center">Total Reviewed</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center text-green-600">Accepted</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center text-red-600">Rejected</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center text-orange-600">Reworked</th>
                  <th className="px-4 sm:px-6 py-5 border-b border-gray-100 dark:border-gray-700 whitespace-nowrap text-center">Performance Score</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50 dark:divide-gray-700">
                {paginatedPerformance.map((row, idx) => (
                  <tr key={idx} className="hover:bg-purple-50/30 dark:hover:bg-purple-900/10 transition-colors">
                    <td className="px-4 sm:px-6 py-5 font-bold text-gray-900 dark:text-white whitespace-nowrap">{row.name}</td>
                    <td className="px-4 sm:px-6 py-5 font-black text-center tabular-nums">{row.totalSubmitted}</td>
                    <td className="px-4 sm:px-6 py-5 font-black text-center text-blue-600 dark:text-blue-400 tabular-nums">{row.dispatched || 0}</td>
                    <td className="px-4 sm:px-6 py-5 font-black text-center tabular-nums">{row.totalReviewed}</td>
                    <td className="px-4 sm:px-6 py-5 font-black text-center text-green-600 tabular-nums">{row.accepted}</td>
                    <td className="px-4 sm:px-6 py-5 font-black text-center text-red-600 tabular-nums">{row.rejected}</td>
                    <td className="px-4 sm:px-6 py-5 font-black text-center text-orange-600 tabular-nums">{row.rework}</td>
                    <td className="px-4 sm:px-6 py-5 text-center">
                      <span className={`px-3 py-1 rounded-full text-[10px] font-black tabular-nums shadow-sm ${row.performanceScore >= 80 ? 'bg-green-100 text-green-700 dark:bg-green-900/40 dark:text-green-400' :
                        row.performanceScore >= 50 ? 'bg-orange-100 text-orange-700 dark:bg-orange-900/40 dark:text-orange-400' :
                          'bg-red-100 text-red-700 dark:bg-red-900/40 dark:text-red-400'
                        }`}>
                        {row.performanceScore}%
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {totalPerformancePages > 1 && (
            <div className="flex flex-col sm:flex-row items-center justify-between gap-4 p-6 bg-gray-50/50 dark:bg-gray-900/50 border-t border-gray-50 dark:border-gray-700">
              <div className="flex items-center gap-3">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Show</label>
                <select
                  value={performancePageSize}
                  onChange={(e) => { setPerformancePageSize(Number(e.target.value)); setPerformancePage(1); }}
                  className="px-3 py-1.5 text-xs font-bold border border-gray-200 dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-gray-700 dark:text-gray-200 focus:outline-none focus:ring-2 focus:ring-purple-500/20 transition-all shadow-sm"
                >
                  {[5, 10, 20, 50].map(size => (
                    <option key={size} value={size}>{size}</option>
                  ))}
                </select>
                <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                  {startIndex + 1}-{Math.min(endIndex, totalPerformanceItems)} of {totalPerformanceItems}
                </span>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setPerformancePage(prev => Math.max(1, prev - 1))}
                  disabled={performancePage === 1}
                  className="p-2 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl disabled:opacity-30 transition-all hover:bg-gray-50 dark:hover:bg-gray-700 shadow-sm"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>

                <div className="flex items-center gap-1">
                  {Array.from({ length: totalPerformancePages }, (_, i) => i + 1)
                    .filter(num =>
                      totalPerformancePages <= 5 ||
                      Math.abs(num - performancePage) <= 1 ||
                      num === 1 ||
                      num === totalPerformancePages
                    )
                    .map((pageNum, idx, arr) => (
                      <React.Fragment key={pageNum}>
                        {idx > 0 && arr[idx - 1] !== pageNum - 1 && (
                          <span className="text-gray-300 mx-1">...</span>
                        )}
                        <button
                          onClick={() => setPerformancePage(pageNum)}
                          className={`min-w-[32px] h-8 text-[10px] font-black rounded-xl transition-all ${performancePage === pageNum
                            ? 'bg-purple-600 text-white shadow-lg shadow-purple-500/30'
                            : 'bg-white dark:bg-gray-800 text-gray-500 dark:text-gray-400 border border-gray-200 dark:border-gray-700 hover:bg-gray-50'
                            }`}
                        >
                          {pageNum}
                        </button>
                      </React.Fragment>
                    ))}
                </div>

                <button
                  onClick={() => setPerformancePage(prev => Math.min(totalPerformancePages, prev + 1))}
                  disabled={performancePage === totalPerformancePages}
                  className="p-2 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl disabled:opacity-30 transition-all hover:bg-gray-50 dark:hover:bg-gray-700 shadow-sm"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  if (loading) {
    return (
      <div className="p-6">
        <div className="text-center py-12">
          <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary-600 mx-auto"></div>
          <p className="mt-4 text-primary-600">Loading analytics...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-6">
        <div className="text-center py-12">
          <p className="text-red-600">Error loading analytics: {error}</p>
          {!isGuest && (
            <button onClick={() => navigate(-1)} className="mt-4 btn-primary">
              Go Back
            </button>
          )}
          {isGuest && (
            <button onClick={handleLogout} className="mt-4 btn-primary bg-red-600 hover:bg-red-700">
              Log out
            </button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="p-2 sm:p-6 space-y-4 sm:space-y-6 bg-gray-50 dark:bg-gray-950 min-h-screen" id="analytics-scroll-container">
      {/* Header with Tabs - Single Row */}
      {form && (
        <div className="bg-white dark:bg-gray-900 p-3 sm:p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-800 flex flex-col lg:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2 sm:gap-4 w-full lg:w-auto">
            {!isGuest && (
              <button
                onClick={() => navigate(-1)}
                className="p-1.5 sm:p-2 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-800 rounded-lg transition-colors"
                title="Go back"
              >
                <ArrowLeft className="w-4 h-4 sm:w-5 sm:h-5" />
              </button>
            )}
            <h1 className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white truncate max-w-[200px] sm:max-w-md">
              {form?.title || "Form"}
            </h1>
          </div>

          {/* Tabs - Center */}
          <div className="flex items-center gap-1 overflow-x-auto pb-1 lg:pb-0 max-w-full no-scrollbar">

            <>
              <button
                onClick={() => setAnalyticsView("dashboard")}
                className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${analyticsView === "dashboard"
                  ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                  : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                  }`}
              >
                <BarChart3 className="w-4 h-4" />
                Dashboard
              </button>
              <button
                onClick={() => setAnalyticsView("question")}
                className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${analyticsView === "question"
                  ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                  : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                  }`}
              >
                <BarChart3 className="w-4 h-4" />
                Questions
              </button>
              <button
                onClick={() => setAnalyticsView("section")}
                className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${analyticsView === "section"
                  ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                  : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                  }`}
              >
                <FileText className="w-4 h-4" />
                Sections
              </button>
              <button
                onClick={() => setAnalyticsView("opsTable")}
                className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${analyticsView === "opsTable"
                  ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                  : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                  }`}
              >
                <Table className="w-4 h-4" />
                OPS Table
              </button>
              {/* <button
                  onClick={() => setAnalyticsView("table")}
                  className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${
                    analyticsView === "table"
                      ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                      : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                  }`}
                >
                  <Table className="w-4 h-4" />
                  Table
                </button> */}
            </>

            <button
              onClick={() => setAnalyticsView("responses")}
              className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${analyticsView === "responses"
                ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                }`}
            >
              <UsersIcon className="w-4 h-4" />
              Responses
            </button>
            {/* {!isInspector && !isGuest && (
              <button
                onClick={() => setAnalyticsView("comparison")}
                className={`px-3 py-2.5 font-semibold transition-all duration-200 flex items-center gap-2 border-b-2 whitespace-nowrap text-sm ${
                  analyticsView === "comparison"
                    ? "text-blue-600 dark:text-blue-400 border-blue-600 dark:border-blue-400"
                    : "text-gray-600 dark:text-gray-400 border-transparent hover:text-gray-900 dark:hover:text-gray-200"
                }`}
              >
                <UsersIcon className="w-4 h-4" />
                Comparison
              </button>
            )} */}
          </div>

          {/* Right Side - Count and Actions */}
          <div className="flex items-center gap-2 sm:gap-3 whitespace-nowrap w-full lg:w-auto justify-between lg:justify-end">
            <div className="flex items-center gap-1.5 sm:gap-2 px-2 sm:px-3 py-1 sm:py-1.5 bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-blue-900/20 dark:to-indigo-900/20 rounded-lg border border-blue-200 dark:border-blue-800">
              <UsersIcon className="w-3.5 h-3.5 sm:w-4 sm:h-4 text-blue-600 dark:text-blue-400" />
              <div className="text-right">
                <div className="text-sm sm:text-base font-bold text-gray-900 dark:text-white">
                  {analytics.total}
                </div>
              </div>
            </div>
            <div className="flex items-center gap-1.5 sm:gap-2">
              <button
                onClick={() => setShowFilterModal(true)}
                className={`p-1.5 sm:p-2 rounded transition-colors relative ${appliedFilters.length > 0
                  ? "text-indigo-600 dark:text-indigo-400 hover:bg-indigo-100 dark:hover:bg-indigo-900/30 bg-indigo-50 dark:bg-indigo-900/20"
                  : "text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-800"
                  }`}
                title="Advanced Filters"
              >
                <Filter className="w-4 h-4" />
                {appliedFilters.length > 0 && (
                  <span className="absolute top-0 right-0 flex items-center justify-center w-3.5 h-3.5 text-[8px] font-bold text-white bg-red-500 rounded-full -translate-y-1 translate-x-1">
                    {appliedFilters.length}
                  </span>
                )}
              </button>
              {/* Other action buttons - grouped for better spacing */}
              <div className="flex items-center gap-1">
                {!isGuest && (
                  <>
                    <button
                      onClick={handleShareAnalytics}
                      className="p-1.5 sm:p-2 text-gray-500 hover:text-blue-600 dark:text-gray-400 dark:hover:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded-lg transition-colors"
                      title="Share via WhatsApp/Email"
                    >
                      <Share2 className="w-4 h-4" />
                    </button>
                    <button
                      onClick={handleAutoSendSetup}
                      className="p-1.5 sm:p-2 text-gray-500 hover:text-indigo-600 dark:text-gray-400 dark:hover:text-indigo-400 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 rounded-lg transition-colors"
                      title="Email Automation"
                    >
                      <Send className="w-4 h-4" />
                    </button>
                  </>
                )}
                <button
                  onClick={handleExportToPDF}
                  disabled={isExporting}
                  className="p-1.5 sm:p-2 text-gray-500 hover:text-red-600 dark:text-gray-400 dark:hover:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors disabled:opacity-50"
                  title="Export to PDF"
                >
                  {isExporting ? (
                    <Loader2 className="w-4 h-4 animate-spin text-red-500" />
                  ) : (
                    <Download className="w-4 h-4" />
                  )}
                </button>
                <button
                  onClick={handleExportToExcel}
                  disabled={isExporting}
                  className="p-1.5 sm:p-2 text-gray-500 hover:text-green-600 dark:text-gray-400 dark:hover:text-green-400 hover:bg-green-50 dark:hover:bg-green-900/20 rounded-lg transition-colors disabled:opacity-50"
                  title="Export to Excel"
                >
                  <Table className="w-4 h-4" />
                </button>
              </div>
            </div>
            {isGuest && (
              <button
                onClick={handleLogout}
                className="px-3 py-1.5 text-xs font-bold text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors border border-red-200 dark:border-red-800"
              >
                Log out
              </button>
            )}
          </div>
        </div>
      )}

      {/* Dashboard View - Always render for PDF export capability, but hide if not active */}
      {(analyticsView === "dashboard" || isExporting) && (
        <div className={analyticsView === "dashboard" ? "space-y-6" : "absolute -left-[9999px] top-0 w-full opacity-0 pointer-events-none"} aria-hidden={analyticsView !== "dashboard"}>
          <div
            className="w-full"
            id="summary-cards"
          >
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 w-full">
              {/* Response Trend Chart - COMPACT */}
              <div className="p-6 bg-gradient-to-br from-white to-gray-50 dark:from-gray-800 dark:to-gray-900 flex flex-col rounded-xl border border-gray-200 dark:border-gray-700 shadow-lg hover:shadow-xl transition-shadow">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center">
                    <div className="p-2 bg-gradient-to-br from-blue-500 to-indigo-600 rounded-lg mr-2">
                      <BarChart3 className="w-4 h-4 text-white" />
                    </div>
                    <div>
                      <h3 className="text-md font-bold text-primary-900 dark:text-white">
                        Response Trend
                      </h3>
                      <p className="text-xs text-primary-500 dark:text-primary-400">
                        Last 7 days
                      </p>
                    </div>
                  </div>
                </div>

                {Object.keys(analytics.responseTrend).length === 0 ? (
                  <div className="flex-1 flex items-center justify-center min-h-[280px]">
                    <div className="text-center">
                      <div className="mb-2">
                        <BarChart3 className="w-8 h-8 text-gray-300 dark:text-gray-600 mx-auto" />
                      </div>
                      <p className="text-sm text-primary-500 dark:text-primary-400 font-medium">
                        No responses yet
                      </p>
                    </div>
                  </div>
                ) : (
                  <div className="flex-1 flex flex-col">
                    <div style={{ height: "293px", position: "relative" }} id="response-trend-chart">
                      <Line
                        data={{
                          labels: analytics.last7Days.map((date) =>
                            new Date(date).toLocaleDateString("en-US", {
                              month: "short",
                              day: "numeric",
                            }),
                          ),
                          datasets: [
                            {
                              label: "Responses %",
                              data: analytics.percentageData,
                              borderColor: "rgb(59, 130, 246)",
                              backgroundColor: "rgba(59, 130, 246, 0.1)",
                              fill: true,
                              tension: 0.4,
                              pointRadius: 4,
                              pointHoverRadius: 6,
                              pointBackgroundColor: "rgb(59, 130, 246)",
                              pointBorderColor: "#fff",
                              pointBorderWidth: 2,
                              borderWidth: 2,
                            },
                          ],
                        }}
                        options={{
                          responsive: true,
                          maintainAspectRatio: false,
                          interaction: {
                            mode: "index" as const,
                            axis: "x" as const,
                            intersect: false,
                          },
                          plugins: {
                            legend: {
                              display: false,
                            },
                            tooltip: {
                              backgroundColor: "rgba(0, 0, 0, 0.8)",
                              titleColor: "#fff",
                              bodyColor: "#fff",
                              cornerRadius: 6,
                              padding: 10,
                              titleFont: { size: 11, weight: "bold" },
                              bodyFont: { size: 11 },
                              callbacks: {
                                title: (context: any) => {
                                  const index = context[0].dataIndex;
                                  const date = analytics.last7Days[index];
                                  if (!date) return "";
                                  const [y, m, d] = date.split("-").map(Number);
                                  return new Date(y, m - 1, d).toLocaleDateString("en-US", {
                                    month: "short",
                                    day: "numeric",
                                    year: "numeric"
                                  });
                                },
                                label: function (context) {
                                  return `Responses: ${context.parsed.y}`;
                                },
                              },
                            },
                          },
                          scales: {
                            y: {
                              beginAtZero: true,
                              max: 100,
                              grid: {
                                color: "rgba(0, 0, 0, 0.05)",
                                drawBorder: false,
                              },
                              ticks: {
                                color: "rgb(107, 114, 128)",
                                font: { size: 10 },
                                callback: function (value) {
                                  return value + "%";
                                },
                              },
                            },
                            x: {
                              grid: {
                                display: false,
                                drawBorder: false,
                              },
                              ticks: {
                                color: "rgb(107, 114, 128)",
                                font: { size: 10 },
                              },
                            },
                          },
                        }}
                      />
                    </div>
                  </div>
                )}
              </div>

              {/* Pie Chart - COMPACT */}
              <OverallQualityPieChart />
            </div>
          </div>

          {/* Question Distribution Chart */}
          {(questionPerformanceStats.length > 0 || trendChartResponses.length > 0) && (
            <div className="w-full" id="question-distribution-card">
              <div className="w-full space-y-6">
                <InspectionStatusLineChart />
                <QuestionStatusDistributionChart />
                <TimeBasedPerformanceGraph />
                <DirectAcceptedPerformanceGraph />
                {renderSummaryTable()}
                {renderPerformanceTable()}
                <InspectorPerformanceChart />
              </div>
            </div>
          )}
        </div>
      )}

      {form && (
        <>
          {/* Question-wise Analytics */}
          {analyticsView === "question" && (
            <div className="space-y-6">
              <div className="card p-3 sm:p-6">
                <ResponseQuestion
                  question={form}
                  responses={filteredResponses}
                />
              </div>
            </div>
          )}
          {/* Section-wise Analytics */}
          {analyticsView === "section" && (
            <div className="space-y-6">
              {filteredSectionStats.length > 0 ? (
                <>
                  <div className="card p-3 sm:p-4 space-y-3">
                    {/* Header */}
                    <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-gray-50/50 dark:bg-gray-800/30 p-4 rounded-xl border border-gray-100 dark:border-gray-700/50">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-indigo-100 dark:bg-indigo-900/40 rounded-lg">
                          <PieChart className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                        </div>
                        <div>
                          <h3 className="text-lg font-bold text-gray-900 dark:text-white tracking-tight">
                            Section Summary
                          </h3>
                          <p className="text-xs text-gray-500 dark:text-gray-400">Section-wise performance breakdown</p>
                        </div>
                      </div>

                      <div className="flex flex-wrap items-center gap-3">
                        {/* Section Selection Dropdown */}
                        <div className="relative">
                          <button
                            onClick={() =>
                              setShowSectionSelector(!showSectionSelector)
                            }
                            className="flex items-center gap-2 px-4 py-2 text-xs font-bold bg-white dark:bg-gray-900 text-indigo-600 dark:text-indigo-400 rounded-lg border-2 border-indigo-100 dark:border-indigo-900/50 hover:border-indigo-500 transition-all shadow-sm"
                          >
                            <Filter className="w-3.5 h-3.5" />
                            Sections ({selectedSectionIds.length})
                            <ChevronDown className={`w-3.5 h-3.5 transition-transform ${showSectionSelector ? 'rotate-180' : ''}`} />
                          </button>

                          {showSectionSelector && (
                            <div className="absolute top-full right-0 mt-2 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl shadow-xl z-[60] min-w-[240px] max-h-80 overflow-y-auto animate-in slide-in-from-top-2 duration-200">
                              <div className="sticky top-0 bg-gray-50 dark:bg-gray-900 p-2 border-b border-gray-100 dark:border-gray-800 z-10">
                                <label className="flex items-center gap-3 px-3 py-2 hover:bg-white dark:hover:bg-gray-800 rounded-lg cursor-pointer transition-colors">
                                  <input
                                    type="checkbox"
                                    checked={
                                      selectedSectionIds.length ===
                                      filteredSectionStats.length &&
                                      filteredSectionStats.length > 0
                                    }
                                    onChange={handleSelectAllSections}
                                    className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                                  />
                                  <span className="text-sm font-bold text-gray-900 dark:text-white">
                                    Select All Sections
                                  </span>
                                </label>
                              </div>

                              <div className="p-1">
                                {filteredSectionStats.map((stat) => {
                                  const selected = selectedSectionIds.includes(
                                    stat.id,
                                  );
                                  return (
                                    <label
                                      key={stat.id}
                                      className={`flex items-center gap-3 px-3 py-2.5 hover:bg-gray-50 dark:hover:bg-gray-700/50 rounded-lg cursor-pointer transition-colors ${selected ? 'bg-indigo-50/30 dark:bg-indigo-900/10' : ''}`}
                                    >
                                      <input
                                        type="checkbox"
                                        checked={selected}
                                        onChange={() =>
                                          toggleSectionSelection(stat.id)
                                        }
                                        className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                                      />
                                      <span className={`text-sm ${selected ? 'font-bold text-indigo-600 dark:text-indigo-400' : 'text-gray-700 dark:text-gray-300'}`}>
                                        {stat.title}
                                      </span>
                                    </label>
                                  );
                                })}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Color Legend */}
                    <div className="flex flex-wrap items-center gap-6 px-4 py-2 border-b border-gray-100 dark:border-gray-800">
                      <div className="flex items-center gap-2">
                        <div className="w-2.5 h-2.5 bg-green-500 rounded-full shadow-[0_0_8px_rgba(34,197,94,0.4)]"></div>
                        <span className="text-[11px] font-bold text-gray-600 dark:text-gray-400 uppercase tracking-wider">
                          {complianceLabels.yes}
                        </span>
                      </div>
                      <div className="flex items-center gap-2">
                        <div className="w-2.5 h-2.5 bg-red-500 rounded-full shadow-[0_0_8px_rgba(239,68,68,0.4)]"></div>
                        <span className="text-[11px] font-bold text-gray-600 dark:text-gray-400 uppercase tracking-wider">
                          {complianceLabels.no}
                        </span>
                      </div>
                      <div className="flex items-center gap-2">
                        <div className="w-2.5 h-2.5 bg-gray-400 rounded-full shadow-[0_0_8px_rgba(156,163,175,0.4)]"></div>
                        <span className="text-[11px] font-bold text-gray-600 dark:text-gray-400 uppercase tracking-wider">
                          {complianceLabels.na}
                        </span>
                      </div>
                    </div>

                    {/* Combined Table with Visualization and Radar Chart */}
                    <div className="flex flex-col lg:flex-row gap-6">
                      {/* Table Container - Always shrinks for radar chart */}
                      <div className="flex-1 min-w-0">
                        <div className="overflow-x-auto no-scrollbar rounded-lg border border-gray-200 dark:border-gray-700">
                          <table className="min-w-full text-xs sm:text-sm border-collapse">
                            <thead className="uppercase tracking-wider text-gray-600 dark:text-gray-300 bg-gray-100 dark:bg-gray-800 z-20">
                              <tr className="bg-gray-200 dark:bg-gray-800">
                                <th rowSpan={2} className="text-left px-4 py-3 border border-gray-300 dark:border-gray-600 min-w-[250px] font-bold">Section Summary</th>
                                <th rowSpan={2} className="text-center px-3 py-3 border border-gray-300 dark:border-gray-600 font-bold">Total</th>
                                <th colSpan={3} className="text-center px-3 py-2 border border-gray-300 dark:border-gray-600 bg-gray-100 dark:bg-gray-700 font-bold">
                                  Section Performance Breakdown
                                </th>
                                <th rowSpan={2} className="text-center px-4 py-3 border border-gray-300 dark:border-gray-600 font-bold">Visualization</th>
                              </tr>
                              <tr className="bg-gray-100 dark:bg-gray-700/50">
                                <th className="text-center px-3 py-2 border border-gray-300 dark:border-gray-600 text-green-700 dark:text-green-400 font-bold">
                                  {complianceLabels.yes}
                                </th>
                                <th className="text-center px-3 py-2 border border-gray-300 dark:border-gray-600 text-red-700 dark:text-red-400 font-bold">
                                  {complianceLabels.no}
                                </th>
                                <th className="text-center px-3 py-2 border border-gray-300 dark:border-gray-600 text-slate-700 dark:text-slate-400 font-bold">
                                  {complianceLabels.na}
                                </th>
                              </tr>
                            </thead>
                            <tbody>
                              {(() => {
                                return (
                                  <>
                                    {sectionSummaryRows.map((row, index) => {
                                      const rowBgColor =
                                        index % 2 === 0
                                          ? "bg-white dark:bg-gray-900"
                                          : "bg-gray-50 dark:bg-gray-800/50";

                                      return (
                                        <tr
                                          key={row.id}
                                          onClick={() => {
                                            setAutoOpenSectionId(null);
                                            setTimeout(
                                              () => setAutoOpenSectionId(row.id),
                                              10,
                                            );
                                          }}
                                          className={`border-b border-gray-300 dark:border-gray-600 hover:bg-gray-100 dark:hover:bg-gray-800 transition-colors cursor-pointer ${rowBgColor}`}
                                        >
                                          {/* Section Column */}
                                          <td className="px-4 py-3 cursor-pointer border border-gray-300 dark:border-gray-600">
                                            <button
                                              onClick={(e) => {
                                                e.stopPropagation();
                                                setAutoOpenSectionId(null);
                                                setTimeout(
                                                  () =>
                                                    setAutoOpenSectionId(row.id),
                                                  10,
                                                );
                                              }}
                                              className="font-bold text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300 text-sm transition-colors text-left"
                                            >
                                              {row.title}
                                            </button>
                                          </td>

                                          {/* Total Column */}
                                          <td className="text-center px-3 py-3 border border-gray-300 dark:border-gray-600">
                                            <div className="font-bold text-gray-900 dark:text-white text-sm">
                                              {row.total}
                                            </div>
                                          </td>

                                          {/* Yes Column */}
                                          <td className="text-center px-3 py-3 border border-gray-300 dark:border-gray-600">
                                            <div className="font-bold text-green-700 dark:text-green-400 text-sm">
                                              {row.yesCount}{" "}
                                              <span className="text-gray-500 dark:text-gray-400 font-medium">
                                                (
                                                {Number.isFinite(row.yesPercent)
                                                  ? row.yesPercent.toFixed(0)
                                                  : "0"}
                                                %)
                                              </span>
                                            </div>
                                          </td>

                                          {/* No Column */}
                                          <td className="text-center px-3 py-3 border-x border-gray-300 dark:border-gray-600">
                                            <div className="font-bold text-red-700 dark:text-red-400 text-sm">
                                              {row.noCount}{" "}
                                              <span className="text-gray-500 dark:text-gray-400 font-medium">
                                                (
                                                {Number.isFinite(row.noPercent)
                                                  ? row.noPercent.toFixed(0)
                                                  : "0"}
                                                %)
                                              </span>
                                            </div>
                                          </td>

                                          {/* N/A Column */}
                                          <td className="text-center px-3 py-3 border-x border-gray-300 dark:border-gray-600">
                                            <div className="font-bold text-slate-700 dark:text-slate-400 text-sm">
                                              {row.naCount}{" "}
                                              <span className="text-gray-500 dark:text-gray-400 font-medium">
                                                (
                                                {Number.isFinite(row.naPercent)
                                                  ? row.naPercent.toFixed(0)
                                                  : "0"}
                                                %)
                                              </span>
                                            </div>
                                          </td>

                                          {/* Visualization Column */}
                                          <td className="px-3 py-3 border-x border-gray-300 dark:border-gray-600">
                                            <div className="flex justify-center">
                                              {generateTableBarChart(
                                                row.yesPercent,
                                                row.noPercent,
                                                row.naPercent,
                                              )}
                                            </div>
                                          </td>
                                        </tr>
                                      );
                                    })}

                                    {/* Comprehensive Total Row */}
                                    <tr className="bg-gray-100 dark:bg-gray-800 font-extrabold border-t-2 border-gray-400 dark:border-gray-500">
                                      <td className="px-4 py-3 text-gray-900 dark:text-gray-100 border-x border-gray-300 dark:border-gray-600">
                                        <div className="flex items-center">
                                          <div className="w-3 h-3 bg-indigo-600 rounded-full mr-3"></div>
                                          <span>TOTAL</span>
                                        </div>
                                      </td>
                                      <td className="text-center px-3 py-3 text-gray-900 dark:text-gray-100 border-x border-gray-300 dark:border-gray-600">
                                        {summaryTotals?.total || 0}
                                      </td>
                                      <td className="text-center px-3 py-3 text-green-700 dark:text-green-400 border-x border-gray-300 dark:border-gray-600">
                                        {summaryTotals?.yesCount || 0} (
                                        {summaryTotals?.total > 0
                                          ? (
                                            (summaryTotals.yesCount /
                                              summaryTotals.total) *
                                            100
                                          ).toFixed(0)
                                          : 0}
                                        %)
                                      </td>
                                      <td className="text-center px-3 py-3 text-red-700 dark:text-red-400 border-x border-gray-300 dark:border-gray-600">
                                        {summaryTotals?.noCount || 0} (
                                        {summaryTotals?.total > 0
                                          ? (
                                            (summaryTotals.noCount /
                                              summaryTotals.total) *
                                            100
                                          ).toFixed(0)
                                          : 0}
                                        %)
                                      </td>
                                      <td className="text-center px-3 py-3 text-slate-700 dark:text-slate-400 border-x border-gray-300 dark:border-gray-600">
                                        {summaryTotals?.naCount || 0} (
                                        {summaryTotals?.total > 0
                                          ? (
                                            (summaryTotals.naCount /
                                              summaryTotals.total) *
                                            100
                                          ).toFixed(0)
                                          : 0}
                                        %)
                                      </td>
                                      <td className="px-3 py-3 border-x border-gray-300 dark:border-gray-600">
                                        <div className="flex justify-center">
                                          {generateTableBarChart(
                                            summaryTotals.total > 0 ? (summaryTotals.yesCount / summaryTotals.total) * 100 : 0,
                                            summaryTotals.total > 0 ? (summaryTotals.noCount / summaryTotals.total) * 100 : 0,
                                            summaryTotals.total > 0 ? (summaryTotals.naCount / summaryTotals.total) * 100 : 0
                                          )}
                                        </div>
                                      </td>
                                    </tr>
                                  </>
                                );
                              })()}
                            </tbody>
                          </table>
                        </div>
                      </div>

                      {/* Radar Chart - Always displayed on right side */}
                      <div className="w-full lg:w-[450px] flex-shrink-0">
                        <div className="bg-white dark:bg-gray-800/40 p-4 sm:p-6 rounded-2xl border border-gray-200 dark:border-gray-700 shadow-md h-full">
                          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-4 sm:mb-6 gap-2">
                            <div>
                              <h4 className="text-base font-bold text-gray-900 dark:text-white uppercase tracking-tight">
                                Performance Radar
                              </h4>
                              <p className="text-[10px] text-gray-500 font-medium">Comparative section analysis</p>
                            </div>
                            <div className="flex flex-wrap items-center gap-2">
                              <div className="flex items-center gap-1">
                                <div className="w-1.5 h-1.5 bg-green-500 rounded-full"></div>
                                <span className="text-[9px] font-bold text-gray-500 uppercase">
                                  {complianceLabels.yes}
                                </span>
                              </div>
                              <div className="flex items-center gap-1">
                                <div className="w-1.5 h-1.5 bg-red-500 rounded-full"></div>
                                <span className="text-[9px] font-bold text-gray-500 uppercase">
                                  {complianceLabels.no}
                                </span>
                              </div>
                              <div className="flex items-center gap-1">
                                <div className="w-1.5 h-1.5 bg-gray-400 rounded-full"></div>
                                <span className="text-[9px] font-bold text-gray-500 uppercase">
                                  {complianceLabels.na}
                                </span>
                              </div>
                            </div>
                          </div>

                          {/* Radar Chart Container */}
                          <div className="h-[300px] sm:h-96">
                            {/* Prepare data for radar chart */}
                            {(() => {
                              // Prepare radar chart data
                              const radarChartData = {
                                labels: visibleSectionStats.map((stat) =>
                                  stat.title.length > 15
                                    ? stat.title.substring(0, 15) + "..."
                                    : stat.title,
                                ),

                                datasets: [
                                  {
                                    label: `${complianceLabels.yes} %`,
                                    data: visibleSectionStats.map((stat) =>
                                      stat.total > 0
                                        ? ((stat.yes + (stat.accepted || 0)) /
                                          stat.total) *
                                        100
                                        : 0,
                                    ),
                                    backgroundColor: "rgba(34, 197, 94, 0.2)",
                                    borderColor: "rgba(34, 197, 94, 1)",
                                    borderWidth: 2,
                                    pointBackgroundColor:
                                      "rgba(34, 197, 94, 1)",
                                    pointBorderColor: "#fff",
                                    pointHoverBackgroundColor: "#fff",
                                    pointHoverBorderColor:
                                      "rgba(34, 197, 94, 1)",
                                  },
                                  {
                                    label: `${complianceLabels.no} %`,
                                    data: visibleSectionStats.map((stat) =>
                                      stat.total > 0
                                        ? ((stat.no + (stat.rejected || 0)) /
                                          stat.total) *
                                        100
                                        : 0,
                                    ),
                                    backgroundColor: "rgba(239, 68, 68, 0.2)",
                                    borderColor: "rgba(239, 68, 68, 1)",
                                    borderWidth: 2,
                                    pointBackgroundColor:
                                      "rgba(239, 68, 68, 1)",
                                    pointBorderColor: "#fff",
                                    pointHoverBackgroundColor: "#fff",
                                    pointHoverBorderColor:
                                      "rgba(239, 68, 68, 1)",
                                  },
                                  {
                                    label: `${complianceLabels.na} %`,
                                    data: visibleSectionStats.map((stat) =>
                                      stat.total > 0
                                        ? ((stat.na + (stat.rework || 0)) /
                                          stat.total) *
                                        100
                                        : 0,
                                    ),
                                    backgroundColor: "rgba(156, 163, 175, 0.2)",
                                    borderColor: "rgba(156, 163, 175, 1)",
                                    borderWidth: 2,
                                    pointBackgroundColor:
                                      "rgba(156, 163, 175, 1)",
                                    pointBorderColor: "#fff",
                                    pointHoverBackgroundColor: "#fff",
                                    pointHoverBorderColor:
                                      "rgba(156, 163, 175, 1)",
                                  },
                                ],
                              };

                              const radarOptions = {
                                responsive: true,
                                maintainAspectRatio: false,
                                scales: {
                                  r: {
                                    angleLines: {
                                      display: true,
                                      color:
                                        document.documentElement.classList.contains(
                                          "dark",
                                        )
                                          ? "rgba(147, 197, 253, 0.4)"
                                          : "rgba(59, 130, 246, 0.4)",
                                      lineWidth: 1.5,
                                    },
                                    grid: {
                                      color:
                                        document.documentElement.classList.contains(
                                          "dark",
                                        )
                                          ? "rgba(147, 197, 253, 0.3)"
                                          : "rgba(59, 130, 246, 0.3)",
                                      lineWidth: 1.5,
                                    },
                                    pointLabels: {
                                      font: {
                                        size: 10,
                                      },
                                      color:
                                        document.documentElement.classList.contains(
                                          "dark",
                                        )
                                          ? "#e5e7eb"
                                          : "#374151",
                                    },
                                    ticks: {
                                      backdropColor: "transparent",
                                      color:
                                        document.documentElement.classList.contains(
                                          "dark",
                                        )
                                          ? "#9ca3af"
                                          : "#6b7280",
                                      font: {
                                        size: 11,
                                      },
                                    },
                                    suggestedMin: 0,
                                    suggestedMax: 100,
                                  },
                                },
                                plugins: {
                                  datalabels: {
                                    display: false,
                                  },
                                  legend: {
                                    position: "bottom",
                                    labels: {
                                      color:
                                        document.documentElement.classList.contains(
                                          "dark",
                                        )
                                          ? "#e5e7eb"
                                          : "#374151",
                                      font: {
                                        size: 10,
                                      },
                                      padding: 15,
                                    },
                                  },
                                  tooltip: {
                                    callbacks: {
                                      label: function (context) {
                                        return `${context.dataset.label}: ${context.raw.toFixed(1)}%`;
                                      },
                                    },
                                  },
                                },
                              };

                              return (
                                <Radar
                                  data={radarChartData}
                                  options={radarOptions}
                                />
                              );
                            })()}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              ) : (
                <div className="card p-6 text-center text-primary-500">
                  No section performance data available yet
                </div>
              )}

              <div className="card p-6">
                <SectionAnalytics
                  question={form}
                  responses={filteredResponses}
                  sectionsStats={filteredSectionsStats}
                  openSectionId={autoOpenSectionId}
                  complianceLabels={complianceLabels}
                />
              </div>

              {/* Analytics Zone */}
              <div className="card p-6 space-y-6">
                <div className="flex items-center justify-between">
                  <h3 className="text-lg font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                    <BarChart3 className="w-5 h-5 text-indigo-600" />
                    Analytics Zone
                  </h3>
                  <div className="flex items-center gap-4 text-xs font-medium">
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-3 bg-amber-500/60 rounded" />
                      <span className="text-gray-600 dark:text-gray-400">Rework</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-3 bg-red-500/60 rounded" />
                      <span className="text-gray-600 dark:text-gray-400">Rejected</span>
                    </div>
                  </div>
                </div>

                <div className="space-y-4">
                  <h4 className="text-sm font-semibold text-gray-700 dark:text-gray-300 uppercase tracking-wider">
                    Advanced Hierarchical Analytics (Zone &gt; Category &gt; Defect)
                  </h4>

                  {(() => {
                    const maxDefectCount = Math.max(
                      ...zoneAnalytics.zoneBreakdown.flatMap(z =>
                        z.categories.flatMap(c =>
                          c.defects.map(d => d.count)
                        )
                      ),
                      1
                    );

                    return (
                      <div className="border border-gray-200 dark:border-gray-700 rounded-2xl overflow-hidden bg-white dark:bg-gray-900 shadow-sm">
                        {/* Header Scale */}
                        <div className="flex bg-gray-50 dark:bg-gray-800/50 border-b border-gray-200 dark:border-gray-700">
                          <div className="w-1/3 min-w-[300px] p-4 border-r border-gray-200 dark:border-gray-700 font-bold text-xs text-gray-500 uppercase tracking-wider">Hierarchy</div>
                          <div className="flex-1 p-4 relative">
                            <div className="flex justify-between text-[10px] font-bold text-gray-400">
                              <span>0</span>
                              <span>{Math.round(maxDefectCount * 0.2)}</span>
                              <span>{Math.round(maxDefectCount * 0.4)}</span>
                              <span>{Math.round(maxDefectCount * 0.6)}</span>
                              <span>{Math.round(maxDefectCount * 0.8)}</span>
                              <span>{maxDefectCount}</span>
                            </div>
                            <div className="absolute inset-x-0 bottom-0 h-1 flex justify-between px-4">
                              {[0, 1, 2, 3, 4, 5].map(i => (
                                <div key={i} className="w-px h-full bg-gray-200 dark:bg-gray-700" />
                              ))}
                            </div>
                          </div>
                        </div>

                        {/* Hierarchical Content */}
                        <div className="divide-y divide-gray-100 dark:divide-gray-800">
                          {zoneAnalytics.zoneBreakdown.length === 0 ? (
                            <div className="p-12 text-center text-gray-500 italic">No defect data available for selected filters</div>
                          ) : (
                            zoneAnalytics.zoneBreakdown.map((zone) => (
                              <div key={zone.zone} className="flex group">
                                {/* Zone Label - Merged Side */}
                                <div className="w-[100px] p-4 flex items-center justify-center bg-indigo-50/30 dark:bg-indigo-900/10 border-r border-gray-200 dark:border-gray-700 shrink-0">
                                  <span className="[writing-mode:vertical-lr] rotate-180 font-bold text-sm text-indigo-700 dark:text-indigo-400 uppercase tracking-widest">{zone.zone}</span>
                                </div>

                                <div className="flex-1 divide-y divide-gray-100 dark:divide-gray-800">
                                  {zone.categories.map((cat) => (
                                    <div key={cat.category} className="flex">
                                      {/* Category Label */}
                                      <div className="w-[150px] p-4 flex items-center bg-gray-50/50 dark:bg-gray-800/20 border-r border-gray-200 dark:border-gray-700 shrink-0">
                                        <span className="font-semibold text-xs text-gray-700 dark:text-gray-300 leading-tight">{cat.category}</span>
                                      </div>

                                      <div className="flex-1 divide-y divide-gray-50 dark:divide-gray-800/50">
                                        {cat.defects.map((defect) => {
                                          const total = defect.count;
                                          const reworkWidth = total > 0 ? (defect.reworkCount / total) * 100 : 0;
                                          const rejectedWidth = total > 0 ? (defect.rejectedCount / total) * 100 : 0;
                                          const volumeWidth = (total / maxDefectCount) * 100;

                                          // Percentages relative to the global maximum for labels
                                          const reworkLabelPct = (defect.reworkCount / maxDefectCount) * 100;
                                          const rejectedLabelPct = (defect.rejectedCount / maxDefectCount) * 100;

                                          return (
                                            <div key={defect.name} className="flex items-center hover:bg-gray-50 dark:hover:bg-gray-800/30 transition-colors">
                                              {/* Defect Name */}
                                              <div className="w-[150px] p-3 border-r border-gray-100 dark:border-gray-800 shrink-0 flex items-center justify-between gap-1">
                                                <span className="text-[11px] font-medium text-gray-600 dark:text-gray-400 leading-tight">{defect.name}</span>
                                                <span className="text-[10px] font-bold text-gray-400 shrink-0">({total})</span>
                                              </div>

                                              {/* Bar Chart Section */}
                                              <div className="flex-1 p-3 px-4 relative flex items-center h-12">
                                                {/* Grid Lines Overlay */}
                                                <div className="absolute inset-0 flex justify-between px-4 pointer-events-none">
                                                  {[0, 1, 2, 3, 4, 5].map(i => (
                                                    <div key={i} className="w-px h-full bg-gray-100/50 dark:bg-gray-800/30" />
                                                  ))}
                                                </div>

                                                {/* Stacked Bar with Volume Normalization */}
                                                <div className="relative flex-1 h-6">
                                                  <div
                                                    style={{ width: `${volumeWidth}%` }}
                                                    className="h-full bg-gray-100 dark:bg-gray-800 rounded-full overflow-hidden flex shadow-inner transition-all duration-500"
                                                  >
                                                    {defect.reworkCount > 0 && (
                                                      <div
                                                        style={{ width: `${reworkWidth}%` }}
                                                        className="h-full bg-gradient-to-r from-amber-400 to-amber-500 relative group/bar"
                                                        title={`Rework: ${defect.reworkCount} (${reworkLabelPct.toFixed(1)}%)`}
                                                      >
                                                        {reworkLabelPct > 15 && (
                                                          <span className="absolute inset-0 flex items-center justify-center text-[9px] font-bold text-amber-900">
                                                            {reworkLabelPct.toFixed(0)}%
                                                          </span>
                                                        )}
                                                      </div>
                                                    )}
                                                    {defect.rejectedCount > 0 && (
                                                      <div
                                                        style={{ width: `${rejectedWidth}%` }}
                                                        className="h-full bg-gradient-to-r from-red-400 to-red-500 relative group/bar"
                                                        title={`Rejected: ${defect.rejectedCount} (${rejectedLabelPct.toFixed(1)}%)`}
                                                      >
                                                        {rejectedLabelPct > 15 && (
                                                          <span className="absolute inset-0 flex items-center justify-center text-[9px] font-bold text-white">
                                                            {rejectedLabelPct.toFixed(0)}%
                                                          </span>
                                                        )}
                                                      </div>
                                                    )}
                                                  </div>
                                                </div>
                                              </div>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              </div>
                            ))
                          )}
                        </div>
                      </div>
                    );
                  })()}
                </div>

                {/* Status Summary */}
                <div className="grid grid-cols-3 gap-4 pt-4 border-t border-gray-100 dark:border-gray-800">
                  <div className="p-4 rounded-xl bg-emerald-50 dark:bg-emerald-900/20 border border-emerald-100 dark:border-emerald-800/50">
                    <p className="text-[10px] font-bold text-emerald-600 dark:text-emerald-400 uppercase tracking-wider mb-1">Accepted</p>
                    <p className="text-2xl font-bold text-emerald-700 dark:text-emerald-300">{zoneAnalytics.inspectionStatus.accepted}</p>
                  </div>
                  <div className="p-4 rounded-xl bg-amber-50 dark:bg-amber-900/20 border border-amber-100 dark:border-amber-800/50">
                    <p className="text-[10px] font-bold text-amber-600 dark:text-amber-400 uppercase tracking-wider mb-1">Rework</p>
                    <p className="text-2xl font-bold text-amber-700 dark:text-amber-300">{zoneAnalytics.inspectionStatus.rework}</p>
                  </div>
                  <div className="p-4 rounded-xl bg-red-50 dark:bg-red-900/20 border border-red-100 dark:border-red-800/50">
                    <p className="text-[10px] font-bold text-red-600 dark:text-red-400 uppercase tracking-wider mb-1">Rejected</p>
                    <p className="text-2xl font-bold text-red-700 dark:text-red-300">{zoneAnalytics.inspectionStatus.rejected}</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Table View */}
          {analyticsView === "table" && (
            <div className="space-y-6">
              {/* Table View Type Selector */}
              <div className="card p-4 flex gap-3 items-center">
                <span className="text-sm font-semibold text-gray-700 dark:text-gray-300">
                  View Type:
                </span>
                <div className="flex gap-2">
                  <button
                    onClick={() => setTableViewType("question")}
                    className={`px-4 py-2 rounded-lg font-medium transition-all ${tableViewType === "question"
                      ? "bg-indigo-600 text-white shadow-md"
                      : "bg-gray-200 dark:bg-gray-700 text-gray-900 dark:text-gray-300 hover:bg-gray-300"
                      }`}
                  >
                    Question Based
                  </button>
                  <button
                    onClick={() => setTableViewType("section")}
                    className={`px-4 py-2 rounded-lg font-medium transition-all ${tableViewType === "section"
                      ? "bg-indigo-600 text-white shadow-md"
                      : "bg-gray-200 dark:bg-gray-700 text-gray-900 dark:text-gray-300 hover:bg-gray-300"
                      }`}
                  >
                    Section Based
                  </button>
                </div>
              </div>

              {/* Question Based Table - All Questions from All Sections */}
              {tableViewType === "question" &&
                form?.sections &&
                form.sections.length > 0 && (
                  <div className="card p-6">
                    <div className="mb-4">
                      <h3 className="text-lg font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                        <BarChart3 className="w-5 h-5 text-indigo-600" />
                        All Questions Analytics - Table View
                      </h3>
                      <p className="text-sm text-gray-600 dark:text-gray-400 mt-1">
                        Showing all questions from all sections including
                        follow-ups
                      </p>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse">
                        <thead>
                          <tr className="bg-gradient-to-r from-indigo-50 to-blue-50 dark:from-gray-700 dark:to-gray-600 border-b-2 border-indigo-200 dark:border-indigo-700">
                            <th className="px-6 py-3 text-left text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              Question
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              Total Responses
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              {complianceLabels.yes}
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              {complianceLabels.no}
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              {complianceLabels.na}
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white">
                              {complianceLabels.yes} %
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200 dark:divide-gray-700">
                          {form.sections.map(
                            (section: Section, sectionIdx: number) => {
                              const allQuestionsInSection =
                                section.questions || [];

                              return (
                                <React.Fragment key={`section-${section.id}`}>
                                  <tr className="bg-indigo-100 dark:bg-indigo-900/40 hover:bg-indigo-100 dark:hover:bg-indigo-900/40">
                                    <td
                                      colSpan={6}
                                      className="px-6 py-4 text-center text-sm font-bold text-indigo-800 dark:text-indigo-300 uppercase tracking-wide"
                                    >
                                      {section.title}
                                    </td>
                                  </tr>
                                  {allQuestionsInSection.map(
                                    (question: any, qIdx: number) => {
                                      const questionResponses =
                                        filteredResponses.filter(
                                          (r) =>
                                            r.answers && r.answers[question.id],
                                        );
                                      const yesCount = questionResponses.filter(
                                        (r) => {
                                          const answer = r.answers[question.id];
                                          if (
                                            typeof answer === "object" &&
                                            answer.status
                                          ) {
                                            const status = String(answer.status)
                                              .toLowerCase()
                                              .trim();
                                            return (
                                              status === "accepted" ||
                                              status === "rework completed" ||
                                              status === "verified"
                                            );
                                          }
                                          const answerStr = String(answer)
                                            .toLowerCase()
                                            .trim();
                                          return (
                                            answerStr.includes("yes") ||
                                            answerStr === "y"
                                          );
                                        },
                                      ).length;
                                      const noCount = questionResponses.filter(
                                        (r) => {
                                          const answer = r.answers[question.id];
                                          if (
                                            typeof answer === "object" &&
                                            answer.status
                                          ) {
                                            return (
                                              String(answer.status)
                                                .toLowerCase()
                                                .trim() === "rejected"
                                            );
                                          }
                                          const answerStr = String(answer)
                                            .toLowerCase()
                                            .trim();
                                          return (
                                            answerStr.includes("no") ||
                                            answerStr === "n"
                                          );
                                        },
                                      ).length;
                                      const naCount = questionResponses.filter(
                                        (r) => {
                                          const answer = r.answers[question.id];
                                          if (
                                            typeof answer === "object" &&
                                            answer.status
                                          ) {
                                            const status = String(answer.status)
                                              .toLowerCase()
                                              .trim();
                                            return (
                                              status === "rework" ||
                                              status === "reworked" ||
                                              status.includes("re-rework")
                                            );
                                          }
                                          const answerStr = String(answer)
                                            .toLowerCase()
                                            .trim();
                                          return (
                                            answerStr.includes("na") ||
                                            answerStr.includes("n/a") ||
                                            answerStr.includes("not applicable")
                                          );
                                        },
                                      ).length;
                                      const total = questionResponses.length;
                                      const yesPercentage =
                                        total > 0
                                          ? ((yesCount / total) * 100).toFixed(
                                            1,
                                          )
                                          : "0.0";

                                      const isFollowUp =
                                        question.parentId ||
                                        question.showWhen?.questionId;

                                      return (
                                        <tr
                                          key={question.id}
                                          className={`hover:bg-indigo-50 dark:hover:bg-gray-700 transition-colors ${isFollowUp
                                            ? "bg-purple-50 dark:bg-purple-900/20"
                                            : "bg-white dark:bg-gray-800"
                                            }`}
                                        >
                                          <td
                                            className={`px-6 py-4 text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium max-w-sm ${isFollowUp ? "pl-12" : ""
                                              }`}
                                          >
                                            <div
                                              className="truncate"
                                              title={
                                                question.text ||
                                                "Unnamed Question"
                                              }
                                            >
                                              {question.text ||
                                                "Unnamed Question"}
                                            </div>
                                          </td>
                                          <td className="px-6 py-4 text-center text-sm font-semibold text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700">
                                            <span className="bg-blue-100 dark:bg-blue-900/30 text-blue-900 dark:text-blue-300 px-3 py-1 rounded-full text-xs">
                                              {total}
                                            </span>
                                          </td>
                                          <td className="px-6 py-4 text-center text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium">
                                            <span className="bg-green-100 dark:bg-green-900/30 text-green-900 dark:text-green-300 px-3 py-1 rounded-full text-xs">
                                              {yesCount}
                                            </span>
                                          </td>
                                          <td className="px-6 py-4 text-center text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium">
                                            <span className="bg-red-100 dark:bg-red-900/30 text-red-900 dark:text-red-300 px-3 py-1 rounded-full text-xs">
                                              {noCount}
                                            </span>
                                          </td>
                                          <td className="px-6 py-4 text-center text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium">
                                            <span className="bg-gray-300 dark:bg-gray-600 text-gray-900 dark:text-gray-200 px-3 py-1 rounded-full text-xs">
                                              {naCount}
                                            </span>
                                          </td>
                                          <td className="px-6 py-4 text-center text-sm font-semibold text-indigo-600 dark:text-indigo-400">
                                            {yesPercentage}%
                                          </td>
                                        </tr>
                                      );
                                    },
                                  )}
                                </React.Fragment>
                              );
                            },
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

              {/* Section Based Table */}
              {tableViewType === "section" &&
                filteredSectionStats.length > 0 && (
                  <div className="card p-6">
                    <div className="mb-4">
                      <h3 className="text-lg font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                        <FileText className="w-5 h-5 text-indigo-600" />
                        Section Analytics - Table View
                      </h3>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse">
                        <thead>
                          <tr className="bg-gradient-to-r from-indigo-50 to-blue-50 dark:from-gray-700 dark:to-gray-600 border-b-2 border-indigo-200 dark:border-indigo-700">
                            <th className="px-6 py-3 text-left text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              Section Name
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              Total
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              {complianceLabels.yes}
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white border-r border-indigo-200 dark:border-indigo-700">
                              {complianceLabels.no}
                            </th>
                            <th className="px-6 py-3 text-center text-sm font-semibold text-gray-900 dark:text-white">
                              {complianceLabels.na}
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200 dark:divide-gray-700">
                          {filteredSectionStats.map(
                            (stat: SectionPerformanceStat, index: number) => {
                              const totalYes = stat.yes + (stat.accepted || 0);
                              const totalNo = stat.no + (stat.rejected || 0);
                              const totalNA = stat.na + (stat.rework || 0);

                              const yesPercentage =
                                stat.total > 0
                                  ? ((totalYes / stat.total) * 100).toFixed(1)
                                  : "0.0";
                              const noPercentage =
                                stat.total > 0
                                  ? ((totalNo / stat.total) * 100).toFixed(1)
                                  : "0.0";
                              const naPercentage =
                                stat.total > 0
                                  ? ((totalNA / stat.total) * 100).toFixed(1)
                                  : "0.0";

                              return (
                                <tr
                                  key={stat.id}
                                  className={`${index % 2 === 0 ? "bg-white dark:bg-gray-800" : "bg-gray-50 dark:bg-gray-750"} hover:bg-indigo-50 dark:hover:bg-gray-700 transition-colors`}
                                >
                                  <td className="px-6 py-4 text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium">
                                    {stat.title}
                                  </td>
                                  <td className="px-6 py-4 text-center text-sm font-semibold text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700">
                                    <span className="bg-blue-100 dark:bg-blue-900/30 text-blue-900 dark:text-blue-300 px-3 py-1 rounded-full text-xs">
                                      {stat.total}
                                    </span>
                                  </td>
                                  <td className="px-6 py-4 text-center text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium">
                                    <span className="bg-green-100 dark:bg-green-900/30 text-green-900 dark:text-green-300 px-3 py-1 rounded-full text-xs">
                                      {totalYes} ({yesPercentage}%)
                                    </span>
                                  </td>
                                  <td className="px-6 py-4 text-center text-sm text-gray-900 dark:text-gray-300 border-r border-gray-200 dark:border-gray-700 font-medium">
                                    <span className="bg-red-100 dark:bg-red-900/30 text-red-900 dark:text-red-300 px-3 py-1 rounded-full text-xs">
                                      {totalNo} ({noPercentage}%)
                                    </span>
                                  </td>
                                  <td className="px-6 py-4 text-center text-sm text-gray-900 dark:text-gray-300 font-medium">
                                    <span className="bg-gray-300 dark:bg-gray-600 text-gray-900 dark:text-gray-200 px-3 py-1 rounded-full text-xs">
                                      {totalNA} ({naPercentage}%)
                                    </span>
                                  </td>
                                </tr>
                              );
                            },
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
            </div>
          )}

          {/* Responses as Table */}
          {analyticsView === "responses" && (
            <div className="space-y-4 sm:space-y-6">
              <div className="card p-3 sm:p-6">
                <div className="mb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="flex flex-col">
                    <h3 className="text-base sm:text-lg font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                      <Table className="w-5 h-5 text-indigo-600" />
                      All Responses
                    </h3>
                    <p className="text-xs sm:text-sm text-gray-600 dark:text-gray-400">
                      Viewing {filteredResponses.length} responses
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2 items-center relative">
                    <button
                      onClick={() =>
                        setShowResponsesFilter(!showResponsesFilter)
                      }
                      className={`px-3 sm:px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs sm:text-sm font-semibold transition-all duration-200 flex items-center gap-2 ${showResponsesFilter ? "ring-2 ring-indigo-400 ring-offset-2 dark:ring-offset-gray-900" : ""}`}
                    >
                      <svg
                        className="w-4 h-4"
                        fill="none"
                        stroke="currentColor"
                        viewBox="0 0 24 24"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          strokeWidth={2}
                          d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4"
                        />
                      </svg>
                      <span className="hidden xs:inline">Filter Sections</span>
                      <span className="xs:hidden">Filter</span>
                    </button>
                    <button
                      onClick={() => handleExportToExcel()}
                      disabled={selectedResponsesSectionIds.length === 0}
                      className="px-3 sm:px-4 py-2 bg-green-600 hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed text-white rounded-lg text-xs sm:text-sm font-semibold transition-colors flex items-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      <span className="hidden xs:inline">Export</span>
                    </button>
                    <button
                      onClick={() => handleExportOPSToExcel()}
                      disabled={selectedResponsesSectionIds.length === 0}
                      className="px-3 sm:px-4 py-2 bg-indigo-700 hover:bg-indigo-800 disabled:bg-gray-400 disabled:cursor-not-allowed text-white rounded-lg text-xs sm:text-sm font-semibold transition-colors flex items-center gap-2"
                      title="Export responses as OPS Excel — questions as columns with answers below"
                    >
                      <Download className="w-4 h-4" />
                      <span>OPS responses as excel</span>
                    </button>
                    <button
                      onClick={handleExportToPDF_OPS}
                      disabled={isExporting}
                      className="p-1.5 sm:p-2 text-gray-500 hover:text-red-600 dark:text-gray-400 dark:hover:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors disabled:opacity-50"
                      title="Export to OPS PDF (A3)"
                    >
                      <Download className="w-4 h-4" />
                      <span>OPS responses as PDF</span>
                      {isExporting ? (
                        <Loader2 className="w-4 h-4 animate-spin text-red-500" />
                      ) : (
                        <FileText className="w-4 h-4" />
                      )}
                    </button>
                    {selectedResponseIds.length > 0 && !isGuest && (
                      <button
                        onClick={() => setShowBulkDeleteConfirm(true)}
                        className="px-3 sm:px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-xs sm:text-sm font-semibold transition-colors flex items-center gap-2"
                      >
                        <Trash2 className="w-4 h-4" />
                        <span className="hidden xs:inline">Delete ({selectedResponseIds.length})</span>
                      </button>
                    )}

                    {showResponsesFilter && (
                      <div className="absolute top-full right-0 mt-2 p-0 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-lg shadow-xl z-50 w-[280px] sm:w-80 animate-in fade-in slide-in-from-top-2 duration-200">
                        <div className="sticky top-0 bg-gradient-to-r from-indigo-50 to-blue-50 dark:from-indigo-900/20 dark:to-blue-900/20 px-3 sm:px-4 py-3 border-b border-gray-200 dark:border-gray-700">
                          <div className="flex items-center justify-between mb-2">
                            <h4 className="text-sm font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                              Select Sections
                            </h4>
                            <button
                              onClick={() => setShowResponsesFilter(false)}
                              className="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 rounded transition-colors"
                            >
                              <svg
                                className="w-4 h-4 text-gray-600 dark:text-gray-400"
                                fill="none"
                                stroke="currentColor"
                                viewBox="0 0 24 24"
                              >
                                <path
                                  strokeLinecap="round"
                                  strokeLinejoin="round"
                                  strokeWidth={2}
                                  d="M6 18L18 6M6 6l12 12"
                                />
                              </svg>
                            </button>
                          </div>
                          <div className="flex gap-2">
                            <button
                              onClick={() =>
                                setSelectedResponsesSectionIds(
                                  form?.sections?.map((s: Section) => s.id) ||
                                  [],
                                )
                              }
                              className="flex-1 px-2 py-1.5 text-[10px] font-semibold text-indigo-600 dark:text-indigo-400 bg-indigo-100 dark:bg-indigo-900/30 hover:bg-indigo-200 dark:hover:bg-indigo-900/50 rounded transition-colors"
                            >
                              Select All
                            </button>
                            <button
                              onClick={() => setSelectedResponsesSectionIds([])}
                              className="flex-1 px-2 py-1.5 text-[10px] font-semibold text-gray-600 dark:text-gray-400 bg-gray-100 dark:bg-gray-700 hover:bg-gray-200 dark:hover:bg-gray-600 rounded transition-colors"
                            >
                              Clear All
                            </button>
                          </div>
                        </div>

                        <div className="p-2 sm:p-4 max-h-64 sm:max-h-96 overflow-y-auto space-y-1">
                          {form?.sections && form.sections.length > 0 ? (
                            form.sections.map((section: Section) => (
                              <label
                                key={section.id}
                                className="flex items-center gap-3 p-2 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-700/50 cursor-pointer transition-colors group"
                              >
                                <div className="relative flex items-center">
                                  <input
                                    type="checkbox"
                                    checked={selectedResponsesSectionIds.includes(
                                      section.id,
                                    )}
                                    onChange={(e) => {
                                      if (e.target.checked) {
                                        setSelectedResponsesSectionIds([
                                          ...selectedResponsesSectionIds,
                                          section.id,
                                        ]);
                                      } else {
                                        setSelectedResponsesSectionIds(
                                          selectedResponsesSectionIds.filter(
                                            (id) => id !== section.id,
                                          ),
                                        );
                                      }
                                    }}
                                    className="w-4 h-4 text-indigo-600 border-gray-300 dark:border-gray-600 rounded cursor-pointer accent-indigo-600"
                                  />
                                </div>
                                <div className="flex-1 min-w-0">
                                  <span className="text-xs font-medium text-gray-900 dark:text-gray-200 block truncate">
                                    {section.title}
                                  </span>
                                </div>
                              </label>
                            ))
                          ) : (
                            <p className="text-xs text-gray-500 dark:text-gray-400 text-center py-4">
                              No sections available
                            </p>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {selectedResponsesSectionIds.length > 0 ? (
                  <>
                    {/* Overall Inspection Statistics Summary Bar */}
                    <div className="bg-gray-50 dark:bg-gray-800/50 rounded-xl border border-gray-200 dark:border-gray-700 p-3 sm:p-4 mb-4">
                      <div className="grid grid-cols-1 xs:grid-cols-2 md:grid-cols-4 gap-4 items-center">
                        <div className="flex items-center gap-3">
                          <div className="p-2 bg-indigo-100 dark:bg-indigo-900/30 rounded-lg">
                            <BarChart3 className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                          </div>
                          <div>
                            <p className="text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase tracking-widest">
                              Summary
                            </p>
                            <p className="text-sm font-bold text-gray-900 dark:text-white">
                              Performance
                            </p>
                          </div>
                        </div>

                        <div className="flex flex-col p-2 bg-green-50 dark:bg-green-900/10 rounded-lg border border-green-100 dark:border-green-900/20">
                          <span className="text-[10px] text-green-700 dark:text-green-400 font-bold uppercase">
                            {complianceLabels.yes}
                          </span>
                          <div className="flex items-baseline justify-between">
                            <span className="text-lg font-black text-green-600 dark:text-green-400">
                              {inspectionStats.accepted}
                            </span>
                            <CheckCircle className="w-4 h-4 text-green-500" />
                          </div>
                        </div>

                        <div className="flex flex-col p-2 bg-red-50 dark:bg-red-900/10 rounded-lg border border-red-100 dark:border-red-900/20">
                          <span className="text-[10px] text-red-700 dark:text-red-400 font-bold uppercase">
                            {complianceLabels.no}
                          </span>
                          <div className="flex items-baseline justify-between">
                            <span className="text-lg font-black text-red-600 dark:text-red-400">
                              {inspectionStats.rejected}
                            </span>
                            <XCircle className="w-4 h-4 text-red-500" />
                          </div>
                        </div>

                        <div className="flex flex-col p-2 bg-amber-50 dark:bg-amber-900/10 rounded-lg border border-amber-100 dark:border-amber-900/20">
                          <span className="text-[10px] text-amber-700 dark:text-amber-400 font-bold uppercase">
                            {complianceLabels.na}
                          </span>
                          <div className="flex items-baseline justify-between">
                            <span className="text-lg font-black text-amber-500 dark:text-amber-400">
                              {inspectionStats.reworked}
                            </span>
                            <span className="text-amber-500 text-sm font-bold">⚠</span>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="overflow-x-auto no-scrollbar rounded-xl border border-gray-200 dark:border-gray-700">
                      <table className="text-xs border-collapse w-full">
                        <thead className="sticky top-18 z-30">
                          <tr className="bg-gray-100 dark:bg-gray-800">
                            <th className="hidden sm:table-cell sticky left-0 z-30 text-center px-3 py-3 font-semibold text-gray-600 dark:text-gray-300 border border-gray-200 dark:border-gray-700 bg-gray-100 dark:bg-gray-800">
                              <input
                                type="checkbox"
                                checked={
                                  selectedResponseIds.length > 0 &&
                                  selectedResponseIds.length ===
                                  filteredResponses.length
                                }
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    setSelectedResponseIds(
                                      filteredResponses.map((r) => r.id),
                                    );
                                  } else {
                                    setSelectedResponseIds([]);
                                  }
                                }}
                                className="w-4 h-4 text-indigo-600 border-gray-300 dark:border-gray-600 rounded cursor-pointer accent-indigo-600"
                              />
                            </th>
                            <th className="sticky left-0 sm:left-12 z-30 text-center px-4 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 whitespace-nowrap bg-gray-100 dark:bg-gray-800 min-w-[120px]">
                              <span>Actions</span>
                            </th>
                            <th className="text-center px-6 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 min-w-24 whitespace-nowrap bg-gray-100 dark:bg-gray-800">
                              Dispatch
                            </th>
                            <th className="text-left px-6 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 min-w-48 whitespace-nowrap bg-gray-50 dark:bg-gray-800">
                              Submitted by
                            </th>
                            <th className="text-left px-6 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 min-w-32 whitespace-nowrap bg-gray-50 dark:bg-gray-800">
                              Status
                            </th>
                            <th className="text-left px-6 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 min-w-40 whitespace-nowrap bg-gray-50 dark:bg-gray-800">
                              Selected Chassis
                            </th>
                            <th className="text-left px-6 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 min-w-48 whitespace-nowrap bg-gray-50 dark:bg-gray-800">
                              Review
                            </th>
                            <th className="text-left px-6 py-3 font-semibold text-gray-600 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-700 min-w-40 whitespace-nowrap">
                              Timestamp
                            </th>
                            <th className="text-center px-4 py-3 font-semibold text-blue-600 dark:text-blue-400 uppercase tracking-wider border border-gray-200 dark:border-gray-700 whitespace-nowrap bg-gray-50 dark:bg-gray-800">
                              Time Taken
                            </th>
                            {form?.sections?.map(
                              (section: Section) =>
                                selectedResponsesSectionIds.includes(
                                  section.id,
                                ) &&
                                section.questions?.map((q: any) => {
                                  const isFollowUp =
                                    q.parentId || q.showWhen?.questionId;
                                  const columnOptions = getUniqueColumnValues(
                                    q.id,
                                    responses,
                                  );

                                  return (
                                    <th
                                      key={q.id}
                                      className={`text-left px-4 py-3 font-semibold text-xs uppercase tracking-wider border border-gray-200 dark:border-gray-700 max-w-xs ${isFollowUp ? "bg-purple-100 dark:bg-purple-900/30" : "bg-gray-100 dark:bg-gray-800"}`}
                                    >
                                      <div className="flex items-center justify-between gap-2">
                                        <div className="line-clamp-2 overflow-hidden text-ellipsis flex-1">
                                          {q.text || "Question"}
                                        </div>
                                        <TableColumnFilter
                                          columnId={q.id}
                                          title={q.text || "Question"}
                                          options={columnOptions}
                                          selectedValues={
                                            columnFilters[q.id] || null
                                          }
                                          onFilterChange={(
                                            columnId,
                                            values,
                                          ) => {
                                            setColumnFilters((prev) => ({
                                              ...prev,
                                              [columnId]: values,
                                            }));
                                          }}
                                        />
                                      </div>
                                    </th>
                                  );
                                }),
                            )}
                          </tr>
                        </thead>

                        <tbody className="divide-y divide-gray-200 dark:divide-gray-700">
                          {filteredResponses.length > 0 ? (
                            filteredResponses.map(
                              (response: Response, idx: number) => (
                                <tr
                                  key={response.id}
                                  className={`${editingResponseId === response.id ? "bg-blue-50 dark:bg-blue-900/20" : idx % 2 === 0 ? "bg-white dark:bg-gray-900" : "bg-gray-50 dark:bg-gray-800"}`}
                                >
                                  <td
                                    className={`hidden sm:table-cell  px-3 py-3 text-center border border-gray-200 dark:border-gray-700 whitespace-nowrap sticky left-0 z-20 ${editingResponseId === response.id ? "bg-blue-50 dark:bg-blue-900/20" : idx % 2 === 0 ? "bg-white dark:bg-gray-900" : "bg-gray-50 dark:bg-gray-800"}`}
                                  >
                                    <input
                                      type="checkbox"
                                      checked={selectedResponseIds.includes(
                                        response.id,
                                      )}
                                      onChange={(e) => {
                                        if (e.target.checked) {
                                          setSelectedResponseIds([
                                            ...selectedResponseIds,
                                            response.id,
                                          ]);
                                        } else {
                                          setSelectedResponseIds(
                                            selectedResponseIds.filter(
                                              (id) => id !== response.id,
                                            ),
                                          );
                                        }
                                      }}
                                      className="w-4 h-4 text-indigo-600 border-gray-300 dark:border-gray-600 rounded cursor-pointer accent-indigo-600"
                                    />
                                  </td>
                                  <td
                                    className={`px-4 py-3 text-center border border-gray-200 dark:border-gray-700 whitespace-nowrap sticky left-0 sm:left-12 z-20 transition-all duration-300 ${editingResponseId === response.id ? "bg-blue-50 dark:bg-blue-900/20" : idx % 2 === 0 ? "bg-white dark:bg-gray-900" : "bg-gray-50 dark:bg-gray-800"}`}
                                  >
                                    <div className="flex items-center gap-1.5 justify-center">
                                      {editingResponseId === response.id ? (
                                        <>
                                          <button
                                            onClick={handleSaveEdit}
                                            disabled={isSaving}
                                            title="Save Response"
                                            className="p-1 text-green-600 dark:text-green-400 hover:bg-green-100 dark:hover:bg-green-900/30 rounded transition-colors disabled:opacity-50"
                                          >
                                            <CheckCircle className="w-4 h-4" />
                                          </button>
                                          <button
                                            onClick={handleCancelEdit}
                                            disabled={isSaving}
                                            title="Cancel"
                                            className="p-1 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-700 rounded transition-colors disabled:opacity-50"
                                          >
                                            <XCircle className="w-4 h-4" />
                                          </button>
                                        </>
                                      ) : (
                                        <>
                                          {/* Focus It */}
                                          <button
                                            onClick={() => handleOpenModal(response)}
                                            className="p-1.5 text-blue-500 hover:text-blue-700 dark:text-blue-400 dark:hover:text-blue-300 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded-lg transition-all"
                                            title="Focus It"
                                          >
                                            <Maximize className="w-4 h-4" />
                                          </button>

                                          {(() => {
                                            const responseTenantId = response.tenantId;
                                            const currentUserTenantId = user?.tenantId;
                                            const isActualOwnTenant = user?.role === 'superadmin' || !responseTenantId || (currentUserTenantId && responseTenantId.toString() === currentUserTenantId.toString());
                                            const isOwnTenant = isActualOwnTenant;

                                            return (
                                              <>
                                                {/* View Details */}
                                                {isOwnTenant && (
                                                  <button
                                                    onClick={() => handleViewDetails(response)}
                                                    className="p-1.5 text-indigo-600 dark:text-indigo-400 hover:text-indigo-800 dark:hover:text-indigo-300 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 rounded-lg transition-all"
                                                    title="View Full Details"
                                                  >
                                                    <Eye className="w-4 h-4" />
                                                  </button>
                                                )}

                                                {/* Review & Discussion */}
                                                {response.isDispatched && (
                                                  <button
                                                    onClick={() => {
                                                      setChatResponse(response);
                                                      setShowChatModal(true);
                                                      setSelectedReviewOptions(prev => ({ ...prev, [response.id]: '' }));
                                                      setReviewedBy(prev => ({ ...prev, [response.id]: null }));
                                                    }}
                                                    className="p-1.5 text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded-lg transition-all"
                                                    title="Review & Discussion"
                                                  >
                                                    <MessageCircle className="w-4 h-4" />
                                                  </button>
                                                )}

                                                {/* Edit & Delete - Admin only */}
                                                {!isGuest && (user?.role === "superadmin" || user?.role === "admin") && isActualOwnTenant && (
                                                  <>
                                                    <button
                                                      onClick={() => handleEditStart(response)}
                                                      className="p-1.5 text-amber-600 dark:text-amber-400 hover:text-amber-800 dark:hover:text-amber-300 hover:bg-amber-50 dark:hover:bg-amber-900/20 rounded-lg transition-all"
                                                      title="Edit Response"
                                                    >
                                                      <Edit className="w-4 h-4" />
                                                    </button>
                                                    <button
                                                      onClick={() => {
                                                        setDeletingResponseId(response.id);
                                                        setShowDeleteConfirm(true);
                                                      }}
                                                      className="p-1.5 text-red-600 dark:text-red-400 hover:text-red-800 dark:hover:text-red-300 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-all"
                                                      title="Delete Response"
                                                    >
                                                      <Trash2 className="w-4 h-4" />
                                                    </button>
                                                  </>
                                                )}
                                              </>
                                            );
                                          })()}
                                        </>
                                      )}
                                    </div>
                                  </td>
                                  <td
                                    className={`px-3 py-3 text-center border border-gray-200 dark:border-gray-700 whitespace-nowrap ${editingResponseId === response.id ? "bg-blue-50 dark:bg-blue-900/20" : idx % 2 === 0 ? "bg-white dark:bg-gray-900" : "bg-gray-50 dark:bg-gray-800"}`}
                                  >
                                    {/* Dispatch cell content */}
                                    {(() => {
                                      const status = responseStatuses[response.id] || "";
                                      const canShowDispatch = status === "Direct Ok" || status === "Rework Accepted" || status === "Rework Completed" || status === "Accepted";
                                      // Check submitter based on email/username like other parts of the code
                                      const userEmail = user?.email || "";
                                      const userUsername = user?.username || "";
                                      const userIdStr = user?._id ? String(user._id) : (user?.id ? String(user.id) : "");
                                      const creatorId = typeof response.createdBy === 'object'
                                        ? (response.createdBy as any)?._id || (response.createdBy as any)?.id
                                        : response.createdBy;
                                      const creatorIdStr = creatorId ? String(creatorId) : "";

                                      const isSubmitter = response.submittedBy === userEmail ||
                                        response.submittedBy === userUsername ||
                                        response.createdBy === userEmail ||
                                        response.createdBy === userUsername ||
                                        response.submitterContact?.email === userEmail || (creatorIdStr && creatorIdStr === userIdStr);

                                      // Debug: Uncomment to see what statuses are being checked
                                      console.log(`Dispatch check for response ${response.id}:`, {
                                        status,
                                        canShow: canShowDispatch,
                                        isSubmitter,
                                        userEmail,
                                        userUsername,
                                        responseSubmittedBy: response.submittedBy,
                                        responseCreatedBy: response.createdBy,
                                        responseSubmitterEmail: response.submitterContact?.email
                                      });

                                      if (!canShowDispatch) {
                                        return <span className="text-gray-400 text-xs">-</span>;
                                      }

                                      // Show enabled state for all users once dispatch is enabled
                                      if (response.isDispatched) {
                                        return (
                                          <div className="flex items-center justify-center">
                                            <input
                                              type="checkbox"
                                              checked={true}
                                              disabled={true}
                                              className="w-4 h-4 text-green-600 border-gray-300 dark:border-gray-600 rounded accent-green-600 opacity-60"
                                              title="Dispatch enabled"
                                            />
                                            <span className="ml-2 text-xs text-green-600 font-medium">Enabled</span>
                                          </div>
                                        );
                                      }

                                      // Only show interactive checkbox for the submitter
                                      return (
                                        <input
                                          type="checkbox"
                                          checked={false}
                                          onChange={async (e) => {
                                            if (e.target.checked && !response.isDispatched) {
                                              try {
                                                await apiClient.updateResponse(response.id, { isDispatched: true });
                                                // Update local state to reflect change immediately
                                                setResponses(prev => prev.map(r =>
                                                  r.id === response.id ? { ...r, isDispatched: true, dispatchedAt: new Date().toISOString() } : r
                                                ));
                                              } catch (error) {
                                                console.error('Failed to enable dispatch:', error);
                                                alert('Failed to enable dispatch. Please try again.');
                                              }
                                            }
                                          }}
                                          className="w-4 h-4 text-green-600 border-gray-300 dark:border-gray-600 rounded cursor-pointer accent-green-600"
                                          title="Enable dispatch (only submitter can do this)"
                                        />
                                      );
                                    })()}
                                  </td>
                                  <td className="px-6 py-3 text-sm text-gray-900 dark:text-white font-bold border border-gray-200 dark:border-gray-700 min-w-48 whitespace-nowrap bg-gray-50/50 dark:bg-gray-800/30">
                                    {response.submittedBy ||
                                      response.createdBy ||
                                      "Anonymous"}
                                  </td>
                                  <td className="px-6 py-3 text-sm font-bold border border-gray-200 dark:border-gray-700 min-w-32 whitespace-nowrap bg-gray-50/50 dark:bg-gray-800/30">
                                    <span
                                      className={`px-2 py-1 rounded-full text-xs ${responseStatuses[response.id] === "Rejected"
                                        ? "bg-red-100 dark:bg-red-900/30 text-red-700 dark:text-red-300"
                                        : responseStatuses[response.id]?.includes("Rework") &&
                                          responseStatuses[response.id] !== "Rework Accepted"
                                          ? "bg-amber-100 dark:bg-amber-900/30 text-amber-700 dark:text-amber-300"
                                          : responseStatuses[response.id] === "Direct Ok" ||
                                            responseStatuses[response.id] === "Rework Accepted" ||
                                            responseStatuses[response.id] === "Accepted"
                                            ? "bg-green-100 dark:bg-green-900/30 text-green-700 dark:text-green-300"
                                            : responseStatuses[response.id] === "Pending Review"
                                              ? "bg-yellow-100 dark:bg-yellow-900/30 text-yellow-700 dark:text-yellow-300"
                                              : "bg-gray-100 dark:bg-gray-800 text-gray-500 dark:text-gray-400"
                                        }`}
                                    >
                                      {responseStatuses[response.id] || "Pending Review"}
                                    </span>
                                  </td>
                                  <td className="px-6 py-3 text-sm text-gray-900 dark:text-white font-medium border border-gray-200 dark:border-gray-700 min-w-40 whitespace-nowrap bg-gray-50/50 dark:bg-gray-800/30">
                                    {response.answers?.chassis_number || "-"}
                                  </td>
                                  <td className="px-6 py-3 text-sm border border-gray-200 dark:border-gray-700 min-w-48 whitespace-nowrap bg-gray-50/50 dark:bg-gray-800/30">
                                    {(() => {
                                      const reviewObj = (response as any).review || (reviewedBy[response.id] ? {
                                        status: reviewedBy[response.id]?.option || reviewedBy[response.id]?.status,
                                        reviewer: reviewedBy[response.id]?.name || reviewedBy[response.id]?.reviewer || 'Reviewer',
                                        flaggedQuestions: reviewedBy[response.id]?.flaggedQuestions || []
                                      } : null);

                                      return reviewObj ? (
                                        <div className="flex flex-col gap-1">
                                          <div className="flex items-center gap-2">
                                            <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold uppercase ${String(reviewObj.status).toLowerCase().trim() === 'accepted' ? 'bg-green-500/20 text-green-700 dark:text-green-400 border border-green-200 dark:border-green-800' :
                                              String(reviewObj.status).toLowerCase().trim() === 'rejected' ? 'bg-red-500/20 text-red-700 dark:text-red-400 border border-red-200 dark:border-red-800' :
                                                'bg-yellow-500/20 text-yellow-700 dark:text-yellow-400 border border-yellow-200 dark:border-yellow-800'
                                              }`}>
                                              {reviewObj.status}
                                            </span>
                                            <span className="text-xs text-gray-500 dark:text-gray-400">
                                              by <span className="font-semibold text-gray-700 dark:text-gray-300">{reviewObj.reviewer}</span>
                                            </span>
                                          </div>
                                          {reviewObj.flaggedQuestions && reviewObj.flaggedQuestions.length > 0 && (
                                            <div className="mt-1 flex flex-wrap gap-1">
                                              {reviewObj.flaggedQuestions.map((q: any, i: number) => (
                                                <span key={i} className="px-1.5 py-0.5 bg-gray-100 dark:bg-gray-800 text-[10px] text-gray-600 dark:text-gray-400 rounded border border-gray-200 dark:border-gray-700 line-clamp-1" title={q}>
                                                  {q}
                                                </span>
                                              ))}
                                            </div>
                                          )}
                                        </div>
                                      ) : (
                                        <span className="text-gray-400 italic text-xs">No review yet</span>
                                      );
                                    })()}
                                  </td>
                                  <td className="px-6 py-3 text-sm text-gray-600 dark:text-gray-400 font-medium border border-gray-200 dark:border-gray-700 min-w-40 whitespace-nowrap">
                                    {getResponseTimestamp(response)
                                      ? new Date(
                                        getResponseTimestamp(response)!,
                                      ).toLocaleString()
                                      : "-"}
                                  </td>
                                  <td className="px-6 py-3 text-sm text-center font-bold text-blue-600 dark:text-blue-400 border border-gray-200 dark:border-gray-700 whitespace-nowrap">
                                    {(() => {
                                      // Check both timeSpent (backend) and totalTimeSpent (frontend type)
                                      const timeSpent =
                                        response.timeSpent ??
                                        response.totalTimeSpent;
                                      return timeSpent !== undefined &&
                                        timeSpent !== null &&
                                        timeSpent > 0 ? (
                                        <div className="flex items-center justify-center gap-1">
                                          <Clock className="w-3.5 h-3.5" />
                                          {timeSpent > 60
                                            ? `${Math.floor(timeSpent / 60)}m ${timeSpent % 60}s`
                                            : `${timeSpent}s`}
                                        </div>
                                      ) : (
                                        "-"
                                      );
                                    })()}
                                  </td>
                                  {form?.sections?.map(
                                    (section: Section) =>
                                      selectedResponsesSectionIds.includes(
                                        section.id,
                                      ) &&
                                      section.questions?.map((q: any) => {
                                        const isFollowUp =
                                          q.parentId || q.showWhen?.questionId;
                                        const isEditing =
                                          editingResponseId === response.id;
                                        const hasCorrectAnswer =
                                          q.correctAnswer !== undefined;
                                        const answer = response.answers?.[q.id];

                                        let isCorrect = false;
                                        if (
                                          hasCorrectAnswer &&
                                          answer !== undefined &&
                                          answer !== null &&
                                          answer !== ""
                                        ) {
                                          const answerStr = Array.isArray(
                                            answer,
                                          )
                                            ? answer.join(", ").toLowerCase()
                                            : String(answer).toLowerCase();
                                          const correctStr = Array.isArray(
                                            q.correctAnswer,
                                          )
                                            ? q.correctAnswer
                                              .join(", ")
                                              .toLowerCase()
                                            : String(
                                              q.correctAnswer,
                                            ).toLowerCase();
                                          isCorrect = answerStr === correctStr;
                                        }

                                        return (
                                          <td
                                            key={`${response.id}-${q.id}`}
                                            className={`px-6 py-3 text-sm border border-gray-200 dark:border-gray-700 min-w-64 break-words ${isFollowUp
                                              ? "bg-purple-50 dark:bg-purple-900/10"
                                              : ""
                                              } ${hasCorrectAnswer && !isEditing
                                                ? isCorrect
                                                  ? "bg-green-100 dark:bg-green-900/30"
                                                  : "bg-red-100 dark:bg-red-900/30"
                                                : ""
                                              }`}
                                          >
                                            {isEditing ? (
                                              // Check if this is a file/image question
                                              (q.type === "file" || q.type === "image" ||
                                                (q.text && (q.text.toLowerCase().includes("image") || q.text.toLowerCase().includes("photo") || q.text.toLowerCase().includes("upload")))) ? (
                                                // File/Image upload inline component - WITHOUT hooks
                                                (() => {
                                                  const urls = getFileUrls(q.id);
                                                  const isUploading = uploadingFileId === q.id;

                                                  return (
                                                    <div className="space-y-2">
                                                      {/* Display existing images */}
                                                      {urls.length > 0 && (
                                                        <div className="flex flex-wrap gap-2">
                                                          {urls.map((url, index) => (
                                                            <div key={index} className="relative group">
                                                              <img
                                                                src={url}
                                                                alt={`Upload ${index + 1}`}
                                                                className="w-16 h-16 object-cover rounded-lg border border-gray-300 dark:border-gray-600 cursor-pointer"
                                                                onClick={() => window.open(url, '_blank')}
                                                              />
                                                              <button
                                                                onClick={() => handleRemoveImage(q.id, index)}
                                                                className="absolute -top-1 -right-1 p-1 bg-red-500 text-white rounded-full opacity-0 group-hover:opacity-100 transition-opacity"
                                                              >
                                                                <X className="w-3 h-3" />
                                                              </button>
                                                            </div>
                                                          ))}
                                                        </div>
                                                      )}

                                                      {/* Upload button */}
                                                      <label className="cursor-pointer inline-block">
                                                        <input
                                                          ref={(el) => { fileInputRefs.current[q.id] = el; }}
                                                          type="file"
                                                          className="hidden"
                                                          accept="image/*"
                                                          onChange={async (e) => {
                                                            const file = e.target.files?.[0];
                                                            if (file) await handleFileUpload(q.id, file);
                                                          }}
                                                        />
                                                        <div className={`flex items-center gap-1 px-3 py-1.5 rounded-lg border-2 border-dashed ${darkMode ? 'border-gray-600 hover:border-blue-400' : 'border-gray-300 hover:border-blue-500'} transition-colors cursor-pointer text-xs font-medium`}>
                                                          {isUploading ? (
                                                            <Loader2 className="w-4 h-4 animate-spin" />
                                                          ) : (
                                                            <Upload className="w-4 h-4" />
                                                          )}
                                                          {isUploading ? 'Uploading...' : 'Upload Image'}
                                                        </div>
                                                      </label>
                                                    </div>
                                                  );
                                                })()
                                              ) : q.type === "chassis-with-zone" || q.type === "chassis-without-zone" ||
                                                q.type === "zone-in" || q.type === "zone-out" || q.type === "chassis" ? (
                                                // For chassis/zone types, show status dropdown and remark
                                                <div className="space-y-1">
                                                  <select
                                                    value={editFormData[q.id]?.status || ''}
                                                    onChange={(e) => {
                                                      const currentVal = editFormData[q.id] || {};
                                                      setEditFormData({
                                                        ...editFormData,
                                                        [q.id]: { ...currentVal, status: e.target.value }
                                                      });
                                                    }}
                                                    className="w-full px-2 py-1 text-xs border border-blue-400 dark:border-blue-600 rounded bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                                                  >
                                                    <option value="">Select Status</option>
                                                    <option value="Accepted">Accepted</option>
                                                    <option value="Rejected">Rejected</option>
                                                    <option value="Rework">Rework</option>
                                                    <option value="Rework Completed">Rework Completed</option>
                                                    <option value="Verified">Verified</option>
                                                  </select>
                                                  <input
                                                    type="text"
                                                    value={editFormData[q.id]?.remark || ''}
                                                    onChange={(e) => {
                                                      const currentVal = editFormData[q.id] || {};
                                                      setEditFormData({
                                                        ...editFormData,
                                                        [q.id]: { ...currentVal, remark: e.target.value }
                                                      });
                                                    }}
                                                    placeholder="Remark"
                                                    className="w-full px-2 py-1 text-xs border border-blue-400 dark:border-blue-600 rounded bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                                                  />
                                                </div>
                                              ) : (
                                                // Text input for other types
                                                <input
                                                  type="text"
                                                  value={typeof editFormData[q.id] === 'object' && 'status' in editFormData[q.id]
                                                    ? editFormData[q.id].status
                                                    : (editFormData[q.id] ? (typeof editFormData[q.id] === 'string' ? editFormData[q.id] : JSON.stringify(editFormData[q.id], null, 2)) : "")}
                                                  onChange={(e) => {
                                                    const val = e.target.value;
                                                    if (editFormData[q.id] && typeof editFormData[q.id] === 'object' && 'status' in editFormData[q.id]) {
                                                      setEditFormData({
                                                        ...editFormData,
                                                        [q.id]: { ...editFormData[q.id], status: val },
                                                      });
                                                    } else {
                                                      let parsed;
                                                      try {
                                                        parsed = JSON.parse(val);
                                                      } catch {
                                                        parsed = val;
                                                      }
                                                      setEditFormData({
                                                        ...editFormData,
                                                        [q.id]: parsed,
                                                      });
                                                    }
                                                  }}
                                                  className="w-full px-2 py-1 text-xs border border-blue-400 dark:border-blue-600 rounded bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                                                  placeholder="Enter answer"
                                                />
                                              )
                                            ) : (
                                              // Display answer (existing code)
                                              <div className="flex flex-col gap-1 max-w-[250px] overflow-auto max-h-[250px]">
                                                {renderAnswerDisplay(answer, q)}
                                                {q.trackResponseRank && response.responseRanks?.[q.id] && (
                                                  <span className={`text-[10px] font-bold min-w-[24px] h-6 px-1.5 rounded-full flex items-center justify-center border shadow-sm w-fit mt-1 ${getRankStyle(answer, darkMode)}`}>
                                                    #{response.responseRanks[q.id]}
                                                  </span>
                                                )}
                                              </div>
                                            )}
                                          </td>
                                        );
                                      }),
                                  )}
                                </tr>
                              ),
                            )
                          ) : (
                            <tr>
                              <td
                                colSpan={
                                  9 +
                                  (form?.sections?.reduce(
                                    (acc: number, sec: Section) =>
                                      selectedResponsesSectionIds.includes(
                                        sec.id,
                                      )
                                        ? acc + (sec.questions?.length || 0)
                                        : acc,
                                    0,
                                  ) || 0)
                                }
                                className="px-4 py-6 text-center text-gray-500 dark:text-gray-400"
                              >
                                No responses yet
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </>
                ) : (
                  <div className="p-6 text-center text-gray-500 dark:text-gray-400">
                    Select at least one section to view responses
                  </div>
                )}
              </div>
            </div>
          )}
        </>
      )}

      {/* Cascading Filter Modal */}
      <CascadingFilterModal
        isOpen={showFilterModal}
        onClose={() => setShowFilterModal(false)}
        questions={allQuestionsWithSections}  // Pass all questions from all sections
        responses={responses}
        onApplyFilters={(filters) => {
          const { dates, locations, ...questionFilters } = filters as any;
          setCascadingFilters(questionFilters);
          if (dates) {
            setDateFilter({
              type: dates.startDate || dates.endDate ? "range" : "all",
              startDate: dates.startDate || "",
              endDate: dates.endDate || "",
            });
          }
          if (locations && locations.length > 0) {
            setLocationFilter(locations);
          }
        }}
      />

      {selectedResponse && selectedFormForModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white dark:bg-gray-900 rounded-lg max-w-4xl w-full my-8 max-h-[90vh] overflow-y-auto shadow-2xl">
            <div className="sticky top-0 bg-white dark:bg-gray-900 px-6 py-4 border-b border-gray-200 dark:border-gray-700 flex items-center justify-between z-10">
              <h2 className="text-xl font-bold text-gray-900 dark:text-white">
                Response Details
              </h2>
              <button
                onClick={() => {
                  setSelectedResponse(null);
                  setSelectedFormForModal(null);
                }}
                className="text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300"
              >
                <X className="w-6 h-6" />
              </button>
            </div>

            <div className="p-6 space-y-6">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <p className="text-sm font-semibold text-gray-600 dark:text-gray-400">
                    Form
                  </p>
                  <p className="text-gray-900 dark:text-white">
                    {selectedFormForModal?.title || "N/A"}
                  </p>
                </div>
                <div>
                  <p className="text-sm font-semibold text-gray-600 dark:text-gray-400">
                    Submitted
                  </p>
                  <p className="text-gray-900 dark:text-white">
                    {getResponseTimestamp(selectedResponse)
                      ? new Date(
                        getResponseTimestamp(selectedResponse)!,
                      ).toLocaleString()
                      : "N/A"}
                  </p>
                </div>
                <div>
                  <p className="text-sm font-semibold text-gray-600 dark:text-gray-400">
                    Time Taken
                  </p>
                  <p className="text-blue-600 dark:text-blue-400 font-bold flex items-center gap-1">
                    <Clock className="w-4 h-4" />
                    {(() => {
                      const timeSpent =
                        selectedResponse.timeSpent ??
                        selectedResponse.totalTimeSpent ??
                        0;
                      return timeSpent > 0
                        ? timeSpent > 60
                          ? `${Math.floor(timeSpent / 60)}m ${timeSpent % 60}s`
                          : `${timeSpent}s`
                        : "N/A";
                    })()}
                  </p>
                </div>
              </div>

              {selectedFormForModal?.sections?.map((section: Section) => (
                <div
                  key={section.id}
                  className="border border-gray-200 dark:border-gray-700 rounded-lg p-4"
                >
                  <h3 className="font-semibold text-lg text-gray-900 dark:text-white mb-4">
                    {section.title}
                  </h3>
                  <div className="space-y-4">
                    {section.questions?.map((question: any) => {
                      const answer = selectedResponse.answers?.[question.id];
                      return (
                        <div
                          key={question.id}
                          className="border-l-4 border-blue-300 dark:border-blue-700 pl-4"
                        >
                          <p className="text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">
                            {question.text}
                          </p>
                          <div className="text-gray-900 dark:text-gray-100 flex flex-col gap-1">
                            {hasAnswerValue(answer) ? (
                              renderAnswerDisplay(answer, question)
                            ) : (
                              <span className="text-gray-400">No response</span>
                            )}
                            {question.trackResponseRank &&
                              selectedResponse.responseRanks?.[question.id] && (
                                <span
                                  className={`text-[10px] font-bold min-w-[24px] h-6 px-1.5 rounded-full flex items-center justify-center border shadow-sm ${getRankStyle(answer, darkMode)}`}
                                >
                                  #{selectedResponse.responseRanks[question.id]}
                                </span>
                              )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}

              {selectedFormForModal?.followUpQuestions?.length > 0 && (
                <div className="border border-gray-200 dark:border-gray-700 rounded-lg p-4">
                  <h3 className="font-semibold text-lg text-gray-900 dark:text-white mb-4">
                    Follow-up Questions
                  </h3>
                  <div className="space-y-4">
                    {selectedFormForModal.followUpQuestions.map(
                      (question: any) => {
                        const answer = selectedResponse.answers?.[question.id];
                        return (
                          <div
                            key={question.id}
                            className="border-l-4 border-purple-300 dark:border-purple-700 pl-4"
                          >
                            <p className="text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">
                              {question.text}
                            </p>
                            <div className="text-gray-900 dark:text-gray-100 flex flex-col gap-1">
                              {hasAnswerValue(answer) ? (
                                renderAnswerDisplay(answer, question)
                              ) : (
                                <span className="text-gray-400">
                                  No response
                                </span>
                              )}
                              {question.trackResponseRank &&
                                selectedResponse.responseRanks?.[
                                question.id
                                ] && (
                                  <span
                                    className={`text-[10px] font-bold min-w-[24px] h-6 px-1.5 rounded-full flex items-center justify-center border shadow-sm ${getRankStyle(answer, darkMode)}`}
                                  >
                                    #
                                    {
                                      selectedResponse.responseRanks[
                                      question.id
                                      ]
                                    }
                                  </span>
                                )}
                            </div>
                          </div>
                        );
                      },
                    )}
                  </div>
                </div>
              )}
            </div>

            <div className="sticky bottom-0 bg-gray-50 dark:bg-gray-800 px-6 py-4 border-t border-gray-200 dark:border-gray-700 flex justify-end">
              <button
                onClick={() => {
                  setSelectedResponse(null);
                  setSelectedFormForModal(null);
                }}
                className="px-4 py-2 bg-gray-300 dark:bg-gray-700 text-gray-900 dark:text-white rounded-lg hover:bg-gray-400 dark:hover:bg-gray-600 transition-colors"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Comparison View - Last 3 Responses */}
      {analyticsView === "comparison" && (
        <div className="bg-gradient-to-br from-gray-50 to-gray-100 dark:from-gray-900 dark:to-gray-800 rounded-xl border border-gray-200 dark:border-gray-800">
          <div className="p-4 sm:p-6">
            {/* View Mode Tabs */}
            <div className="flex items-center justify-between mb-6">
              <div className="flex gap-1 bg-white dark:bg-gray-700 rounded-lg p-1 w-fit border border-gray-200 dark:border-gray-600">
                <button
                  onClick={() => setComparisonViewMode("dashboard")}
                  className={`flex items-center gap-2 px-3.5 py-1.5 text-sm font-medium rounded-md transition-all duration-200 ${comparisonViewMode === "dashboard"
                    ? "text-white shadow-sm"
                    : "text-gray-900 dark:text-gray-100 hover:text-black dark:hover:text-white"
                    }`}
                  style={{
                    backgroundColor:
                      comparisonViewMode === "dashboard"
                        ? "#1e3a8a"
                        : "transparent",
                  }}
                >
                  <BarChart3 className="w-4 h-4" />
                  Dashboard
                </button>
                <button
                  onClick={() => setComparisonViewMode("responses")}
                  className={`flex items-center gap-2 px-3.5 py-1.5 text-sm font-medium rounded-md transition-all duration-200 ${comparisonViewMode === "responses"
                    ? "text-white shadow-sm"
                    : "text-gray-900 dark:text-gray-100 hover:text-black dark:hover:text-white"
                    }`}
                  style={{
                    backgroundColor:
                      comparisonViewMode === "responses"
                        ? "#1e3a8a"
                        : "transparent",
                  }}
                >
                  <FileText className="w-4 h-4" />
                  Responses
                </button>
              </div>

              <div className="flex items-center gap-6 mx-4">
                <div className="text-center">
                  <h2 className="text-lg font-bold text-gray-900 dark:text-white leading-tight">
                    {form?.title}
                  </h2>
                  <p className="text-xs text-gray-600 dark:text-gray-400">
                    Last 5 Responses Comparison
                  </p>
                </div>
              </div>
            </div>

            {/* Content Area */}
            {comparisonViewMode === "dashboard" ? (
              <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-5 gap-4">
                {(() => {
                  const last5 = filteredResponses
                    .filter((r) => getResponseTimestamp(r))
                    .sort((a, b) => {
                      const dateA = new Date(
                        getResponseTimestamp(a)!,
                      ).getTime();
                      const dateB = new Date(
                        getResponseTimestamp(b)!,
                      ).getTime();
                      return dateB - dateA;
                    })
                    .slice(0, 5);

                  if (last5.length === 0) {
                    return (
                      <div className="col-span-full flex flex-col items-center justify-center min-h-64 py-12">
                        <UsersIcon className="w-12 h-12 text-gray-300 dark:text-gray-600 mb-4" />
                        <p className="text-gray-600 dark:text-gray-400 font-medium">
                          No responses to compare
                        </p>
                      </div>
                    );
                  }

                  return last5.map((response, idx) => {
                    const sectionStats = getSectionYesNoStats(
                      form,
                      response.answers || {},
                    );
                    const filteredSectionStats = sectionStats.filter(
                      (stat) =>
                        stat.yes > 0 ||
                        stat.no > 0 ||
                        stat.na > 0 ||
                        (stat.accepted && stat.accepted > 0) ||
                        (stat.rejected && stat.rejected > 0) ||
                        (stat.rework && stat.rework > 0),
                    );

                    const totalQuestions = filteredSectionStats.reduce(
                      (sum, stat) => sum + stat.total,
                      0,
                    );
                    const totalYes = filteredSectionStats.reduce(
                      (sum, stat) => sum + stat.yes + (stat.accepted || 0),
                      0,
                    );
                    const totalNo = filteredSectionStats.reduce(
                      (sum, stat) => sum + stat.no + (stat.rejected || 0),
                      0,
                    );
                    const totalNA = filteredSectionStats.reduce(
                      (sum, stat) => sum + stat.na + (stat.rework || 0),
                      0,
                    );
                    const totalAnswered = totalYes + totalNo + totalNA;

                    const overallScore =
                      totalQuestions > 0
                        ? ((totalYes / totalQuestions) * 100).toFixed(1)
                        : "0.0";
                    const responseRate =
                      totalQuestions > 0
                        ? ((totalAnswered / totalQuestions) * 100).toFixed(1)
                        : "0.0";
                    const yesPercent =
                      totalAnswered > 0
                        ? ((totalYes / totalAnswered) * 100).toFixed(1)
                        : "0.0";
                    const noPercent =
                      totalAnswered > 0
                        ? ((totalNo / totalAnswered) * 100).toFixed(1)
                        : "0.0";
                    const naPercent =
                      totalAnswered > 0
                        ? ((totalNA / totalAnswered) * 100).toFixed(1)
                        : "0.0";

                    return (
                      <div
                        key={response.id}
                        className="border border-gray-200 dark:border-gray-700 rounded-lg overflow-hidden shadow-sm hover:shadow-md transition-shadow bg-white dark:bg-gray-800 flex flex-col h-full"
                      >
                        <div className="bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-gray-700 dark:to-gray-600 px-4 py-3 border-b border-gray-200 dark:border-gray-700">
                          <div className="flex flex-col items-center text-center">
                            <p className="text-xs font-semibold text-blue-700 dark:text-blue-400 uppercase mb-1">
                              Submission #{idx + 1}
                            </p>
                            <p className="text-xs text-gray-600 dark:text-gray-400">
                              {getResponseTimestamp(response)
                                ? new Date(
                                  getResponseTimestamp(response)!,
                                ).toLocaleDateString("en-US", {
                                  month: "short",
                                  day: "numeric",
                                  hour: "2-digit",
                                  minute: "2-digit",
                                })
                                : "N/A"}
                            </p>
                            <p className="text-2xl font-bold text-blue-900 dark:text-blue-300 mt-2">
                              {overallScore}%
                            </p>
                            <p className="text-xs font-semibold text-blue-700 dark:text-blue-400 uppercase mt-1">
                              Overall Score
                            </p>
                          </div>
                        </div>

                        <div className="p-4 space-y-3 flex-1">
                          <div className="grid grid-cols-3 gap-2">
                            <div className="bg-indigo-50 dark:bg-indigo-900/20 p-2 rounded border border-indigo-200 dark:border-indigo-700 text-center">
                              <p className="text-xs font-semibold text-indigo-700 dark:text-indigo-400 uppercase">
                                Sections
                              </p>
                              <p className="text-xl font-bold text-indigo-900 dark:text-indigo-300">
                                {filteredSectionStats.length}
                              </p>
                            </div>
                            <div className="bg-green-50 dark:bg-green-900/20 p-2 rounded border border-green-200 dark:border-green-700 text-center">
                              <p className="text-xs font-semibold text-green-700 dark:text-green-400 uppercase">
                                Rate
                              </p>
                              <p className="text-xl font-bold text-green-900 dark:text-green-300">
                                {responseRate}%
                              </p>
                            </div>
                            <div className="bg-purple-50 dark:bg-purple-900/20 p-2 rounded border border-purple-200 dark:border-purple-700 text-center">
                              <p className="text-xs font-semibold text-purple-700 dark:text-purple-400 uppercase">
                                Questions
                              </p>
                              <p className="text-xl font-bold text-purple-900 dark:text-purple-300">
                                {totalQuestions}
                              </p>
                            </div>
                          </div>

                          <div className="border-t border-gray-200 dark:border-gray-700 pt-3">
                            <p className="text-xs font-semibold text-gray-900 dark:text-white mb-2 text-center">
                              Distribution
                            </p>
                            <div className="space-y-1">
                              <div className="text-center p-2 bg-green-100/60 dark:bg-green-900/20 rounded border border-green-200 dark:border-green-700">
                                <p className="text-xs font-semibold text-green-700 dark:text-green-400">
                                  {complianceLabels.yes}
                                </p>
                                <p className="text-sm font-bold text-green-800 dark:text-green-300">
                                  {totalYes} ({yesPercent}%)
                                </p>
                              </div>
                              <div className="text-center p-2 bg-red-100/60 dark:bg-red-900/20 rounded border border-red-200 dark:border-red-700">
                                <p className="text-xs font-semibold text-red-700 dark:text-red-400">
                                  {complianceLabels.no}
                                </p>
                                <p className="text-sm font-bold text-red-800 dark:text-red-300">
                                  {totalNo} ({noPercent}%)
                                </p>
                              </div>
                              <div className="text-center p-2 bg-yellow-100/60 dark:bg-yellow-900/20 rounded border border-yellow-200 dark:border-yellow-700">
                                <p className="text-xs font-semibold text-yellow-700 dark:text-yellow-400">
                                  {complianceLabels.na}
                                </p>
                                <p className="text-sm font-bold text-yellow-800 dark:text-yellow-300">
                                  {totalNA} ({naPercent}%)
                                </p>
                              </div>
                            </div>
                          </div>

                          {filteredSectionStats.length > 0 && (
                            <div className="border-t border-gray-200 dark:border-gray-700 pt-3">
                              <p className="text-xs font-semibold text-gray-900 dark:text-white mb-3">
                                Sections
                              </p>
                              <div className="space-y-4">
                                {filteredSectionStats.map((row) => {
                                  const rowYes = row.yes + (row.accepted || 0);
                                  const rowNo = row.no + (row.rejected || 0);
                                  const rowNA = row.na + (row.rework || 0);
                                  const total = rowYes + rowNo + rowNA;
                                  const yesPercent =
                                    total > 0
                                      ? ((rowYes / total) * 100).toFixed(1)
                                      : 0;
                                  const noPercent =
                                    total > 0
                                      ? ((rowNo / total) * 100).toFixed(1)
                                      : 0;
                                  const naPercent =
                                    total > 0
                                      ? ((rowNA / total) * 100).toFixed(1)
                                      : 0;

                                  const chartData = {
                                    labels: [
                                      `${complianceLabels.yes} (${yesPercent}%)`,
                                      `${complianceLabels.no} (${noPercent}%)`,
                                      `${complianceLabels.na} (${naPercent}%)`,
                                    ],
                                    datasets: [
                                      {
                                        data: [rowYes, rowNo, rowNA],
                                        backgroundColor: [
                                          "#1e3a8a",
                                          "#3b82f6",
                                          "#93c5fd",
                                        ],
                                        borderColor: [
                                          "#1e3a8a",
                                          "#3b82f6",
                                          "#93c5fd",
                                        ],
                                        borderWidth: 2,
                                        borderRadius: 4,
                                      },
                                    ],
                                  };

                                  return (
                                    <div
                                      key={row.id}
                                      className="p-3 bg-gradient-to-br from-gray-50 to-gray-100 dark:from-gray-700/40 dark:to-gray-800/40 rounded-lg border border-gray-200 dark:border-gray-600"
                                    >
                                      <p className="font-semibold text-gray-900 dark:text-white text-[11px] mb-3">
                                        {row.title}
                                      </p>

                                      <div className="flex gap-3">
                                        <div className="flex-1 flex items-center justify-center">
                                          <div className="w-24 h-24">
                                            <Doughnut
                                              data={chartData}
                                              options={{
                                                responsive: true,
                                                maintainAspectRatio: true,
                                                plugins: {
                                                  legend: {
                                                    display: false,
                                                  },
                                                  tooltip: {
                                                    backgroundColor:
                                                      "rgba(0, 0, 0, 0.8)",
                                                    titleColor: "#ffffff",
                                                    bodyColor: "#ffffff",
                                                    borderColor: "#ffffff",
                                                    borderWidth: 1,
                                                    callbacks: {
                                                      label: (context) => {
                                                        return `${context.label}: ${context.parsed}`;
                                                      },
                                                    },
                                                  },
                                                  datalabels: {
                                                    color: "#ffffff",
                                                    font: {
                                                      weight: "bold",
                                                      size: 10,
                                                    },
                                                    formatter: (
                                                      value,
                                                      context,
                                                    ) => {
                                                      const total =
                                                        context.dataset.data.reduce(
                                                          (a, b) => (a as number) + (b as number),
                                                          0,
                                                        );
                                                      const percentage = (
                                                        ((value as number) / (total as number)) *
                                                        100
                                                      ).toFixed(0);
                                                      return `${percentage}%`;
                                                    },
                                                  },
                                                },
                                              }}
                                            />
                                          </div>
                                        </div>

                                        <div className="flex-1 flex flex-col justify-center gap-2 text-xs">
                                          <div className="flex items-center gap-2">
                                            <div
                                              className="w-3 h-3 rounded-full"
                                              style={{
                                                backgroundColor: "#1e3a8a",
                                              }}
                                            ></div>
                                            <span className="text-gray-700 dark:text-gray-300">
                                              {complianceLabels.yes}:{" "}
                                              <span className="font-bold">
                                                {rowYes}
                                              </span>{" "}
                                              ({yesPercent}%)
                                            </span>
                                          </div>
                                          <div className="flex items-center gap-2">
                                            <div
                                              className="w-3 h-3 rounded-full"
                                              style={{
                                                backgroundColor: "#3b82f6",
                                              }}
                                            ></div>
                                            <span className="text-gray-700 dark:text-gray-300">
                                              {complianceLabels.no}:{" "}
                                              <span className="font-bold">
                                                {rowNo}
                                              </span>{" "}
                                              ({noPercent}%)
                                            </span>
                                          </div>
                                          <div className="flex items-center gap-2">
                                            <div
                                              className="w-3 h-3 rounded-full"
                                              style={{
                                                backgroundColor: "#93c5fd",
                                              }}
                                            ></div>
                                            <span className="text-gray-700 dark:text-gray-300">
                                              {complianceLabels.na}:{" "}
                                              <span className="font-bold">
                                                {rowNA}
                                              </span>{" "}
                                              ({naPercent}%)
                                            </span>
                                          </div>
                                          <div className="border-t border-gray-300 dark:border-gray-500 mt-2 pt-2">
                                            <p className="font-semibold text-gray-900 dark:text-white">
                                              Total: <span>{total}</span>
                                            </p>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}

                          {response.submissionMetadata?.location && (
                            <div className="border-t border-gray-200 dark:border-gray-700 pt-3">
                              <p className="text-xs font-semibold text-gray-900 dark:text-white mb-1">
                                Location
                              </p>
                              <p className="text-xs text-gray-700 dark:text-gray-300 truncate">
                                {response.submissionMetadata.location.city ||
                                  response.submissionMetadata.location.region ||
                                  response.submissionMetadata.location
                                    .country ||
                                  "N/A"}
                              </p>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
            ) : (
              <div className="card p-6">
                {filteredResponses.length === 0 ? (
                  <div className="flex flex-col items-center justify-center min-h-96 py-12">
                    <UsersIcon className="w-12 h-12 text-gray-300 dark:text-gray-600 mb-4" />
                    <p className="text-gray-600 dark:text-gray-400 font-medium">
                      No responses to compare
                    </p>
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-gray-200 dark:border-gray-700">
                    <table className="w-full text-sm border-collapse">
                      <thead>
                        <tr className="bg-indigo-50 dark:bg-indigo-900/20">
                          <th className="sticky left-0 z-20 text-left px-4 py-3 font-semibold text-gray-900 dark:text-white border border-gray-200 dark:border-gray-700 min-w-40 bg-indigo-50 dark:bg-indigo-900/20">
                            Question
                          </th>
                          {filteredResponses
                            .filter((r) => getResponseTimestamp(r))
                            .sort((a, b) => {
                              const dateA = new Date(
                                getResponseTimestamp(a)!,
                              ).getTime();
                              const dateB = new Date(
                                getResponseTimestamp(b)!,
                              ).getTime();
                              return dateB - dateA;
                            })
                            .slice(0, 5)
                            .map((response, idx) => (
                              <th
                                key={response.id}
                                className="text-center px-3 py-2 font-semibold text-gray-900 dark:text-white border border-gray-200 dark:border-gray-700 min-w-28 bg-gradient-to-b from-blue-50 to-indigo-50 dark:from-blue-900/30 dark:to-indigo-900/30"
                              >
                                <div className="flex flex-col gap-0.5">
                                  <span className="text-xs text-gray-600 dark:text-gray-400 leading-tight font-medium">
                                    Sub #{idx + 1}
                                  </span>
                                  <span className="text-xs font-semibold text-gray-600 dark:text-gray-400 leading-tight">
                                    {getResponseTimestamp(response)
                                      ? new Date(
                                        getResponseTimestamp(response)!,
                                      ).toLocaleDateString("en-US", {
                                        month: "short",
                                        day: "numeric",
                                      })
                                      : "N/A"}
                                  </span>
                                </div>
                              </th>
                            ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-200 dark:divide-gray-700">
                        {form?.sections?.flatMap((section) =>
                          section.questions?.map((question, qIdx) => {
                            const last5Responses = filteredResponses
                              .filter((r) => getResponseTimestamp(r))
                              .sort((a, b) => {
                                const dateA = new Date(
                                  getResponseTimestamp(a)!,
                                ).getTime();
                                const dateB = new Date(
                                  getResponseTimestamp(b)!,
                                ).getTime();
                                return dateB - dateA;
                              })
                              .slice(0, 5);

                            return (
                              <tr
                                key={question.id}
                                className={
                                  qIdx % 2 === 0
                                    ? "bg-white dark:bg-gray-900"
                                    : "bg-gray-50 dark:bg-gray-800/50"
                                }
                              >
                                <td className="sticky left-0 z-10 px-4 py-3 font-medium text-gray-900 dark:text-white border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 min-w-60">
                                  <div className="flex flex-col">
                                    <span className="text-sm font-semibold break-words whitespace-normal">
                                      {question.text || "Question"}
                                    </span>
                                    {question.description && (
                                      <span className="text-xs text-gray-500 dark:text-gray-400 mt-1 break-words whitespace-normal">
                                        {question.description}
                                      </span>
                                    )}
                                  </div>
                                </td>
                                {last5Responses.map((response) => {
                                  const answer =
                                    response.answers?.[question.id];
                                  const hasAnswer =
                                    answer !== null &&
                                    answer !== undefined &&
                                    answer !== "";

                                  return (
                                    <td
                                      key={`${response.id}-${question.id}`}
                                      className="text-center px-3 py-2 border border-gray-200 dark:border-gray-700 min-w-[120px]"
                                    >
                                      {hasAnswer ? (
                                        <div className="flex flex-col items-center justify-center max-w-[200px] overflow-auto max-h-[150px] gap-1">
                                          {renderAnswerDisplay(
                                            answer,
                                            question,
                                          )}
                                          {response.responseRanks?.[
                                            question.id
                                          ] && (
                                              <span
                                                className={`text-[10px] font-bold min-w-[24px] h-6 px-1.5 rounded-full flex items-center justify-center border shadow-sm ${getRankStyle(answer, darkMode)}`}
                                              >
                                                #
                                                {
                                                  response.responseRanks[
                                                  question.id
                                                  ]
                                                }
                                              </span>
                                            )}
                                        </div>
                                      ) : (
                                        <span className="text-xs text-gray-400 dark:text-gray-500">
                                          —
                                        </span>
                                      )}
                                    </td>
                                  );
                                })}
                              </tr>
                            );
                          }),
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {analyticsView === "opsTable" && (
        <OPSMasterListTable
          form={form}
          filteredResponses={filteredResponses}
          responseStatuses={responseStatuses}
          getResponseTimestamp={getResponseTimestamp}
        />
      )}
      {/* Delete Confirmation Modal */}
      {showDeleteConfirm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 rounded-lg shadow-2xl max-w-sm">
            <div className="p-6">
              <div className="flex items-center justify-center w-12 h-12 mx-auto bg-red-100 dark:bg-red-900/30 rounded-full mb-4">
                <Trash2 className="w-6 h-6 text-red-600 dark:text-red-400" />
              </div>
              <h3 className="text-lg font-semibold text-gray-900 dark:text-white text-center mb-2">
                Delete Response
              </h3>
              <p className="text-sm text-gray-600 dark:text-gray-400 text-center mb-6">
                Are you sure you want to delete this response? This action
                cannot be undone.
              </p>
              <div className="flex gap-2 justify-center">
                <button
                  onClick={() => {
                    setShowDeleteConfirm(false);
                    setDeletingResponseId(null);
                  }}
                  disabled={isDeleting}
                  className="px-4 py-2 text-gray-700 dark:text-gray-300 bg-gray-100 dark:bg-gray-700 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-600 transition-colors font-medium disabled:opacity-50"
                >
                  Cancel
                </button>
                <button
                  onClick={handleDeleteResponse}
                  disabled={isDeleting}
                  className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors font-medium disabled:opacity-50 flex items-center gap-2"
                >
                  {isDeleting ? (
                    <>
                      <div className="animate-spin rounded-full h-4 w-4 border-2 border-white border-t-transparent"></div>
                      Deleting...
                    </>
                  ) : (
                    "Delete"
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Bulk Delete Confirmation Modal */}
      {showBulkDeleteConfirm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 rounded-lg shadow-2xl max-w-sm">
            <div className="p-6">
              <div className="flex items-center justify-center w-12 h-12 mx-auto bg-red-100 dark:bg-red-900/30 rounded-full mb-4">
                <Trash2 className="w-6 h-6 text-red-600 dark:text-red-400" />
              </div>
              <h3 className="text-lg font-semibold text-gray-900 dark:text-white text-center mb-2">
                Delete Selected Responses
              </h3>
              <p className="text-sm text-gray-600 dark:text-gray-400 text-center mb-2">
                Are you sure you want to delete {selectedResponseIds.length}{" "}
                response(s)? This action cannot be undone.
              </p>
              <p className="text-xs text-gray-500 dark:text-gray-500 text-center mb-6">
                This will permanently remove the selected responses from the
                system.
              </p>
              <div className="flex gap-2 justify-center">
                <button
                  onClick={() => {
                    setShowBulkDeleteConfirm(false);
                  }}
                  disabled={isDeleting}
                  className="px-4 py-2 text-gray-700 dark:text-gray-300 bg-gray-100 dark:bg-gray-700 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-600 transition-colors font-medium disabled:opacity-50"
                >
                  Cancel
                </button>
                <button
                  onClick={handleBulkDeleteResponses}
                  disabled={isDeleting}
                  className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors font-medium disabled:opacity-50 flex items-center gap-2"
                >
                  {isDeleting ? (
                    <>
                      <div className="animate-spin rounded-full h-4 w-4 border-2 border-white border-t-transparent"></div>
                      Deleting...
                    </>
                  ) : (
                    <>Delete {selectedResponseIds.length} Response(s)</>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Action Menu Modal */}
      {showActionMenuModal && actionResponse && (
        <div className="fixed inset-0 bg-black/60 dark:bg-black/75 flex items-center justify-center p-4 z-[120] backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden transform animate-in zoom-in-95 duration-200">
            <div className="p-5 border-b border-gray-100 dark:border-gray-700 flex items-center justify-between">
              <h3 className="text-lg font-bold text-gray-900 dark:text-white">
                Response Actions
              </h3>
              <button
                onClick={() => {
                  setShowActionMenuModal(false);
                  setActionResponse(null);
                }}
                className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-full transition-colors text-gray-500"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 space-y-2">
              <div className="px-2 py-1 mb-2">
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Selected Response</p>
                <p className="text-sm font-bold text-gray-700 dark:text-gray-300 truncate">
                  {actionResponse.answers?.chassis_number || actionResponse.submittedBy || "Anonymous"}
                </p>
              </div>

              {(() => {
                const isOwnTenant = user?.role === 'superadmin' ||
                  user?.role === 'admin' ||
                  user?.role === 'subadmin' ||
                  user?.role === 'inspector' ||
                  !responseTenantId ||
                  (currentUserTenantId && responseTenantId.toString() === currentUserTenantId.toString());

                const isActualOwnTenant = user?.role === 'superadmin' ||
                  !responseTenantId ||
                  (currentUserTenantId && responseTenantId.toString() === currentUserTenantId.toString());

                return (
                  <>
                    {/* Focus It */}
                    <button
                      onClick={() => {
                        handleOpenModal(actionResponse);
                        setShowActionMenuModal(false);
                      }}
                      className="w-full flex items-center gap-3 p-3.5 hover:bg-blue-50 dark:hover:bg-blue-900/20 text-gray-700 dark:text-gray-200 rounded-xl transition-colors font-semibold text-sm"
                    >
                      <div className="p-2 bg-blue-100 dark:bg-blue-900/40 rounded-lg text-blue-600 dark:text-blue-400">
                        <Maximize className="w-4 h-4" />
                      </div>
                      FOCUS IT
                    </button>

                    {/* View Details */}
                    {isOwnTenant && (
                      <button
                        onClick={() => {
                          handleViewDetails(actionResponse);
                          setShowActionMenuModal(false);
                        }}
                        className="w-full flex items-center gap-3 p-3.5 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 text-gray-700 dark:text-gray-200 rounded-xl transition-colors font-semibold text-sm"
                      >
                        <div className="p-2 bg-indigo-100 dark:bg-indigo-900/40 rounded-lg text-indigo-600 dark:text-indigo-400">
                          <Eye className="w-4 h-4" />
                        </div>
                        View Full Details
                      </button>
                    )}

                    {/* Chat / Review */}
                    {actionResponse.isDispatched && (
                      <button
                        onClick={() => {
                          setChatResponse(actionResponse);
                          setShowChatModal(true);
                          setSelectedReviewOptions(prev => ({ ...prev, [actionResponse.id]: '' }));
                          setReviewedBy(prev => ({ ...prev, [actionResponse.id]: null }));
                          setShowActionMenuModal(false);
                        }}
                        className="w-full flex items-center gap-3 p-3.5 hover:bg-blue-50 dark:hover:bg-blue-900/20 text-gray-700 dark:text-gray-200 rounded-xl transition-colors font-semibold text-sm"
                      >
                        <div className="p-2 bg-blue-100 dark:bg-blue-900/40 rounded-lg text-blue-600 dark:text-blue-400">
                          <MessageCircle className="w-4 h-4" />
                        </div>
                        Review & Discussion
                      </button>
                    )}

                    {/* Edit - Admin only */}
                    {!isGuest && (user?.role === "superadmin" || user?.role === "admin") && isActualOwnTenant && (
                      <button
                        onClick={() => {
                          handleEditStart(actionResponse);
                          setShowActionMenuModal(false);
                        }}
                        className="w-full flex items-center gap-3 p-3.5 hover:bg-amber-50 dark:hover:bg-amber-900/20 text-gray-700 dark:text-gray-200 rounded-xl transition-colors font-semibold text-sm"
                      >
                        <div className="p-2 bg-amber-100 dark:bg-amber-900/40 rounded-lg text-amber-600 dark:text-amber-400">
                          <Edit className="w-4 h-4" />
                        </div>
                        Edit Response
                      </button>
                    )}

                    {/* Delete - Admin only */}
                    {!isGuest && (user?.role === "superadmin" || user?.role === "admin") && isActualOwnTenant && (
                      <button
                        onClick={() => {
                          setDeletingResponseId(actionResponse.id);
                          setShowDeleteConfirm(true);
                          setShowActionMenuModal(false);
                        }}
                        className="w-full flex items-center gap-3 p-3.5 hover:bg-red-50 dark:hover:bg-red-900/20 text-red-600 dark:text-red-400 rounded-xl transition-colors font-semibold text-sm"
                      >
                        <div className="p-2 bg-red-100 dark:bg-red-900/40 rounded-lg text-red-600 dark:text-red-400">
                          <Trash2 className="w-4 h-4" />
                        </div>
                        Delete Response
                      </button>
                    )}
                  </>
                );
              })()}
            </div>

            <div className="p-4 bg-gray-50 dark:bg-gray-800/50 border-t border-gray-100 dark:border-gray-700 text-center">
              <button
                onClick={() => {
                  setShowActionMenuModal(false);
                  setActionResponse(null);
                }}
                className="text-xs font-bold text-gray-500 uppercase tracking-widest hover:text-gray-700 dark:hover:text-gray-300 transition-colors"
              >
                Close Menu
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Share Analytics Modal */}
      <ShareAnalyticsModal
        isOpen={shareAnalyticsModal.open}
        onClose={() =>
          setShareAnalyticsModal((prev) => ({ ...prev, open: false }))
        }
        formId={shareAnalyticsModal.formId}
        formTitle={shareAnalyticsModal.formTitle}
        analyticsData={fullAnalyticsData}
      />

      {/* Auto Send Modal */}
      <AutoSendModal
        isOpen={autoSendModal.open}
        onClose={() =>
          setAutoSendModal((prev) => ({ ...prev, open: false }))
        }
        formId={autoSendModal.formId}
        formTitle={autoSendModal.formTitle}
      />

      {/* Toast Notification */}
      {toast && (
        <div
          className={`fixed bottom-4 right-4 px-6 py-3 rounded-lg shadow-lg text-white font-medium z-50 animate-fadeIn ${toast.type === "success"
            ? "bg-green-500 dark:bg-green-600"
            : toast.type === "info"
              ? "bg-blue-500 dark:bg-blue-600"
              : "bg-red-500 dark:bg-red-600"
            }`}
        >
          <div className="flex items-center gap-2">
            {toast.type === "success" ? (
              <CheckCircle className="w-5 h-5" />
            ) : toast.type === "info" ? (
              <Info className="w-5 h-5" />
            ) : (
              <XCircle className="w-5 h-5" />
            )}
            {toast.message}
          </div>
        </div>
      )}

      {/* Chat Modal */}
      {showChatModal && chatResponse && (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-[60] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-900 rounded-xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="px-6 py-4 border-b border-gray-200 dark:border-gray-700 bg-indigo-600">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-white/20 rounded-lg">
                    <MessageCircle className="w-5 h-5 text-white" />
                  </div>
                  <div>
                    <h2 className="text-lg font-bold text-white">
                      Question Filter: {chatResponse.submittedBy || 'Inspector'}
                    </h2>
                    <p className="text-xs text-white/70">Chassis: {(() => {
                      const chassisQ = form?.sections?.flatMap(s => s.questions || []).find(q => q.type === 'chassis' || q.type === 'chassisWithZone' || q.type === 'chassisWithoutZone' || q.text?.toLowerCase().includes('chassis'));
                      const chassisVal = chatResponse.answers?.[chassisQ?.id || ''];
                      if (typeof chassisVal === 'object' && chassisVal?.chassisNumber) {
                        return chassisVal.chassisNumber;
                      }
                      if (typeof chassisVal === 'string' && chassisVal.trim()) {
                        return chassisVal;
                      }
                      return 'N/A';
                    })()}</p>
                  </div>
                </div>

                {/* Review Options - Right side of header */}
                <div className="flex items-center gap-3">

                  {/* === AFTER REVIEW: Show review badge + score === */}
                  {(() => {
                    const responseId = chatResponse?.id || '';

                    // First check if review is directly attached to chatResponse (immediate display)
                    let reviewer = null;
                    let reviewOption = '';

                    if (chatResponse?.review) {
                      reviewer = chatResponse.review.reviewer;
                      reviewOption = chatResponse.review.option || chatResponse.review.reviewOption;
                    } else {
                      // Fallback to state-based approach
                      reviewer = reviewedBy[responseId];
                      reviewOption = selectedReviewOptions[responseId];
                    }

                    // Accept reviewer object with either .id or ._id field
                    const hasReviewer = reviewer && (reviewer.id || (reviewer as any)._id || reviewer.name || reviewer.email);
                    const hasOption = reviewOption && reviewOption !== '';



                    if (!hasReviewer || !hasOption) return null;

                    const reviewerName = reviewer.name || reviewer.email || user?.name || user?.email || 'Reviewer';
                    const isAccepted = reviewOption === 'Accepted';
                    const isRejected = reviewOption === 'Rejected';
                    const emoji = isAccepted ? '✅' : isRejected ? '❌' : '🔄';
                    const badgeClass = isAccepted
                      ? 'bg-green-500/20 border-green-400 text-green-100'
                      : isRejected
                        ? 'bg-red-500/20 border-red-400 text-red-100'
                        : 'bg-yellow-500/20 border-yellow-400 text-yellow-100';

                    const scoreVal = performanceScores[chatResponse?.submittedBy || ''] ??
                      performanceScores[(chatResponse?.createdBy as any)?._id || chatResponse?.createdBy as string || ''];

                    return (
                      <div className={`flex items-center gap-2 px-3 py-1.5 text-xs font-bold rounded-lg border ${badgeClass}`}>
                        <span>
                          <span className="font-bold">Status:</span>{' '}
                          {emoji} {reviewOption} by {reviewerName}
                        </span>

                      </div>
                    );
                  })()}

                  {/* === BEFORE REVIEW: Show 3 buttons or "pending" state === */}
                  {!reviewSubmitted[`${String(user?._id)}-${String(chatResponse?.id)}`] &&
                    !reviewedBy[chatResponse?.id || ''] &&
                    (responseStatuses[chatResponse?.id] === "Direct Ok" ||
                      responseStatuses[chatResponse?.id] === "Rework Accepted" ||
                      responseStatuses[chatResponse?.id] === "Accepted" ||
                      responseStatuses[chatResponse?.id] === "Pending Review" ||
                      responseStatuses[chatResponse?.id] === "Rework Completed") &&
                    (() => {
                      const userEmail = user?.email || "";
                      const userUsername = user?.username || "";
                      const userId = user?._id || user?.id;
                      const userIdStr = userId ? String(userId) : "";
                      const creatorId = typeof chatResponse?.createdBy === "object"
                        ? (chatResponse.createdBy as any)?._id || (chatResponse.createdBy as any)?.id
                        : chatResponse?.createdBy;
                      const creatorIdStr = creatorId ? String(creatorId) : "";

                      const isSubmitter = chatResponse?.submittedBy === userEmail ||
                        chatResponse?.submittedBy === userUsername ||
                        chatResponse?.submitterContact?.email === userEmail ||
                        (creatorIdStr && creatorIdStr === userIdStr);
                      return !isSubmitter;
                    })() && (
                      pendingReviewOption ? (
                        /* Pending state: show label + cancel */
                        <div className="flex items-center gap-2">
                          <span className={`px-3 py-1 text-xs font-bold rounded ${pendingReviewOption === 'Rejected' ? 'bg-red-500/30 text-red-100 border border-red-400' : 'bg-yellow-500/30 text-yellow-100 border border-yellow-400'
                            }`}>
                            {pendingReviewOption === 'Rejected' ? '❌' : '🔄'} {pendingReviewOption} — Select questions below
                          </span>
                          <button
                            onClick={() => setPendingReviewOption(null)}
                            className="px-2 py-1 text-xs font-bold rounded bg-white/10 text-white hover:bg-white/20 transition-all"
                          >
                            Cancel
                          </button>
                        </div>
                      ) : (
                        /* Normal state: 3 review buttons - Only show if user is not the submitter */
                        (() => {
                          const userEmail = user?.email || "";
                          const userUsername = user?.username || "";
                          const userIdStr = user?._id ? String(user._id) : (user?.id ? String(user.id) : "");
                          const creatorId = typeof chatResponse?.createdBy === "object"
                            ? (chatResponse.createdBy as any)?._id || (chatResponse.createdBy as any)?.id
                            : chatResponse?.createdBy;
                          const creatorIdStr = creatorId ? String(creatorId) : "";

                          const isSubmitter = chatResponse?.submittedBy === userEmail ||
                            chatResponse?.submittedBy === userUsername ||
                            chatResponse?.submitterContact?.email === userEmail || (creatorIdStr && creatorIdStr === userIdStr);

                          return !isSubmitter ? (
                            <div className="flex gap-2">
                              {['Accepted', 'Rejected', 'Rework'].map((option) => (
                                <button
                                  key={option}
                                  onClick={() => {
                                    if (option === 'Accepted') {
                                      handleReviewSubmit(chatResponse!.id, option);
                                    } else {
                                      setPendingReviewOption(option);
                                    }
                                  }}
                                  className={`px-3 py-1 text-xs font-bold rounded transition-all border ${option === 'Accepted'
                                    ? 'bg-green-500/30 border-green-400 text-green-100 hover:bg-green-500/50'
                                    : option === 'Rejected'
                                      ? 'bg-red-500/30 border-red-400 text-red-100 hover:bg-red-500/50'
                                      : 'bg-yellow-500/30 border-yellow-400 text-yellow-100 hover:bg-yellow-500/50'
                                    }`}
                                >
                                  {option === 'Accepted' ? '✅' : option === 'Rejected' ? '❌' : '🔄'} {option}
                                </button>
                              ))}
                            </div>
                          ) : null;
                        })()
                      )
                    )}

                  <button
                    onClick={() => {
                      const responseId = searchParams.get('responseId');
                      if (responseId) {
                        navigate('/inspector/chat');
                      } else {
                        setShowChatModal(false);
                        setChatResponse(null);
                        setPendingReviewOption(null);
                        setSelectedReviewOptions(prev => ({ ...prev, [chatResponse?.id || '']: '' }));
                        setReviewedBy(prev => ({ ...prev, [chatResponse?.id || '']: null }));
                      }
                    }}
                    className="p-2 hover:bg-white/10 rounded-full text-white transition-colors"
                    title="Close Chat"
                  >
                    <X className="w-6 h-6" />
                  </button>
                </div>
              </div>
            </div>

            <div className="flex-1 overflow-hidden grid grid-cols-1 md:grid-cols-2">
              {/* Left Column: Filters */}
              <div className="p-6 bg-gray-50 dark:bg-gray-900/50 border-r border-gray-200 dark:border-gray-700 overflow-y-auto space-y-6">
                <div>
                  <div className="space-y-4">
                    <p className="text-xl text-gray-700 dark:text-gray-300">Chassis Number : {(() => {
                      const chassisQ = form?.sections?.flatMap(s => s.questions || []).find(q => q.type === 'chassis' || q.type === 'chassisWithZone' || q.type === 'chassisWithoutZone' || q.text?.toLowerCase().includes('chassis'));
                      const chassisVal = chatResponse.answers?.[chassisQ?.id || ''];
                      if (typeof chassisVal === 'object' && chassisVal?.chassisNumber) {
                        return chassisVal.chassisNumber;
                      }
                      if (typeof chassisVal === 'string' && chassisVal.trim()) {
                        return chassisVal;
                      }
                      return 'N/A';
                    })()}</p>

                    {/* Show question selection panel when Rejected/Rework is pending */}
                    {pendingReviewOption && (
                      <>
                        <div className="space-y-2">
                          <div className={`flex items-center gap-2 mb-3 px-3 py-2 rounded-lg ${pendingReviewOption === 'Rejected' ? 'bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800' : 'bg-yellow-50 dark:bg-yellow-900/20 border border-yellow-200 dark:border-yellow-800'}`}>
                            <span className="text-lg">{pendingReviewOption === 'Rejected' ? '❌' : '🔄'}</span>
                            <div>
                              <p className={`text-xs font-bold ${pendingReviewOption === 'Rejected' ? 'text-red-700 dark:text-red-300' : 'text-yellow-700 dark:text-yellow-300'}`}>
                                {pendingReviewOption} Review
                              </p>
                              <p className="text-xs text-gray-500 dark:text-gray-400">Select the questions with issues, fill in corrections, then click "Send & Submit Review"</p>
                            </div>
                          </div>
                          <label className="text-xs font-semibold text-gray-700 dark:text-gray-300 ml-1">
                            Select Questions to Flag
                          </label>
                          <div className="max-h-[400px] overflow-y-auto border border-gray-200 dark:border-gray-700 rounded-xl shadow-inner scrollbar-thin scrollbar-thumb-gray-300 dark:scrollbar-thumb-gray-600 bg-white dark:bg-gray-800">
                            {form?.sections?.flatMap(s =>
                              (s.questions || []).filter((q: any) => !q.parentId && !q.showWhen?.questionId)
                            ).map(q => (
                              <div key={q.id} className="border-b border-gray-100 dark:border-gray-700 last:border-0">
                                <div className="flex items-start gap-3 p-3 hover:bg-indigo-50/50 dark:hover:bg-indigo-900/10 cursor-default transition-colors group">
                                  <label className="mt-1 cursor-pointer">
                                    <input
                                      type="checkbox"
                                      checked={chatFilters.questions.includes(q.id)}
                                      onChange={(e) => {
                                        const checked = e.target.checked;
                                        setChatFilters(prev => ({
                                          ...prev,
                                          questions: checked
                                            ? [...prev.questions, q.id]
                                            : prev.questions.filter(id => id !== q.id)
                                        }));
                                      }}
                                      className="w-4 h-4 text-indigo-600 border-gray-300 rounded focus:ring-indigo-500 cursor-pointer"
                                    />
                                  </label>
                                  <div className="flex-1">
                                    <span className="text-sm text-gray-700 dark:text-gray-300 font-medium group-hover:text-indigo-600 dark:group-hover:text-indigo-400 transition-colors block mb-1">
                                      {q.text}
                                    </span>
                                    {chatFilters.questions.includes(q.id) && (
                                      <div className="mt-2 space-y-2">
                                        <QuestionSuggestionRenderer
                                          question={q}
                                          currentAnswer={responses.find(r => r.id === chatResponse.id)?.answers?.[q.id]}
                                          value={chatFilters.suggestedAnswers?.[q.id] || {}}
                                          onChange={(val) => setChatFilters(prev => ({
                                            ...prev,
                                            suggestedAnswers: { ...prev.suggestedAnswers, [q.id]: val }
                                          }))}
                                        />
                                      </div>
                                    )}
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="pt-6 flex items-center justify-between">
                          <button
                            onClick={() => setChatFilters({ chassisNumber: "", location: "", questions: [], selectedCategories: {}, zoneType: "both", suggestedAnswers: {} })}
                            className="text-xs font-bold text-indigo-600 dark:text-indigo-400 hover:underline flex items-center gap-1"
                          >
                            Clear All Filters
                          </button>
                          <div className="flex gap-2">
                            <button
                              onClick={() => setShowChatModal(false)}
                              className="px-4 py-2 text-xs font-bold text-gray-500 hover:bg-gray-200 dark:hover:bg-gray-800 rounded-lg transition-colors"
                            >
                              Cancel
                            </button>
                            <button
                              onClick={() => setShowChatModal(false)}
                              className="px-4 py-2 text-xs font-extrabold text-white bg-indigo-600 rounded-lg hover:bg-indigo-700 shadow-lg shadow-indigo-200 dark:shadow-none transition-all active:scale-95"
                            >
                              Apply Filters
                            </button>
                          </div>
                        </div>
                        <p className="text-[11px] text-indigo-600 dark:text-indigo-400 font-semibold mt-2">
                          ✏️ Type your message below and click <b>Send Feedback</b> to submit the review.
                        </p>
                      </>)}

                  </div>
                </div>
              </div>
              {/* Right Column: Chat history and input */}
              <div className="flex-1 flex flex-col bg-white dark:bg-gray-900 border-t md:border-t-0 p-6 overflow-hidden">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-xs font-bold text-gray-500 dark:text-gray-400 uppercase tracking-widest">
                    Message Center
                  </h3>
                  <div className="flex items-center gap-1.5 px-2 py-0.5 bg-green-100 dark:bg-green-900/30 rounded-full border border-green-200 dark:border-green-800">
                    <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse" />
                    <span className="text-[10px] font-bold text-green-700 dark:text-green-400">Live Context</span>
                  </div>
                </div>

                <div className="flex-1 bg-gray-50 dark:bg-gray-800/20 rounded-2xl border border-gray-100 dark:border-gray-800 p-4 mb-4 overflow-y-auto space-y-4 flex flex-col scrollbar-thin scrollbar-thumb-gray-200 dark:scrollbar-thumb-gray-700">
                  {chatMessages.length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center text-gray-400 gap-3 opacity-50">
                      <div className="p-4 bg-gray-100 dark:bg-gray-800 rounded-full ring-8 ring-gray-50 dark:ring-gray-900/50">
                        <MessageCircle className="w-10 h-10" />
                      </div>
                      <div className="text-center">
                        <p className="text-sm font-bold text-gray-600 dark:text-gray-300">No active conversation</p>
                        <p className="text-xs">Send a message to start the thread.</p>
                      </div>
                    </div>
                  ) : (
                    chatMessages.map((msg, i) => (
                      <div key={i} className={`flex flex-col ${String(msg.from?._id || msg.from) === String(user?._id || (user as any)?.id) ? 'items-end' : 'items-start'} animate-in fade-in slide-in-from-bottom-2 duration-300`}>
                        <div className={`max-w-[90%] px-4 py-2.5 rounded-2xl text-sm leading-relaxed shadow-sm ${String(msg.from?._id || msg.from) === String(user?._id || (user as any)?.id)
                          ? 'bg-[#dcf8c6] text-gray-900 rounded-br-lg rounded-tr-lg rounded-tl-sm'
                          : 'bg-white dark:bg-gray-100 text-gray-900 border border-gray-100 dark:border-gray-700 rounded-bl-lg rounded-tl-lg rounded-tr-sm'
                          }`}>
                          {msg.questionContexts && msg.questionContexts.length > 0 ? (
                            <div className="space-y-3">
                              {msg.questionContexts.map((ctx: any, idx: number) => (
                                <div key={idx} className="space-y-2">
                                  <p className="text-[12px] font-bold text-gray-500 dark:text-gray-400 border-b border-indigo-100 dark:border-indigo-800/50 pb-0.5">
                                    {ctx.title}
                                  </p>

                                  {ctx.suggestion && (
                                    <div className="mt-1 p-2 bg-amber-50 dark:bg-amber-900/20 rounded-lg border border-amber-100 dark:border-amber-800">
                                      <p className="text-[10px] font-bold text-amber-600 uppercase mb-1">Review Feedback:</p>
                                      {ctx.question ? (
                                        <QuestionSuggestionRenderer
                                          question={ctx.question}
                                          value={ctx.suggestion}
                                          currentAnswer={ctx.answer}
                                          onChange={(newSuggestion) => {
                                            // Update the suggestion in chatFilters
                                            setChatFilters(prev => ({
                                              ...prev,
                                              suggestedAnswers: {
                                                ...prev.suggestedAnswers,
                                                [ctx.question.id]: newSuggestion
                                              }
                                            }));
                                          }}
                                        />
                                      ) : (
                                        // Fallback to read-only display if question data is missing
                                        <div className="text-xs text-gray-600 dark:text-gray-400">
                                          {renderAnswerDisplay(ctx.suggestion, { type: 'text' } as any)}
                                        </div>
                                      )}
                                    </div>
                                  )}
                                </div>
                              ))}
                            </div>
                          ) : msg.questionTitles && msg.questionTitles.length > 0 && (
                            <div className="mb-2 p-2 bg-indigo-50/50 dark:bg-indigo-900/20 rounded-xl border border-indigo-100/50 dark:border-indigo-800/30">
                              <p className="text-[10px] uppercase font-black text-indigo-500 dark:text-indigo-400 mb-1.5 flex items-center gap-1">
                                <Filter className="w-2.5 h-2.5" />
                                Linked Questions
                              </p>
                              <div className="flex flex-wrap gap-1">
                                {msg.questionTitles.map((title: string, idx: number) => (
                                  <span key={idx} className="px-1.5 py-0.5 bg-white dark:bg-gray-700 text-[9px] font-bold text-indigo-600 dark:text-indigo-300 rounded-md border border-indigo-100 dark:border-indigo-800">
                                    {title}
                                  </span>
                                ))}
                              </div>
                            </div>
                          )}
                          {renderMessageWithImages(msg.message)}
                        </div>
                        <div className="flex items-center gap-1 mt-1.5 px-1 opacity-60">
                          <span className="text-[9px] font-medium text-gray-500 dark:text-gray-400">
                            {String(msg.from?._id || msg.from) === String(user?._id || (user as any)?.id) ? 'You' : (msg.from?.name || 'Inspector')} • {new Date(msg.createdAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                          </span>
                          {String(msg.from?._id || msg.from) !== String(user?._id || (user as any)?.id) && (
                            <button
                              onClick={() => {
                                setNewMessage(`Replying to: "${msg.message.substring(0, 30)}..." \n`);
                                const textarea = document.querySelector('textarea');
                                if (textarea) textarea.focus();
                              }}
                              className="flex items-center gap-1 text-[10px] font-bold text-indigo-600 hover:underline ml-2 pointer-events-auto"
                            >
                              <Reply className="w-3 h-3" />
                              Reply
                            </button>
                          )}
                        </div>
                      </div>
                    ))
                  )}
                </div>

                {(() => {
                  const userEmail = user?.email || "";
                  const userUsername = user?.username || "";
                  const userIdStr = user?._id ? String(user._id) : (user?.id ? String(user.id) : "");
                  const creatorId = typeof chatResponse?.createdBy === "object"
                    ? (chatResponse.createdBy as any)?._id || (chatResponse.createdBy as any)?.id
                    : chatResponse?.createdBy;
                  const creatorIdStr = creatorId ? String(creatorId) : "";

                  const isSubmitter = chatResponse?.submittedBy === userEmail ||
                    chatResponse?.submittedBy === userUsername ||
                    chatResponse?.submitterContact?.email === userEmail || (creatorIdStr && creatorIdStr === userIdStr);

                  // For submitters, only show message input if no review option is pending
                  if (isSubmitter && pendingReviewOption) {
                    return null;
                  }

                  return (
                    <div className="space-y-3">
                      <div className="relative">
                        <textarea
                          value={newMessage}
                          onChange={(e) => setNewMessage(e.target.value)}
                          placeholder={isSubmitter ? "Send a message..." : "Type your feedback to the inspector..."}
                          className="w-full px-5 py-4 bg-gray-50 dark:bg-gray-800 border-2 border-transparent focus:border-indigo-500 dark:focus:border-indigo-400 rounded-2xl text-sm focus:ring-0 transition-all resize-none shadow-inner text-gray-800 dark:text-gray-200"
                          rows={3}
                        />
                      </div>
                      <button
                        onClick={async () => {
                          // If a Rejected/Rework review is pending and user is not submitter, send message + submit review together
                          if (pendingReviewOption && !isSubmitter) {
                            setPendingReviewOption(null); // Clear pending state immediately
                            const reviewNote = newMessage.trim() || `Please review and correct the flagged questions (${pendingReviewOption}).`;
                            await handleSendMessage(reviewNote);
                            await handleReviewSubmit(chatResponse!.id, pendingReviewOption);
                          } else {
                            await handleSendMessage();
                          }
                        }}
                        disabled={isSendingMessage || !newMessage.trim()}
                        className={`w-full flex items-center justify-center gap-2 px-6 py-3.5 text-white text-sm font-black rounded-2xl disabled:opacity-50 disabled:cursor-not-allowed transition-all active:scale-[0.98] shadow-xl ${pendingReviewOption && !isSubmitter
                          ? (pendingReviewOption === 'Rejected'
                            ? 'bg-red-600 hover:bg-red-700 shadow-red-200 dark:shadow-none'
                            : 'bg-yellow-600 hover:bg-yellow-700 shadow-yellow-200 dark:shadow-none')
                          : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-200 dark:shadow-none'
                          }`}
                      >
                        {isSendingMessage ? (
                          <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                        ) : (
                          <>
                            <span>{pendingReviewOption && !isSubmitter ? `Send Feedback & Submit ${pendingReviewOption}` : 'Send Message'}</span>
                            <ChevronRight className="w-4 h-4" />
                          </>
                        )}
                      </button>
                      <div className="flex justify-center gap-2">
                        <p className="text-[10px] text-center text-gray-400 font-medium">
                          Message will be sent to <b>{isSubmitter ? 'the reviewer' : (chatResponse.submittedBy || 'the submitter')}</b>
                        </p>
                      </div>
                    </div>
                  );
                })()}


              </div>
            </div>
          </div>
        </div>
      )}
      {/* Confirm Update Popup for Edit Response */}
      {showEditConfirmPopup && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[10000] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl p-6 max-w-md w-full border border-gray-200 dark:border-gray-700 animate-in fade-in zoom-in duration-300">
            <h3 className="text-xl font-bold text-gray-900 dark:text-white mb-4 flex items-center gap-2">
              <Calendar className="w-5 h-5 text-blue-600" />
              Confirm Update
            </h3>

            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                  Date (DD/MM/YY)
                </label>
                <input
                  type="text"
                  value={popupDate}
                  onChange={(e) => setPopupDate(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  placeholder="e.g. 15/06/26"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1">
                  Issuance / Revision Details
                </label>
                <input
                  type="text"
                  value={popupIssuanceDetails}
                  onChange={(e) => setPopupIssuanceDetails(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  placeholder="Enter revision details"
                />
              </div>
            </div>

            <div className="mt-6 flex justify-end gap-3">
              <button
                onClick={() => {
                  setShowEditConfirmPopup(false);
                  setPendingEditAnswers(null);
                }}
                className="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleConfirmEditSave}
                disabled={isSaving}
                className="px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition-colors flex items-center gap-2 disabled:opacity-50"
              >
                {isSaving ? (
                  <>
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Saving...
                  </>
                ) : (
                  <>
                    <Save className="w-4 h-4" />
                    Confirm Update
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}