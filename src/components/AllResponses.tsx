import React, { useState, useEffect, useMemo, useRef } from "react";
import { useNavigate } from "react-router-dom";
import {
  Eye,
  Calendar,
  FileText,
  User,
  X,
  CheckCircle,
  Clock,
  XCircle,
  Download,
  Trash2,
  Edit2,
  TrendingUp,
  BarChart3,
  PieChart,
  Activity,
  Zap,
  Target,
  Award,
  Users,
  FileCheck,
  AlertTriangle,
  Save,
  ChevronDown,
  MapPin,
  List,
  RefreshCw,
  Upload,
  Mail,
  MessageCircle,
  MessageSquare,
} from "lucide-react";
import { Bar, Line, Pie, Radar } from "react-chartjs-2";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
  ArcElement,
  RadialLinearScale,
} from "chart.js";
import { apiClient } from "../api/client";
import { formatTimestamp } from "../utils/dateUtils";
import { useNotification } from "../context/NotificationContext";
import { useLogo } from "../context/LogoContext";
import { useTheme } from "../context/ThemeContext";
import { generateResponseExcelReport } from "../utils/responseExportUtils";
import { generateAndDownloadPDF, exportAllResponsesToZip } from "../utils/pdfExportUtils";
import JSZip from "jszip";
import FilePreview from "./FilePreview";
import ResponseEdit from "./ResponseEdit";
import DashboardSummaryCard from "./DashboardSummaryCard";
import AnswerTemplateImport from "./AnswerTemplateImport";
import { isImageUrl } from "../utils/answerTemplateUtils";
import ImageLink from "./ImageLink";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
  ArcElement,
  RadialLinearScale
);

function formatSectionLabel(label: string, maxLength = 20): string {
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
}

interface Form {
  _id: string;
  id?: string;
  title: string;
  description?: string;
  parentFormId?: string;
  sections?: any[];
  followUpQuestions?: any[];
}

interface LocationData {
  country?: string;
  countryCode?: string;
  region?: string;
  city?: string;
  latitude?: number;
  longitude?: number;
  timezone?: string;
  isp?: string;
}

interface SubmissionMetadata {
  ipAddress?: string;
  userAgent?: string;
  browser?: string;
  device?: string;
  os?: string;
  location?: LocationData;
  submittedAt?: string;
  source?: string;
}

interface Response {
  _id: string;
  id: string;
  questionId: string;
  formId?: string;
  parentResponseId?: string;
  answers: Record<string, any>;
  responseRanks?: Record<string, number>;
  createdAt: string;
  updatedAt: string;
  assignedTo?: string;
  status?: string;
  yesNoScore?: {
    yes: number;
    total: number;
  };
  submissionMetadata?: SubmissionMetadata;
  dealerName?: string;
}

interface GroupedResponse {
  modelNo: string;
  responses: (Response & { formTitle: string; dealerName?: string })[];
  submissionDate: string;
}

type SectionStat = {
  id: string;
  title: string;
  yes: number;
  no: number;
  na: number;
  total: number;
};

export default function AllResponses() {
  const { darkMode } = useTheme();
  const navigate = useNavigate();
  const { showSuccess, showError, showConfirm } = useNotification();
  const { logo } = useLogo();

  const [responses, setResponses] = useState<(Response & { formTitle: string; dealerName?: string })[]>([]);
  const [groupedResponsesList, setGroupedResponsesList] = useState<GroupedResponse[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedResponse, setSelectedResponse] = useState<(Response & { formTitle: string }) | null>(null);
  const [selectedForm, setSelectedForm] = useState<Form | null>(null);
  const [formLoading, setFormLoading] = useState(false);
  const [updatingStatus, setUpdatingStatus] = useState(false);
  const [showStatusUpdate, setShowStatusUpdate] = useState(false);
  const [viewMode, setViewMode] = useState<"dashboard" | "responses">("dashboard");
  const [pendingSectionId, setPendingSectionId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedFormIds, setSelectedFormIds] = useState<string[]>([]);
  const [showFormFilter, setShowFormFilter] = useState(false);
  const [exportingExcel, setExportingExcel] = useState(false);
  const [deletingResponseId, setDeletingResponseId] = useState<string | null>(null);
  const [deletingGroupModelNo, setDeletingGroupModelNo] = useState<string | null>(null);
  const [editingResponse, setEditingResponse] = useState<(Response & { formTitle: string }) | null>(null);
  const [editingForm, setEditingForm] = useState<Form | null>(null);
  const [savingEdit, setSavingEdit] = useState(false);
  const [editingFormLoading, setEditingFormLoading] = useState(false);
  const [generatingPDF, setGeneratingPDF] = useState(false);
  const [exportingZip, setExportingZip] = useState(false);
  const [isAnswerTemplateOpen, setIsAnswerTemplateOpen] = useState(false);
  const [selectedPDFType, setSelectedPDFType] = useState<any>(null);
  const [selectedSectionId, setSelectedSectionId] = useState<string | null>(null);
  const [sectionResponsesMap, setSectionResponsesMap] = useState<Record<string, any[]>>({});
  const [sectionChartTypes, setSectionChartTypes] = useState<Record<string, "pie" | "bar">>({});
  const [expandResponseRateBreakdown, setExpandResponseRateBreakdown] = useState(false);
  const sectionRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const [showDownloadOptions, setShowDownloadOptions] = useState(false);
  const [showResponseDropdown, setShowResponseDropdown] = useState(false);
  const [formsMap, setFormsMap] = useState<Map<string, Form>>(new Map());

  useEffect(() => {
    fetchData();
  }, []);

  // Helper to get answer from response
  const getAnswer = (response: Response, questionId: string): string => {
    const answer = response.answers?.[questionId];
    if (!answer) return "";
    if (typeof answer === "string") return answer;
    if (typeof answer === "object") {
      if (answer.status) return answer.status;
      if (answer.chassisNumber) return answer.chassisNumber;
      return JSON.stringify(answer);
    }
    return String(answer);
  };

  // Add this function to merge answers from multiple responses in a group
  const getMergedAnswers = (responses: (Response & { formTitle: string })[]) => {
    const mergedMap = new Map<string, { questionText: string; values: Set<string>; questionId: string }>();

    responses.forEach((response) => {
      const form = formsMap.get(response.questionId || response.formId || '');

      if (form?.sections) {
        for (const section of form.sections) {
          for (const question of section.questions || []) {
            const answer = response.answers?.[question.id];
            if (answer && answer !== "" && answer !== null) {
              const questionText = question.text || question.label || question.id;
              const answerStr = typeof answer === "object" ? JSON.stringify(answer) : String(answer);

              if (!mergedMap.has(question.id)) {
                mergedMap.set(question.id, {
                  questionId: question.id,
                  questionText: questionText,
                  values: new Set()
                });
              }
              mergedMap.get(question.id)!.values.add(answerStr);
            }
          }
        }
      }

      // Add any unmapped answers
      for (const [key, value] of Object.entries(response.answers || {})) {
        if (value && !mergedMap.has(key) && !key.startsWith("__")) {
          const answerStr = typeof value === "object" ? JSON.stringify(value) : String(value);
          mergedMap.set(key, {
            questionId: key,
            questionText: key,
            values: new Set([answerStr])
          });
        } else if (value && !key.startsWith("__")) {
          const answerStr = typeof value === "object" ? JSON.stringify(value) : String(value);
          if (mergedMap.has(key)) {
            mergedMap.get(key)!.values.add(answerStr);
          }
        }
      }
    });

    // Convert to array and format values
    return Array.from(mergedMap.values()).map(item => ({
      ...item,
      displayValue: item.values.size === 1
        ? Array.from(item.values)[0]
        : Array.from(item.values).join(" | ")
    }));
  };

  // Add this function to check if a question is the Model question
  const isModelQuestion = (questionId: string, questionText: string): boolean => {
    return questionId === "q_basic_model" ||
      questionText.toLowerCase().includes("model") ||
      questionId.toLowerCase().includes("model");
  };

  const getModelNo = (response: Response, form: Form | undefined): string => {
    if (!form) return "N/A";

    // Direct check for q_basic_model in answers
    if (response.answers?.["q_basic_model"]) {
      const modelValue = response.answers["q_basic_model"];
      if (modelValue && modelValue !== "") {
        return String(modelValue);
      }
    }

    // Also check for other possible model question IDs
    const possibleModelIds = ["q_basic_model", "model", "modelNo", "model_number"];
    for (const id of possibleModelIds) {
      if (response.answers?.[id]) {
        const value = response.answers[id];
        if (value && value !== "") {
          return String(value);
        }
      }
    }

    // Search through form sections for model question
    const headerSection = form.sections?.find((s: any) => s.id === "sec_basic_info");
    if (headerSection?.questions) {
      for (const question of headerSection.questions) {
        const questionText = (question.text || question.label || "").toLowerCase();
        if (questionText.includes("model") || question.id === "q_basic_model") {
          const answer = response.answers?.[question.id];
          if (answer && answer !== "") {
            return String(answer);
          }
        }
      }
    }

    return "N/A";
  };

  const fetchData = async () => {
    try {
      setLoading(true);
      const [responsesData, formsData] = await Promise.all([
        apiClient.getResponses(),
        apiClient.getForms(),
      ]);

      const formsMapTemp = new Map<string, Form>();
      formsData.forms.forEach((form: any) => {
        if (form?._id) formsMapTemp.set(form._id, form as Form);
        if (form?.id) formsMapTemp.set(form.id, form as Form);
      });
      setFormsMap(formsMapTemp);

      // Pre-calculate dealer question IDs for each form
      const dealerQuestionMap = new Map<string, string>();

      formsMapTemp.forEach((form: Form) => {
        const formId = form._id || form.id;
        if (!formId) return;

        if (form.sections && form.sections.length > 0) {
          const firstSection = form.sections[0];
          if (firstSection.questions && firstSection.questions.length > 0) {
            for (const question of firstSection.questions) {
              const questionText = (question.text || question.label || '').toLowerCase();
              const isDealerField = questionText.includes('dealer') ||
                questionText.includes('distributor') ||
                questionText.includes('agent') ||
                questionText.includes('store') ||
                questionText.includes('business');

              if (isDealerField) {
                dealerQuestionMap.set(formId, question.id);
                break;
              }
            }
          }
        }
      });

      const extractDealerName = (response: Response, form: Form | undefined): { name: string | null, rank: number | null } => {
        if (!form || !response.answers) return { name: null, rank: null };

        const formId = form._id || form.id;
        if (!formId) return { name: null, rank: null };

        const dealerQuestionId = dealerQuestionMap.get(formId);
        if (dealerQuestionId) {
          const answer = response.answers[dealerQuestionId];
          if (answer && hasAnswerValue(answer)) {
            return {
              name: String(answer),
              rank: response.responseRanks?.[dealerQuestionId] || null
            };
          }
        }

        if (form.sections && form.sections.length > 0) {
          const firstSection = form.sections[0];
          if (firstSection.questions && firstSection.questions.length > 0) {
            for (const question of firstSection.questions) {
              const answer = response.answers[question.id];
              if (answer && hasAnswerValue(answer)) {
                return {
                  name: String(answer),
                  rank: response.responseRanks?.[question.id] || null
                };
              }
            }
          }
        }

        return { name: null, rank: null };
      };

      const responsesWithTitles = responsesData.responses.map(
        (response: Response) => {
          const form = formsMapTemp.get(response.questionId || response.formId || '');
          const dealerInfo = extractDealerName(response, form);

          return {
            ...response,
            formTitle: form?.title || "Unknown Form",
            dealerName: dealerInfo.name || "Unknown",
            dealerRank: dealerInfo.rank,
          };
        }
      );

      setResponses(responsesWithTitles);

      // Group responses by Model No
      const groupedMap = new Map<string, GroupedResponse>();

      responsesWithTitles.forEach((response: Response & { formTitle: string }) => {
        const form = formsMapTemp.get(response.questionId || response.formId || '');
        const modelNo = getModelNo(response, form);

        if (groupedMap.has(modelNo)) {
          const existing = groupedMap.get(modelNo)!;
          existing.responses.push(response);
          if (new Date(response.createdAt) > new Date(existing.submissionDate)) {
            existing.submissionDate = response.createdAt;
          }
        } else {
          groupedMap.set(modelNo, {
            modelNo: modelNo,
            responses: [response],
            submissionDate: response.createdAt
          });
        }
      });

      const groupedArray = Array.from(groupedMap.values()).sort((a, b) =>
        new Date(b.submissionDate).getTime() - new Date(a.submissionDate).getTime()
      );

      setGroupedResponsesList(groupedArray);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load responses");
    } finally {
      setLoading(false);
    }
  };

  const handleViewDetails = (group: GroupedResponse) => {
    console.log("=== NAVIGATING WITH GROUPED RESPONSES ===");
    console.log("Group Model No:", group.modelNo);
    console.log("Number of responses:", group.responses.length);

    // Store the grouped responses in sessionStorage as a backup
    const storageKey = `grouped_${group.responses[0]._id}`;
    sessionStorage.setItem(storageKey, JSON.stringify({
      groupedResponses: group.responses,
      isGrouped: group.responses.length > 1,
      modelNo: group.modelNo
    }));

    // Navigate with state
    navigate(`/responses/${group.responses[0]._id}`, {
      state: {
        groupedResponses: group.responses,
        isGrouped: group.responses.length > 1,
        modelNo: group.modelNo,
        storageKey: storageKey
      }
    });
  };

  const handleEditResponse = async (response: Response & { formTitle: string }) => {
    setEditingResponse(response);
    setEditingForm(null);
    setEditingFormLoading(true);
    try {
      const formIdentifier = response.questionId || response.formId;
      const formData = await apiClient.getForm(formIdentifier);
      const form = formData.form;
      setEditingForm(form);
    } catch (err) {
      console.error("Failed to load form for editing:", err);
      showError("Failed to load form for editing.");
      setEditingResponse(null);
    } finally {
      setEditingFormLoading(false);
    }
  };

  const handleCloseEdit = () => {
    setEditingResponse(null);
    setEditingForm(null);
    setSavingEdit(false);
    setEditingFormLoading(false);
  };

  const handleSaveEditedResponse = async (updated: any) => {
    if (savingEdit) return;
    setSavingEdit(true);
    try {
      await apiClient.updateResponse(updated.id, { answers: updated.answers });
      fetchData();
      showSuccess("Response updated successfully.");
      handleCloseEdit();
    } catch (err) {
      console.error("Failed to update response:", err);
      showError("Failed to update response.");
    } finally {
      setSavingEdit(false);
    }
  };

  const handleDeleteResponse = (response: Response & { formTitle: string }) => {
    showConfirm(
      "Are you sure you want to delete this request?",
      async () => {
        setDeletingResponseId(response.id);
        try {
          await apiClient.deleteResponse(response.id);
          fetchData();
          showSuccess("Request deleted successfully.");
        } catch (err) {
          console.error("Failed to delete response:", err);
          showError("Failed to delete request.");
        } finally {
          setDeletingResponseId(null);
        }
      },
      "Delete Request",
      "Delete",
      "Cancel"
    );
  };

  const handleDeleteGroup = (group: GroupedResponse) => {
    const count = group.responses.length;
    showConfirm(
      `Delete all ${count} response${count !== 1 ? 's' : ''} for Model No: "${group.modelNo}"? This cannot be undone.`,
      async () => {
        setDeletingGroupModelNo(group.modelNo);
        try {
          // Delete all responses in the group in parallel
          await Promise.all(
            group.responses.map((r) => apiClient.deleteResponse(r._id || r.id))
          );
          fetchData();
          showSuccess(`All ${count} response${count !== 1 ? 's' : ''} for "${group.modelNo}" deleted successfully.`);
        } catch (err) {
          console.error("Failed to delete group responses:", err);
          showError("Failed to delete some responses. Please try again.");
        } finally {
          setDeletingGroupModelNo(null);
        }
      },
      "Delete All Responses",
      "Delete All",
      "Cancel"
    );
  };

  const hasAnswerValue = (value: any) => {
    if (value === null || value === undefined) return false;
    if (typeof value === "string") return value.trim() !== "";
    if (Array.isArray(value)) return value.length > 0;
    if (typeof value === "object") return Object.keys(value).length > 0;
    return true;
  };

  const renderAnswerDisplay = (value: any, question?: any): React.ReactNode => {
    if (!value) return <span className="text-gray-400">No response</span>;
    if (typeof value === "string") {
      if (isImageUrl(value)) return <ImageLink text={value} />;
      return value;
    }
    if (typeof value === "object") {
      if (value.status) return value.status;
      if (value.chassisNumber) return `Chassis: ${value.chassisNumber}`;
      if (value.remark) return value.remark;
      if (value.url && isImageUrl(value.url)) return <ImageLink text={value.url} />;
      return JSON.stringify(value);
    }
    return String(value);
  };

  const getRankStyle = (answer: any, darkMode: boolean = false) => {
    if (!answer) return "";
    const str = String(answer).trim().toLowerCase();
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const colors = [
      { l: "bg-blue-50 text-blue-700 border-blue-200", d: "bg-blue-900/30 text-blue-300 border-blue-800" },
      { l: "bg-emerald-50 text-emerald-700 border-emerald-200", d: "bg-emerald-900/30 text-emerald-300 border-emerald-800" },
      { l: "bg-amber-50 text-amber-700 border-amber-200", d: "bg-amber-900/30 text-amber-300 border-amber-800" },
      { l: "bg-purple-50 text-purple-700 border-purple-200", d: "bg-purple-900/30 text-purple-300 border-purple-800" },
      { l: "bg-pink-50 text-pink-700 border-pink-200", d: "bg-pink-900/30 text-pink-300 border-pink-800" },
      { l: "bg-indigo-50 text-indigo-700 border-indigo-200", d: "bg-indigo-900/30 text-indigo-300 border-indigo-800" },
      { l: "bg-teal-50 text-teal-700 border-teal-200", d: "bg-teal-900/30 text-teal-300 border-teal-800" },
      { l: "bg-cyan-50 text-cyan-700 border-cyan-200", d: "bg-cyan-900/30 text-cyan-300 border-cyan-800" }
    ];
    const color = colors[Math.abs(hash) % colors.length];
    return darkMode ? color.d : color.l;
  };

  const getStatusInfo = (status: string) => {
    switch (status?.toLowerCase()) {
      case "pending": return { color: "text-yellow-600", bgColor: "bg-yellow-50", icon: Clock, label: "Pending" };
      case "confirmed": return { color: "text-blue-600", bgColor: "bg-blue-50", icon: CheckCircle, label: "Confirmed" };
      case "verified": return { color: "text-green-600", bgColor: "bg-green-50", icon: CheckCircle, label: "Verified" };
      case "rejected": return { color: "text-red-600", bgColor: "bg-red-50", icon: XCircle, label: "Rejected" };
      default: return { color: "text-gray-600", bgColor: "bg-gray-50", icon: Clock, label: "Unknown" };
    }
  };

  // Get unique forms for filter
  const uniqueForms = useMemo(() => {
    const formMap = new Map<string, { id: string; title: string }>();
    responses.forEach(response => {
      const key = response.questionId || response.formId || '';
      if (key && !formMap.has(key)) {
        formMap.set(key, {
          id: key,
          title: response.formTitle
        });
      }
    });
    return Array.from(formMap.values());
  }, [responses]);

  // Filter groups
  const filteredGroups = useMemo(() => {
    if (searchQuery === "" && selectedFormIds.length === 0) return groupedResponsesList;

    return groupedResponsesList.filter(group => {
      const matchesSearch = group.modelNo.toLowerCase().includes(searchQuery.toLowerCase());
      const matchesForm = selectedFormIds.length === 0 ||
        group.responses.some(r => selectedFormIds.includes(r.questionId || r.formId || ''));
      return matchesSearch && matchesForm;
    });
  }, [groupedResponsesList, searchQuery, selectedFormIds]);

  // Get all answers from a response for display
  const getAllAnswersList = (response: Response & { formTitle: string }) => {
    const form = formsMap.get(response.questionId || response.formId || '');
    const answers: { questionText: string; answer: any; questionId: string }[] = [];

    if (form?.sections) {
      for (const section of form.sections) {
        for (const question of section.questions || []) {
          const answer = response.answers?.[question.id];
          if (answer && answer !== "" && answer !== null) {
            answers.push({
              questionId: question.id,
              questionText: question.text || question.label || question.id,
              answer: answer
            });
          }
        }
      }
    }

    // Add any unmapped answers
    for (const [key, value] of Object.entries(response.answers || {})) {
      if (value && !answers.some(a => a.questionId === key) && !key.startsWith("__")) {
        answers.push({
          questionId: key,
          questionText: key,
          answer: value
        });
      }
    }

    return answers;
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary-600"></div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-slate-50 to-blue-100/50 dark:from-gray-900 dark:to-gray-800 p-4 sm:p-6 md:p-8">
      {/* Page Header */}
      <div className="mb-6 sm:mb-8 bg-gradient-to-r from-blue-50/80 to-indigo-50/80 dark:from-blue-900/10 dark:to-indigo-900/10 p-4 sm:p-5 rounded-2xl border border-blue-100 dark:border-blue-800/20 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="flex items-center gap-4 min-w-0">
            <div className="p-3 bg-blue-100 dark:bg-blue-900/40 rounded-xl shadow-inner flex-shrink-0">
              <FileText className="w-6 h-6 text-blue-600 dark:text-blue-400" />
            </div>
            <div className="min-w-0">
              <h1 className="text-xl sm:text-2xl font-bold text-gray-900 dark:text-white leading-tight">
                Customer Requests
              </h1>
              <p className="text-gray-500 dark:text-gray-400 text-xs sm:text-sm font-medium mt-0.5">
                Grouped by Model No • View and manage all customer interactions
              </p>
            </div>
          </div>

          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3 lg:flex-1 lg:max-w-2xl lg:justify-end">
            <div className="relative flex-1 sm:max-w-xs">
              <input
                type="text"
                placeholder="Search Model No..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-10 pr-4 py-2.5 bg-white dark:bg-gray-800/50 border border-blue-200 dark:border-blue-700/50 rounded-xl text-gray-900 dark:text-white placeholder-blue-300 focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all shadow-sm text-sm"
              />
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={fetchData}
                disabled={loading}
                className="p-2.5 text-gray-500 hover:text-blue-600 bg-white dark:bg-gray-800/50 border border-blue-100 dark:border-blue-700/50 rounded-xl transition-all hover:shadow-md"
              >
                <RefreshCw className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} />
              </button>

              <button
                onClick={() => setIsAnswerTemplateOpen(true)}
                className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white rounded-xl font-bold transition-all shadow-md text-sm"
              >
                <Upload className="w-4 h-4" />
                <span>Import</span>
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Grouped Responses by Model No */}
      <div className="space-y-6">
        {filteredGroups.map((groupItem) => (
          <div key={groupItem.modelNo} className="bg-gradient-to-br from-white to-blue-50/50 dark:from-gray-700 dark:to-blue-900/10 rounded-2xl border border-blue-100 dark:border-blue-800/30 shadow-sm hover:shadow-lg transition-all duration-300">
            {/* Group Header - Shows Model No once */}
            <div className="flex items-center justify-between p-6 pb-4 border-b-2 border-blue-200 dark:border-blue-800/50">
              <div className="flex items-center gap-3">
                <div className="p-3 bg-blue-100 dark:bg-blue-900/30 rounded-lg">
                  <FileText className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-gray-900 dark:text-white">
                    Model No: {groupItem.modelNo}
                  </h3>
                  <span className="text-sm text-gray-500 dark:text-gray-400">
                    {groupItem.responses.length} response{groupItem.responses.length !== 1 ? 's' : ''} • Latest: {formatTimestamp(groupItem.submissionDate)}
                  </span>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => handleViewDetails(groupItem)}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-sm font-medium transition-all shadow-md hover:shadow-lg"
                >
                  View All ({groupItem.responses.length})
                </button>
                <button
                  onClick={() => handleDeleteGroup(groupItem)}
                  disabled={deletingGroupModelNo === groupItem.modelNo}
                  title="Delete all responses for this model"
                  className="flex items-center gap-1.5 px-3 py-2 bg-red-600 hover:bg-red-700 disabled:opacity-60 disabled:cursor-not-allowed text-white rounded-lg text-sm font-medium transition-all shadow-md hover:shadow-lg"
                >
                  {deletingGroupModelNo === groupItem.modelNo ? (
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  ) : (
                    <Trash2 className="w-4 h-4" />
                  )}
                  <span>{deletingGroupModelNo === groupItem.modelNo ? "Deleting..." : "Delete"}</span>
                </button>
              </div>
            </div>

            {/* Merged Answers - Single Card for the entire group 
            <div className="p-6">
              <div className="bg-gradient-to-br from-blue-50/70 to-indigo-50/50 dark:from-blue-900/15 dark:to-indigo-900/10 rounded-xl border border-blue-100 dark:border-blue-800/30 p-5">
                <div className="mb-3 pb-2 border-b border-blue-200 dark:border-blue-700">
                  <h4 className="text-sm font-semibold text-gray-700 dark:text-gray-300">
                    Merged Answers ({groupItem.responses.length} responses combined)
                  </h4>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                    {groupItem.responses.length === 1 ? "Single response" : "Values from multiple responses are combined with |"}
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                  {getMergedAnswers(groupItem.responses).map(({ questionId, questionText, displayValue, values }) => {
                    // Skip showing Model question separately since it's already in header
                    if (isModelQuestion(questionId, questionText)) return null;

                    const isMerged = values.size > 1;

                    return (
                      <div key={questionId} className="bg-white dark:bg-gray-800/50 rounded-lg p-3 border border-gray-100 dark:border-gray-700">
                        <p className="text-[10px] font-semibold text-blue-600 dark:text-blue-400 uppercase tracking-wider mb-1 truncate" title={questionText}>
                          {questionText}
                          {isMerged && <span className="ml-1 text-orange-500 text-[9px]">(merged)</span>}
                        </p>
                        <p className={`text-sm break-words ${isMerged ? 'text-orange-700 dark:text-orange-300' : 'text-gray-800 dark:text-gray-200'}`}>
                          {displayValue}
                        </p>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>*/}
          </div>
        ))}

        {filteredGroups.length === 0 && (
          <div className="text-center py-16 bg-gradient-to-br from-blue-50 to-indigo-50/50 dark:from-blue-900/10 dark:to-indigo-900/10 rounded-2xl border border-blue-100 dark:border-blue-800/30">
            <div className="w-16 h-16 bg-blue-50 dark:bg-blue-900/20 rounded-full mx-auto mb-4 flex items-center justify-center">
              <FileText className="w-8 h-8 text-blue-600 dark:text-blue-400" />
            </div>
            <h3 className="text-lg font-semibold text-gray-900 dark:text-white mb-2">
              No Customer Requests
            </h3>
            <p className="text-sm text-gray-600 dark:text-gray-400 max-w-md mx-auto">
              {selectedFormIds.length === 0
                ? "There are currently no customer service requests."
                : "No requests match your current filters."}
            </p>
          </div>
        )}
      </div>
      {/* Response Preview Modal - Keeping existing modal code */}
      {selectedResponse && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[9999] overflow-y-auto p-2">
          <div className="bg-gradient-to-br from-white to-blue-50/30 dark:from-gray-800 dark:to-blue-900/10 rounded-2xl shadow-2xl max-w-7xl w-full my-auto max-h-[95vh] flex flex-col">
            {/* Modal content - keep existing */}
            <div className="sticky top-0 bg-gradient-to-r from-blue-50 to-indigo-50 dark:from-blue-900/20 dark:to-indigo-900/20 px-6 py-3 border-b border-blue-200 flex justify-between items-center">
              <h3 className="text-lg font-bold text-gray-900 dark:text-white">{selectedResponse.formTitle}</h3>
              <button onClick={() => setSelectedResponse(null)} className="p-2 hover:bg-gray-200 rounded-lg">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="overflow-y-auto flex-1 p-6">
              <pre className="whitespace-pre-wrap">{JSON.stringify(selectedResponse.answers, null, 2)}</pre>
            </div>
          </div>
        </div>
      )}

      {/* Edit Modal */}
      {editingResponse && editingForm && !editingFormLoading && (
        <ResponseEdit
          response={editingResponse as any}
          question={editingForm as any}
          onSave={handleSaveEditedResponse}
          onCancel={handleCloseEdit}
        />
      )}

      {/* Loading overlay for edit */}
      {editingResponse && editingFormLoading && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white dark:bg-gray-800 rounded-xl px-6 py-4 flex items-center gap-3">
            <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-primary-600"></div>
            <div className="text-primary-600">Loading form details...</div>
          </div>
        </div>
      )}

      <AnswerTemplateImport
        isOpen={isAnswerTemplateOpen}
        onClose={() => setIsAnswerTemplateOpen(false)}
        onSuccess={() => fetchData()}
      />
    </div>
  );
}