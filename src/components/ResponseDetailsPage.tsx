import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { useParams, useNavigate, useLocation } from "react-router-dom";
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
  ArrowLeft,
  TrendingUp, Printer
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
import type { ActiveElement } from "chart.js";
import { apiClient } from "../api/client";
import { formatTimestamp } from "../utils/dateUtils";
import { useNotification } from "../context/NotificationContext";
import { useLogo } from "../context/LogoContext";
import { generateResponseExcelReport } from "../utils/responseExportUtils";
import { ProgressCallback, generateAndDownloadPDF, exportAllResponsesToZip } from "../utils/pdfExportUtils";
import FilePreview from "./FilePreview";
import ResponseEdit from "./ResponseEdit";
import DashboardSummaryCard from "./DashboardSummaryCard";
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

const getRankStyle = (answer: any, darkMode: boolean = false) => {
  if (answer === null || answer === undefined) return "";
  const str = typeof answer === 'object' ? JSON.stringify(answer) : String(answer).trim().toLowerCase();
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  const colors = [
    { l: "bg-blue-50 text-blue-700 border-blue-200", d: "bg-blue-900/30 text-blue-300 border-blue-800" },
    { l: "bg-emerald-50 text-emerald-700 border-emerald-200", d: "bg-emerald-900/30 text-emerald-300 border-emerald-800" },
    { l: "bg-amber-50 text-amber-700 border-amber-200", d: "bg-amber-900/30 text-amber-300 border-amber-800" },
    { l: "bg-orange-50 text-orange-700 border-orange-200", d: "bg-orange-900/30 text-orange-300 border-orange-800" },
    { l: "bg-rose-50 text-rose-700 border-rose-200", d: "bg-rose-900/30 text-rose-300 border-rose-800" },
    { l: "bg-purple-50 text-purple-700 border-purple-200", d: "bg-purple-900/30 text-purple-300 border-purple-800" },
    { l: "bg-pink-50 text-pink-700 border-pink-200", d: "bg-pink-900/30 text-pink-300 border-pink-800" },
    { l: "bg-indigo-50 text-indigo-700 border-indigo-200", d: "bg-indigo-900/30 text-indigo-300 border-indigo-800" },
    { l: "bg-teal-50 text-teal-700 border-teal-200", d: "bg-teal-900/30 text-teal-300 border-teal-800" },
    { l: "bg-cyan-50 text-cyan-700 border-cyan-200", d: "bg-cyan-900/30 text-cyan-300 border-cyan-800" }
  ];
  const color = colors[Math.abs(hash) % colors.length];
  return darkMode ? color.d : color.l;
};

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
}

interface Response {
  _id: string;
  id: string;
  questionId: string;
  formId?: string;
  parentResponseId?: string;
  answers: Record<string, any>;
  createdAt: string;
  updatedAt: string;
  assignedTo?: string;
  status?: string;
  yesNoScore?: {
    yes: number;
    total: number;
  };
  responseRanks?: Record<string, number>;
  submissionMetadata?: SubmissionMetadata;
  dealerName?: string;
}

type SectionStat = {
  id: string;
  title: string;
  yes: number;
  no: number;
  na: number;
  correct: number;
  wrong: number;
  total: number;
  answeredCount: number;
  hasYesNo: boolean;
};

interface OPSTemplateProps {
  form: any;
  response: any;
  submissionHistory?: Array<{ no: number; date: string; issuanceDetails: string }>;
  sameFormatResponses?: any[];  // Add this - all responses with same Format No
  isLoadingSameFormatResponses?: boolean;  // Add this - loading state
  formatQuestionId?: string;  // Add this - ID of the Format No question
  controlQuestionId?: string;  // Add this - ID of the Control No question
  cellTranslations?: Map<string, string>;
  selectedLang?: string;
}

function OPSTemplate({
  form,
  response,
  submissionHistory = [],
  sameFormatResponses = [],
  isLoadingSameFormatResponses = false,

}: OPSTemplateProps) {

  const ASSETS = {
    logo: "/assets/Companylogo.png",
    stop: "/assets/Safetyposter.png",
    noRun: "/assets/Dontrun.png",
    noMob: "/assets/dontusemobile.png",
    ppeG: "/assets/PPEGuide.png",
    ppeGl: "/assets/PPEGUIDE2.png",
    fiveS: "/assets/5S_Guidelines.png",
    qr: "/assets/Qrcode.png",
    shift: "/assets/Shift_timing.png"
  };


  const [translating, setTranslating] = useState(false);

  const INDIAN_LANGUAGES = [
    { code: 'en', label: 'English (off)' },
    { code: 'hi', label: 'Hindi (हिन्दी)' },
    { code: 'bn', label: 'Bengali (বাংলা)' },
    { code: 'te', label: 'Telugu (తెలుగు)' },
    { code: 'mr', label: 'Marathi (मराठी)' },
    { code: 'ta', label: 'Tamil (தமிழ்)' },
    { code: 'gu', label: 'Gujarati (ગુજરાતી)' },
    { code: 'kn', label: 'Kannada (ಕನ್ನಡ)' },
    { code: 'ml', label: 'Malayalam (മലയാളം)' },
    { code: 'pa', label: 'Punjabi (ਪੰਜਾਬੀ)' },
    { code: 'or', label: 'Odia (ଓଡ଼ିଆ)' },
    { code: 'as', label: 'Assamese (অসমীয়া)' },
    { code: 'ur', label: 'Urdu (اردو)' },
    { code: 'sd', label: 'Sindhi (سنڌي)' },
    { code: 'kok', label: 'Konkani (कोंकणी)' },
    { code: 'mai', label: 'Maithili (मैथिली)' },
  ];

  // Columns to translate (matching your PDF generator's TRANSLATE_COL_POSITIONS)
  const TRANSLATE_COL_INDICES = new Set([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
  const [selectedLang, setSelectedLang] = useState<string>('en');
  const [cellTranslations, setCellTranslations] = useState<Map<string, string>>(new Map());
  const handleTranslateOPS = useCallback(async (langCode: string) => {
    if (langCode === 'en') {
      setCellTranslations(new Map());
      setSelectedLang('en');
      return;
    }
    if (!sameFormatResponses.length || !form) return;

    setTranslating(true);
    setSelectedLang(langCode);

    try {
      // DEBUG: Log what we're working with
      console.log('=== TRANSLATION DEBUG ===');
      console.log('sameFormatResponses length:', sameFormatResponses.length);
      console.log('Form sections:', form.sections?.length);

      // Find process steps section
      const allSections = form.sections || [];
      console.log('All sections:', allSections.map((s: any) => ({ id: s.id, title: s.title, questionCount: s.questions?.length })));

      // Find process steps section - look for questions with step pattern
      const processStepsSection = allSections.find((s: any) => {
        return s.questions?.some((q: any) => {
          const qId = (q.id || q._id || "").toLowerCase();
          return qId.includes("step") && (qId.includes("activity") || qId.includes("method"));
        });
      }) || allSections.find((s: any) => {
        const id = (s.id || s._id || "").toLowerCase();
        const title = (s.title || "").toLowerCase();
        return (id.includes("process") || id.includes("step") ||
          title.includes("process") || title.includes("step")) &&
          !id.includes("illust");
      }) || allSections.reduce((best: any, s: any) =>
        (s.questions?.length || 0) > (best?.questions?.length || 0) ? s : best, null);

      if (!processStepsSection) {
        console.error('No process steps section found!');
        return;
      }

      console.log('Process steps section found:', processStepsSection.title, 'Questions:', processStepsSection.questions?.length);

      const processQuestions = (processStepsSection.questions || [])
        .filter((q: any) => !q.showWhen?.questionId);

      console.log('Process questions (filtered):', processQuestions.length);
      console.log('Process questions details:', processQuestions.map((q: any, idx: number) => ({
        index: idx,
        id: q.id,
        label: q.label || q.text,
        type: q.type
      })));

      // Build translation jobs
      const jobs: Array<{ key: string; text: string }> = [];

      sameFormatResponses.forEach((resp, respIdx) => {
        console.log(`Response ${respIdx} answers keys:`, Object.keys(resp?.answers || {}));

        processQuestions.forEach((q: any, colIdx: number) => {
          if (!TRANSLATE_COL_INDICES.has(colIdx)) return;

          const raw = resp?.answers?.[q.id || q._id];
          console.log(`  Question ${colIdx} (${q.id}): raw value =`, raw);

          if (!raw) return;

          let text = '';
          if (typeof raw === 'string') text = raw;
          else if (typeof raw === 'object') {
            if (raw.status) text = raw.status;
            else if (raw.chassisNumber) text = raw.chassisNumber;
            else if (raw.remark) text = raw.remark;
            else { try { text = JSON.stringify(raw); } catch { text = String(raw); } }
          } else text = String(raw);

          console.log(`    Extracted text: "${text}"`);

          if (text.trim()) {
            jobs.push({ key: `${respIdx}:${colIdx}`, text: text.trim() });
            console.log(`    ✅ Added to translation jobs`);
          } else {
            console.log(`    ❌ Skipped - empty text`);
          }
        });
      });

      console.log(`Total translation jobs: ${jobs.length}`);

      if (jobs.length === 0) {
        console.warn('No translation jobs created - nothing to translate');
        setTranslating(false);
        return;
      }

      // Continue with translation...
      const BATCH = 10;
      const newMap = new Map<string, string>();

      for (let i = 0; i < jobs.length; i += BATCH) {
        const batch = jobs.slice(i, i + BATCH);
        await Promise.all(batch.map(async ({ key, text }) => {
          try {
            const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=en|${langCode}&de=appuprasanna460@gmail.com`;
            const res = await fetch(url);
            const data = await res.json();
            if (data?.responseStatus === 200 && data?.responseData?.translatedText) {
              const translated = data.responseData.translatedText;
              if (translated !== text) {
                newMap.set(key, translated);
              }
            }
          } catch (err) {
            console.error(`Translation failed for "${text}":`, err);
          }
        }));
      }

      console.log(`Translated ${newMap.size} cells to ${langCode}`);
      setCellTranslations(newMap);
    } catch (err) {
      console.error('Translation failed:', err);
    } finally {
      setTranslating(false);
    }
  }, [sameFormatResponses, form]);


  // Add these new state variables at the beginning of OPSTemplate component (after the existing hooks)
  const [currentPage, setCurrentPage] = useState(1);
  const [showFullscreenPreview, setShowFullscreenPreview] = useState(false);
  const [currentPreviewPage, setCurrentPreviewPage] = useState(1);
  const responsesPerPage = 5;
  const submissionHistoryPerPage = 5;

  // Paginate responses
  const paginatedResponses = useMemo(() => {
    const start = (currentPage - 1) * responsesPerPage;
    const end = start + responsesPerPage;
    return sameFormatResponses.slice(start, end);
  }, [sameFormatResponses, currentPage]);
  const currentPageResponses = paginatedResponses;

  // Paginate submission history
  const paginatedSubmissionHistory = useMemo(() => {
    const allHistoryEntries: Array<{ no: number; date: string; issuanceDetails: string; responseIndex: number }> = [];

    paginatedResponses.forEach((resp, respIdx) => {   // <-- was sameFormatResponses
      const history = resp.answers?.__submissionHistory || [];
      history.forEach((entry: any, entryIdx: number) => {
        allHistoryEntries.push({
          no: entry.no || entryIdx + 1,
          date: entry.date,
          issuanceDetails: entry.issuanceDetails,
          responseIndex: respIdx,
        });
      });
    });

    allHistoryEntries.sort((a, b) => {
      if (a.date && b.date) return new Date(b.date).getTime() - new Date(a.date).getTime();
      return b.no - a.no;
    });

    // No pagination offset needed — already scoped to current 5 responses
    return allHistoryEntries.slice(0, submissionHistoryPerPage);
  }, [paginatedResponses]);                          // <-- dependency updated

  const totalHistoryEntries = useMemo(() => {
    let count = 0;
    paginatedResponses.forEach((resp) => {    // <-- was sameFormatResponses
      count += (resp.answers?.__submissionHistory || []).length;
    });
    return count;
  }, [paginatedResponses]);
  const totalPages = Math.ceil(sameFormatResponses.length / responsesPerPage);

  // Handle page change
  const handlePageChange = (page: number) => {
    setCurrentPage(page);
    setCurrentPreviewPage(page);
    // Scroll to top when changing page
    setTimeout(() => {
      const opsContainer = document.querySelector('[data-ops-template="true"]');
      if (opsContainer) {
        opsContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }, 100);
  };
  const FullscreenPreview = () => {
    const totalPreviewPages = totalPages;

    const goToPrevPage = () => {
      if (currentPreviewPage > 1) {
        setCurrentPreviewPage(p => p - 1);
        setCurrentPage(p => p - 1);
      }
    };

    const goToNextPage = () => {
      if (currentPreviewPage < totalPreviewPages) {
        setCurrentPreviewPage(p => p + 1);
        setCurrentPage(p => p + 1);
      }
    };

    useEffect(() => {
      const handleKey = (e: KeyboardEvent) => {
        if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === 'PageDown') goToNextPage();
        if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') goToPrevPage();
        if (e.key === 'Escape') setShowFullscreenPreview(false);
      };
      window.addEventListener('keydown', handleKey);
      return () => window.removeEventListener('keydown', handleKey);
    }, [currentPreviewPage]);

    const previewPaginatedResponses = useMemo(() => {
      const start = (currentPreviewPage - 1) * responsesPerPage;
      return sameFormatResponses.slice(start, start + responsesPerPage);
    }, [sameFormatResponses, currentPreviewPage]);

    const previewPaginatedHistory = useMemo(() => {
      const allHistoryEntries: Array<{ no: number; date: string; issuanceDetails: string; responseIndex: number }> = [];
      previewPaginatedResponses.forEach((resp, respIdx) => {
        const history = resp.answers?.__submissionHistory || [];
        history.forEach((entry: any, entryIdx: number) => {
          allHistoryEntries.push({
            no: entry.no || entryIdx + 1,
            date: entry.date,
            issuanceDetails: entry.issuanceDetails,
            responseIndex: respIdx,
          });
        });
      });
      allHistoryEntries.sort((a, b) => {
        if (a.date && b.date) return new Date(b.date).getTime() - new Date(a.date).getTime();
        return b.no - a.no;
      });
      return allHistoryEntries.slice(0, submissionHistoryPerPage);
    }, [previewPaginatedResponses]);

    const pastProblemMergedPreview = useMemo(() => {
      if (!pastProbsQuestions[0]) return { value: "", isMerged: false };
      const values: string[] = [];
      previewPaginatedResponses.forEach(resp => {
        const raw = resp?.answers?.[pastProbsQuestions[0]?.id || pastProbsQuestions[0]?._id];
        if (raw) values.push(String(raw));
      });
      const uniqueValues = Array.from(new Set(values));
      return { value: uniqueValues.join(", "), isMerged: uniqueValues.length > 1 && previewPaginatedResponses.length > 1 };
    }, [previewPaginatedResponses, pastProbsQuestions]);

    const [zoomLevel, setZoomLevel] = useState(1);
    const OPS_NATURAL_WIDTH = 1587;

    // ── Slide area ref for scroll-reset on page change ──
    const slideScrollRef = useRef<HTMLDivElement>(null);
    useEffect(() => {
      if (slideScrollRef.current) slideScrollRef.current.scrollTop = 0;
    }, [currentPreviewPage]);

    // ── Compute zoom to fill the available slide area ──
    const slideAreaRef = useRef<HTMLDivElement>(null);
    useEffect(() => {
      const compute = () => {
        if (slideAreaRef.current) {
          const availW = slideAreaRef.current.clientWidth - 32; // 16px padding each side
          setZoomLevel(Math.min(availW / OPS_NATURAL_WIDTH, 1));
        }
      };
      compute();
      const ro = new ResizeObserver(compute);
      if (slideAreaRef.current) ro.observe(slideAreaRef.current);
      return () => ro.disconnect();
    }, []);
    useEffect(() => {
      const interval = setInterval(() => {
        if (currentPreviewPage < totalPreviewPages) {
          goToNextPage();
        } else {
          setCurrentPreviewPage(1);
          setCurrentPage(1);
        }
      }, 3000);
      return () => clearInterval(interval);
    }, [currentPreviewPage, totalPreviewPages]);
    // ── Shared cell styles (same as main OPSTemplate) ──
    const BORDER = "1px solid #999";
    const BORDER2 = "2px solid #000";
    const C: React.CSSProperties = { border: BORDER, padding: "1px 1.5px", fontSize: "7pt", verticalAlign: "top", wordBreak: "break-word", lineHeight: "1.4" };
    const H: React.CSSProperties = { ...C, background: "#d9d9d9", fontWeight: 700, textAlign: "center", verticalAlign: "middle", fontSize: "6.5pt" };
    const L: React.CSSProperties = { ...C, background: "#e8e8e8", fontWeight: 700, fontSize: "6.5pt", verticalAlign: "middle", lineHeight: "1" };
    const V: React.CSSProperties = { ...C, background: "#fff", verticalAlign: "middle" };
    const T: React.CSSProperties = { width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: "7pt", lineHeight: "1" };
    const hdrCell = { ...C, background: "#d9d9d9", fontWeight: 700, textAlign: "center", verticalAlign: "middle" };
    const cellBase = { border: "1px solid #999", padding: "1px 2px", fontSize: "6pt", lineHeight: 1.2, overflowWrap: "break-word", overflow: "hidden" };
    const lblCell = { ...cellBase, background: "#e8e8e8", fontWeight: 700, verticalAlign: "middle" };
    const boldBorder = { border: "2px solid #000" };
    const live = (v: string) => (v && v !== "—") ? "#15803d" : "#999";

    // ── Computed header values for this preview page ──
    const deptC = getCombinedHeaderAnswer(headerQuestions, 0, previewPaginatedResponses);
    const lineC = getCombinedHeaderAnswer(headerQuestions, 1, previewPaginatedResponses);
    const modelC = getCombinedHeaderAnswer(headerQuestions, 2, previewPaginatedResponses);
    const stationC = getCombinedHeaderAnswer(headerQuestions, 3, previewPaginatedResponses);
    const formatNoC = getCombinedHeaderAnswer(docControlQuestions, 0, previewPaginatedResponses);
    const controlNoC = getCombinedHeaderAnswer(docControlQuestions, 1, previewPaginatedResponses);

    const DEF_FIFO = `1. Bin/trolley must be changed only after complete usage of all material in it.\n2. Empty bin/trolley should be replaced with new one.\n3. Don't top up partially filled bin.\n4. Follow FIFO on line during Process.\n5. Do not use next bin / Trolley material until running not consumed.`;
    const DEF_NONLUB = "Do not use any lubrication if not specified in OPS / Process Sheet.";
    const DEF_ENV = `1. Do waste segregation.\n2. Switch off idle lights & machines.\n3. Ensure 3R Principal in daily activities.\n4. If there was any leakage, communicate to Sub Leader.`;
    const DEF_SAFE = `1. Follow POS sheet in case of any Chemical.\n2. Follow MSDS/SDS in case of any emergency regarding chemical.\n3. Follow your PPE's.`;
    const DEF_PROC_INS = [
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
    const TROUBLE_ROWS = [
      "Equipment Trouble / Machine Break Down",
      "A Trouble You Are Responsible For",
      "Empty Marshal Carrier",
      "Stock Out / Material Shortage",
      "A Trouble From Different Section",
    ];

    const getPreviewIllustrationImages = (resp: any): string[] => {
      const imgQ = illustrationQuestions.find(
        (q: any) => q.type === "file" || q.type === "image" ||
          (q.id || "").toLowerCase().includes("image") ||
          (q.id || "").toLowerCase().includes("illust")
      ) || illustrationQuestions[0];
      if (!imgQ || !resp?.answers) return [];
      const imgVal = resp.answers[imgQ.id || imgQ._id];
      if (!imgVal) return [];
      if (Array.isArray(imgVal)) return imgVal.map((item: any) => typeof item === "string" ? item : item?.url).filter(Boolean);
      if (typeof imgVal === "string") {
        try {
          const parsed = JSON.parse(imgVal);
          if (Array.isArray(parsed)) return parsed.map((item: any) => typeof item === "string" ? item : item?.url).filter(Boolean);
          if (parsed?.url) return [parsed.url];
        } catch { }
        return [imgVal];
      }
      if (typeof imgVal === "object" && imgVal?.url) return [imgVal.url];
      return [];
    };

    // Progress dots — show up to 15, then use counter
    const MAX_DOTS = 15;
    const showDots = totalPreviewPages <= MAX_DOTS;

    return (
      <div
        style={{
          position: "fixed",
          inset: 0,
          zIndex: 10000,
          background: "#1a1a2e",           // deep navy — like PPT presenter bg
          display: "flex",
          flexDirection: "column",
          fontFamily: "Arial, sans-serif",
          userSelect: "none",
        }}
      >
        {/* ══════════════════════════════════════════
          TOP BAR  — mimics PPT presenter toolbar
      ══════════════════════════════════════════ */}
        <div style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          padding: "0 20px",
          height: 52,
          background: "#16213e",
          borderBottom: "1px solid #0f3460",
          flexShrink: 0,
          gap: 12,
        }}>
          {/* Left: title */}
          <div style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 0 }}>
            <div style={{
              width: 32, height: 32, borderRadius: 6,
              background: "#0f3460",
              display: "flex", alignItems: "center", justifyContent: "center",
            }}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#60a5fa" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="2" y="3" width="20" height="14" rx="2" />
                <line x1="8" y1="21" x2="16" y2="21" />
                <line x1="12" y1="17" x2="12" y2="21" />
              </svg>
            </div>
            <div style={{ minWidth: 0 }}>
              <p style={{ color: "#e2e8f0", fontSize: 13, fontWeight: 600, margin: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                OPS Template — {form?.title || "Operation Standard"}
              </p>
              <p style={{ color: "#64748b", fontSize: 11, margin: 0 }}>
                {previewPaginatedResponses.length} response{previewPaginatedResponses.length !== 1 ? 's' : ''} on this slide
              </p>
            </div>
          </div>

          {/* Center: slide counter + nav */}
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button
              onClick={goToPrevPage}
              disabled={currentPreviewPage === 1}
              title="Previous slide (←)"
              style={{
                width: 30, height: 30, borderRadius: 6,
                background: currentPreviewPage === 1 ? "transparent" : "#0f3460",
                border: "1px solid " + (currentPreviewPage === 1 ? "#1e293b" : "#1d4ed8"),
                cursor: currentPreviewPage === 1 ? "not-allowed" : "pointer",
                display: "flex", alignItems: "center", justifyContent: "center",
                opacity: currentPreviewPage === 1 ? 0.3 : 1,
                transition: "all 0.15s",
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#93c5fd" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="15 18 9 12 15 6" />
              </svg>
            </button>

            <div style={{
              background: "#0f3460",
              border: "1px solid #1d4ed8",
              borderRadius: 6,
              padding: "4px 14px",
              display: "flex", alignItems: "center", gap: 6,
            }}>
              <span style={{ color: "#93c5fd", fontSize: 13, fontWeight: 700 }}>{currentPreviewPage}</span>
              <span style={{ color: "#475569", fontSize: 12 }}>/</span>
              <span style={{ color: "#64748b", fontSize: 13 }}>{totalPreviewPages}</span>
            </div>

            <button
              onClick={goToNextPage}
              disabled={currentPreviewPage === totalPreviewPages}
              title="Next slide (→)"
              style={{
                width: 30, height: 30, borderRadius: 6,
                background: currentPreviewPage === totalPreviewPages ? "transparent" : "#0f3460",
                border: "1px solid " + (currentPreviewPage === totalPreviewPages ? "#1e293b" : "#1d4ed8"),
                cursor: currentPreviewPage === totalPreviewPages ? "not-allowed" : "pointer",
                display: "flex", alignItems: "center", justifyContent: "center",
                opacity: currentPreviewPage === totalPreviewPages ? 0.3 : 1,
                transition: "all 0.15s",
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#93c5fd" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="9 18 15 12 9 6" />
              </svg>
            </button>
          </div>

          {/* Right: keyboard hint + close */}
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ color: "#334155", fontSize: 11, whiteSpace: "nowrap" }}>← → navigate · Esc close</span>
            <button
              onClick={() => { setShowFullscreenPreview(false); setCurrentPreviewPage(currentPage); }}
              title="Close (Esc)"
              style={{
                width: 32, height: 32, borderRadius: 6,
                background: "transparent",
                border: "1px solid #334155",
                cursor: "pointer",
                display: "flex", alignItems: "center", justifyContent: "center",
                color: "#94a3b8",
                fontSize: 16,
                transition: "all 0.15s",
              }}
              onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.background = "#7f1d1d"; (e.currentTarget as HTMLButtonElement).style.borderColor = "#ef4444"; }}
              onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.background = "transparent"; (e.currentTarget as HTMLButtonElement).style.borderColor = "#334155"; }}
            >
              ✕
            </button>
          </div>
        </div>

        {/* ══════════════════════════════════════════
          MAIN AREA  — slide canvas
      ══════════════════════════════════════════ */}
        <div
          ref={slideAreaRef}
          style={{
            flex: 1,
            overflow: "hidden",
            display: "flex",
            alignItems: "flex-start",
            justifyContent: "center",
            padding: "20px 16px 8px",
            background: "#1a1a2e",
          }}
        >
          {/* Slide shadow frame — mimics PPT slide card */}
          <div
            style={{
              background: "#fff",
              boxShadow: "0 8px 40px rgba(0,0,0,0.6), 0 2px 8px rgba(0,0,0,0.4)",
              borderRadius: 3,
              overflow: "auto",
              maxHeight: "100%",
              width: OPS_NATURAL_WIDTH * zoomLevel,
              flexShrink: 0,
            }}
            ref={slideScrollRef}
          >
            {/* The actual OPS content at zoom scale */}
            <div
              style={{
                width: OPS_NATURAL_WIDTH,
                transformOrigin: "top left",
                transform: `scale(${zoomLevel})`,
                // When scaled down, the layout space collapses — compensate with negative margin:
                marginBottom: `${-(OPS_NATURAL_WIDTH * (1 - zoomLevel) * 0.8)}px`,
              }}
            >
              {/* ── SLIDE CONTENT (same structure as main OPSTemplate) ── */}
              <div style={{ width: "100%", fontFamily: "Arial,sans-serif", fontSize: "7pt", background: "#fff", color: "#000" }}>

                {/* RETENTION BAR */}
                <table style={{ width: "100%", borderCollapse: "collapse", border: "2px solid #000" }}>
                  <tbody><tr>
                    <td style={{ padding: "1px 6px", textAlign: "right", fontWeight: 700, fontSize: "6.5pt" }}>
                      Retention Period : 20 years after Model is discontinued
                    </td>
                  </tr></tbody>
                </table>

                {/* HEADER TABLE */}
                <table style={{ ...T, border: BORDER2, borderTop: "none" }}>
                  <colgroup>
                    <col style={{ width: "5.5%" }} /><col style={{ width: "3.5%" }} /><col style={{ width: "5%" }} />
                    <col style={{ width: "3.5%" }} /><col style={{ width: "8%" }} /><col style={{ width: "7%" }} />
                    <col style={{ width: "4%" }} /><col style={{ width: "11%" }} /><col style={{ width: "9%" }} />
                    <col style={{ width: "6%" }} /><col style={{ width: "6%" }} /><col style={{ width: "6%" }} />
                    <col style={{ width: "4%" }} /><col style={{ width: "4%" }} /><col style={{ width: "8%" }} />
                    <col style={{ width: "9%" }} />
                  </colgroup>
                  <tbody>
                    <tr style={{ height: 2 }}>
                      <td rowSpan={8} style={{ border: BORDER2, textAlign: "center", verticalAlign: "top", padding: 0, background: "#ffffffff" }}>
                        <img src={ASSETS.logo} alt="Logo" style={{ width: "100%", height: 70, objectFit: "contain" }} />
                      </td>
                      <td style={{ ...L, lineHeight: "1.5", width: "2%" }}>{label(headerQuestions, 0, "Dept. / Section")} :</td>
                      <td style={{ ...V, fontWeight: 700, color: live(deptC), lineHeight: "1.2", width: "10%" }}>{deptC || "—"}</td>
                      <td style={{ ...L, lineHeight: "1.5", width: "2.5%" }}>{label(headerQuestions, 1, "Line / Zones")} :</td>
                      <td style={{ ...V, fontWeight: 700, color: live(lineC), lineHeight: "1.2", width: "12%" }}>{lineC || "—"}</td>
                      <td colSpan={4} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", padding: 2 }}>
                        <div style={{ fontSize: "9pt", fontWeight: 700, letterSpacing: 1 }}>Operation Standard</div>
                      </td>
                      {Array.from({ length: 3 }).map((_, idx) => (
                        <td key={`empty-${idx}`} rowSpan={7} style={{ border: BORDER2, verticalAlign: "top", padding: 0 }} />
                      ))}
                      {Array.from({ length: 3 }).map((_, idx) => {
                        const colIdx = idx + 4;
                        return (
                          <td key={idx} rowSpan={7} style={{ border: BORDER2, verticalAlign: "top", padding: 0, background: "#fff" }}>
                            <table style={{ width: "100%", height: "100%", borderCollapse: "collapse" }}>
                              {Array.from({ length: 5 }).map((_, i) => {
                                const entry = i < previewPaginatedHistory.length ? previewPaginatedHistory[i] : null;
                                let display = "\u00A0";
                                if (entry) {
                                  if (colIdx === 4) display = String(entry.no);
                                  else if (colIdx === 5) {
                                    try { display = new Date(entry.date).toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "2-digit" }); }
                                    catch { display = entry.date || "\u00A0"; }
                                  } else { display = entry.issuanceDetails || "\u00A0"; }
                                }
                                return (
                                  <tr key={i}>
                                    <td style={{ borderBottom: "1px solid #ccc", height: 36, padding: "0 2px", fontSize: "5.5pt", textAlign: "center", color: entry ? "#15803d" : "transparent", fontWeight: entry ? 700 : 400, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                                      {display}
                                    </td>
                                  </tr>
                                );
                              })}
                            </table>
                          </td>
                        );
                      })}


                      <td rowSpan={8} style={{ border: BORDER2, verticalAlign: "top", padding: "2px 3px", fontSize: "6.5pt" }}>
                        <div style={{ fontWeight: 700, color: "#c00", lineHeight: "1", marginBottom: 2 }}>{label(docControlQuestions, 4, "Format No. AP")} :</div>
                        <div style={{ fontWeight: 700, fontSize: "7pt", lineHeight: "1.2", marginBottom: 1, color: live(formatNoC) }}>{formatNoC || "—"}</div>
                        <div style={{ borderTop: "0.5px solid #999", margin: "2px 0" }} />
                        <div style={{ fontWeight: 700, color: "#c00", lineHeight: "1", marginBottom: 1 }}>{label(docControlQuestions, 5, "Control No. AP")} :</div>
                        <div style={{ fontWeight: 700, fontSize: "7pt", marginBottom: 1, lineHeight: "1.2", color: live(controlNoC) }}>{controlNoC || "—"}</div>
                        <div style={{ borderTop: "0.5px solid #999", margin: "2px 0" }} />
                        <div style={{ fontWeight: 700, marginBottom: 1 }}>QR Code :</div>
                        <img src={ASSETS.qr} alt="QR" style={{ width: 40, height: 36, objectFit: "contain" }} />
                      </td>
                    </tr>
                    <tr style={{ height: 4 }}>
                      <td style={{ ...L, lineHeight: "1.5" }}>{label(headerQuestions, 2, "Model AP")}</td>
                      <td style={{ ...V, fontWeight: 700, color: live(modelC) }}>{modelC || "—"}</td>
                      <td style={{ ...L, lineHeight: "1.2" }}>{label(headerQuestions, 3, "Process / Station AP")} :</td>
                      <td style={{ ...V, fontWeight: 700, color: live(stationC) }}>{stationC || "—"}</td>
                      <td colSpan={4} style={{ ...H, border: BORDER2, fontSize: "5.5pt" }}>Your Work When Trouble Stopped The Production Line</td>
                    </tr>
                    <tr style={{ height: 2 }}>
                      <td rowSpan={6} colSpan={2} style={{ border: BORDER2, verticalAlign: "top", fontSize: "5pt", padding: "2px 3px" }}>
                        <div style={{ fontWeight: 900, marginBottom: 5, fontSize: 8 }}>REJECTION HANDLING :-</div>
                        <div style={{ marginBottom: 2, fontSize: 7 }}>Clearly Identify Rejected / NG parts.</div>
                        <div style={{ lineHeight: "1.5", fontSize: 7 }}>Keep them properly with proper identification at defined Location.</div>
                      </td>
                      <td rowSpan={6} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", fontWeight: 700, fontSize: "6pt", padding: 1, lineHeight: 1.5 }}>Measuring<br />Instruments<br />or Gauges</td>
                      <td rowSpan={6} style={{ border: BORDER2, verticalAlign: "top", fontSize: "5pt", padding: 0 }}>
                        {["Always use Calibrated Measuring Instruments / Gauges.", "Ensure Zero setting before use.", "Do Not Use Unidentified Measuring Tool / Gauges.", "In case of any abnormality, inform Line leader and Quality Engineer."].map((txt, i, arr) => (
                          <div key={i} style={{ padding: "2px 1px", borderBottom: i < arr.length - 1 ? "0.5px solid #ccc" : "none", lineHeight: "1.5" }}>{txt}</div>
                        ))}
                      </td>
                      <td rowSpan={6} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", padding: 2 }}>
                        <img src={ASSETS.stop} alt="Stop Call Wait" style={{ maxWidth: "100%", height: 90, objectFit: "contain" }} />
                      </td>
                      <td style={{ ...H }}>S. No.</td>
                      <td style={H}>Trouble</td>
                      <td style={H}>Your task</td>
                    </tr>
                    <tr style={{ height: 4 }}>
                      <td style={{ ...C, textAlign: "center" }}>1</td>
                      <td style={{ ...C, fontSize: "5.5pt" }}>{TROUBLE_ROWS[0]}</td>
                      <td rowSpan={5} style={{ fontSize: "5pt", textAlign: "center", verticalAlign: "middle", lineHeight: "1.4" }}>Stop The Line<br />Inform the Zone Leader<br />Write on card if mentioned in OPS</td>
                    </tr>
                    {TROUBLE_ROWS.slice(1).map((row, i) => (
                      <tr key={i} style={{ height: 7 }}>
                        <td style={{ ...C, textAlign: "center" }}>{i + 2}</td>
                        <td style={{ ...C, fontSize: "5.5pt" }}>{row}</td>
                        {i === 3 && (
                          <>
                            <td style={{ ...H, border: BORDER2, fontSize: "6pt" }}>Prepared</td>
                            <td style={{ ...H, fontSize: "6pt" }}>Checked</td>
                            <td style={{ ...H, fontSize: "6pt" }}>Approved</td>
                            <td style={{ ...H, fontSize: "6pt" }}>No.</td>
                            <td style={{ ...H, fontSize: "6pt" }}>DD/MM/YY</td>
                            <td style={{ ...H, border: BORDER2, fontSize: "6pt" }}>Issuance / Revision details</td>
                          </>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>

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
                    <col style={{ width: "8%" }} />
                    <col style={{ width: "17%" }} />
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
                      <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
                        <div style={{ fontWeight: 800, marginBottom: 0, fontSize: "5.5pt" }}>FIFO System</div>
                        {DEF_FIFO.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, lineHeight: "1.4", fontSize: "5.5pt" }}>{l}</div>)}
                      </td>
                      <td style={{ ...C, verticalAlign: "top", padding: 0 }}>
                        <div style={{ padding: "1px 2px", borderBottom: "0.5px solid #ccc", fontSize: "5pt", lineHeight: "1.4" }}>{DEF_NONLUB}</div>
                        <div style={{ display: "flex", borderBottom: "0.5px solid #ccc" }}>
                          <div style={{ flex: 1, borderRight: "0.5px solid #ccc", padding: "1px 2px", textAlign: "center", fontWeight: 700, fontSize: "4.5pt", background: "#d9d9d9", lineHeight: "1.4" }}>No mobile on shopfloor</div>
                          <div style={{ flex: 1, padding: "1px 2px", textAlign: "center", fontWeight: 700, fontSize: "4.5pt", background: "#d9d9d9", lineHeight: "1.4" }}>Do not run on shopfloor</div>
                        </div>
                        <div style={{ display: "flex" }}>
                          <div style={{ flex: 1, borderRight: "0.5px solid #ccc", padding: 1, textAlign: "center" }}>
                            <img src={ASSETS.noMob} alt="No Mobile" style={{ width: 44, height: 55, objectFit: "contain" }} />
                          </div>
                          <div style={{ flex: 1, padding: 1, textAlign: "center" }}>
                            <img src={ASSETS.noRun} alt="No Run" style={{ width: 44, height: 55, objectFit: "contain" }} />
                          </div>
                        </div>
                      </td>
                      <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 1 }}>
                        <img src={ASSETS.ppeG} alt="Full PPE Uniform" style={{ width: "100%", height: 95, objectFit: "contain" }} />
                      </td>
                      <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 1 }}>
                        <img src={ASSETS.ppeGl} alt="Station PPE" style={{ width: "100%", maxHeight: 95, objectFit: "contain" }} />
                      </td>
                      <td style={{ ...C, textAlign: "center", verticalAlign: "top", padding: 1 }}>
                        <img src={ASSETS.shift} alt="Shift Timings" style={{ width: "100%", height: 90, objectFit: "contain" }} />
                      </td>
                      <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
                        <div style={{ fontWeight: 700, color: "#166534", marginBottom: 1, fontSize: "5pt" }}>Environmental Issues</div>
                        {DEF_ENV.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, fontSize: "5pt", lineHeight: "1.4" }}>{l}</div>)}
                      </td>
                      <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
                        <div style={{ fontWeight: 700, color: "#991b1b", marginBottom: 1, fontSize: "5pt" }}>Safety Issues</div>
                        {DEF_SAFE.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, lineHeight: "1.4", fontSize: "5pt" }}>{l}</div>)}
                      </td>
                      <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 1 }}>
                        <img src={ASSETS.fiveS} alt="5S Guidelines" style={{ width: "100%", height: 95, objectFit: "fill" }} />
                      </td>
                      <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
                        {DEF_PROC_INS.map((l, i) => (
                          <div key={i} style={{ marginBottom: 0, lineHeight: "1.4", fontSize: "5pt" }}>
                            {l}
                          </div>
                        ))}
                      </td>
                    </tr>
                  </tbody>
                </table>
                {/* PROCESS STEPS TABLE */}
                <table style={{ width: "100%", borderCollapse: "collapse", borderLeft: BORDER2, borderRight: BORDER2, borderBottom: BORDER2, tableLayout: "fixed" }}>
                  <colgroup>
                    <col style={{ width: "8%" }} />
                    <col style={{ width: "2.5%" }} />
                    {processStepColumns.map((col, i) => <col key={i} style={{ width: col.width }} />)}
                  </colgroup>
                  <thead>
                    <tr>
                      <th style={{ ...hdrCell, background: "#ffff00", color: "#000" }}>
                        {illustrationQuestions[0]?.text || illustrationQuestions[0]?.label || "Illustrations"}
                      </th>
                      <th style={hdrCell}>SN</th>
                      {processStepColumns.map((col) => (
                        <th key={col.questionId} style={{ ...hdrCell, whiteSpace: "pre-line", lineHeight: 1.2 }}>{col.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewPaginatedResponses?.map((resp, respIdx) => {
                      const globalIdx = (currentPreviewPage - 1) * responsesPerPage + respIdx;
                      const illustrationImages = getPreviewIllustrationImages(resp);
                      return (
                        <tr key={`resp-${globalIdx}`} style={{ minHeight: 50 }}>
                          <td style={{ ...cellBase, background: "#ffffffff", textAlign: "center", verticalAlign: "middle", padding: 2, height: 50 }}>
                            {illustrationImages.length > 0 ? (
                              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, justifyContent: 'center' }}>
                                {illustrationImages.map((url, idx) => (
                                  <img key={idx} src={url} alt={`Illus ${idx + 1}`}
                                    style={{ width: 50, height: 40, objectFit: 'contain', cursor: 'pointer', border: '0.5px solid #ccc', borderRadius: 2 }}
                                    onClick={() => window.open(url, '_blank')} />
                                ))}
                              </div>
                            ) : (
                              <div style={{ width: "100%", height: 40, display: "flex", alignItems: "center", justifyContent: "center", color: "#bbb", fontSize: 20 }}>📷</div>
                            )}
                          </td>
                          <td style={{ ...cellBase, textAlign: "center", fontWeight: 700, fontSize: "9pt", verticalAlign: "middle" }}>{globalIdx + 1}</td>
                          {processStepColumns.map((col) => {
                            const rawVal = resp?.answers?.[col.questionId];
                            let cellVal = "";
                            if (rawVal !== null && rawVal !== undefined && rawVal !== "") {
                              if (typeof rawVal === "object") { try { cellVal = JSON.stringify(rawVal); } catch { cellVal = String(rawVal); } }
                              else { cellVal = String(rawVal); }
                            }
                            return (
                              <td key={col.questionId} style={{ ...cellBase, fontWeight: 500, color: cellVal ? '#000' : 'transparent', height: 50, minHeight: 50, padding: 3, verticalAlign: "top", background: "#fff", lineHeight: 1.3, textAlign: "left" }}>
                                {cellVal || '\u00A0'}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                    {previewPaginatedResponses && previewPaginatedResponses.length < responsesPerPage &&
                      Array.from({ length: responsesPerPage - previewPaginatedResponses.length }).map((_, i) => (
                        <tr key={`blank-${i}`} style={{ minHeight: 50 }}>
                          <td style={{ ...cellBase, background: "#ffffffff", height: 50 }} />
                          <td style={{ ...cellBase, textAlign: 'center', fontWeight: 700, fontSize: '9pt', verticalAlign: 'middle', color: '#ccc' }}>
                            {(currentPreviewPage - 1) * responsesPerPage + previewPaginatedResponses.length + i + 1}
                          </td>
                          {processStepColumns.map(col => (
                            <td key={col.questionId} style={{ ...cellBase, height: 50, background: "#fff" }}>&nbsp;</td>
                          ))}
                        </tr>
                      ))
                    }
                  </tbody>
                </table>

                {/* ABNORMALITY + PAST PROBLEMS */}
                <table style={{ width: "100%", borderCollapse: "collapse", borderLeft: BORDER2, borderRight: BORDER2, borderBottom: BORDER2, tableLayout: "fixed" }}>
                  <tbody>
                    <tr>
                      <td rowSpan={2} style={{ ...cellBase, padding: 4, verticalAlign: "top", fontSize: "6.5pt", width: "12%" }}>
                        <div style={{ fontWeight: 700, marginBottom: 2 }}>{label(pastProbsQuestions, 0, "Abnormality handling route")} :</div>
                        <div>In case of any abnormality inform the Zone In-Charge</div>
                        <div style={{ marginTop: 2 }}>Flow of Communication :-</div>
                        <div>Operator ▶ Team Member ▶ Section Mgr ▶ As required</div>
                      </td>
                      <td style={{ ...lblCell, padding: "2px", textAlign: "center", fontSize: "7pt" }}>{label(pastProbsQuestions, 1, "Past Problem Details")}</td>
                    </tr>
                    <tr>
                      <td style={{ ...cellBase, padding: 4, verticalAlign: "top", minHeight: 60, height: 60, fontSize: "7pt", background: "#fff", fontStyle: pastProblemMergedPreview.isMerged ? "italic" : "normal", color: pastProblemMergedPreview.value ? (pastProblemMergedPreview.isMerged ? "#b45309" : "#000") : "#bbb" }}>
                        {pastProblemMergedPreview.value || <span style={{ color: "#ccc", fontStyle: "italic" }}>No data recorded</span>}
                      </td>
                    </tr>
                  </tbody>
                </table>
                {/* PAGE NUMBER */}
                <table style={{ width: "100%", borderCollapse: "collapse", border: BORDER2 }}>
                  <tbody><tr>
                    <td style={{ ...cellBase, padding: 2 }}>{form?.title || "Operation Standard"}</td>
                    <td style={{ ...hdrCell as any, padding: 3, fontSize: "8.5pt" }}>
                      Page Number : {currentPreviewPage} / {totalPreviewPages}
                    </td>
                  </tr></tbody>
                </table>

              </div>
            </div>
          </div>
        </div>

        {/* ══════════════════════════════════════════
          BOTTOM BAR  — progress dots + info strip
      ══════════════════════════════════════════ */}
        {/* ══════════════════════════════════════════
  BOTTOM BAR — dynamic progress indicator
══════════════════════════════════════════ */}
        <div style={{
          height: 44,
          background: "#16213e",
          borderTop: "1px solid #0f3460",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          gap: 6,
          flexShrink: 0,
          padding: "0 16px",
        }}>
          {totalPreviewPages <= 20 ? (
            // Show dots when 20 pages or less
            Array.from({ length: totalPreviewPages }, (_, i) => i + 1).map(pageNum => {
              const isActive = pageNum === currentPreviewPage;
              return (
                <button
                  key={pageNum}
                  onClick={() => { setCurrentPreviewPage(pageNum); setCurrentPage(pageNum); }}
                  title={`Slide ${pageNum}`}
                  style={{
                    width: isActive ? 28 : 8,
                    height: 8,
                    borderRadius: 4,
                    background: isActive ? "#3b82f6" : "#1e3a5f",
                    border: isActive ? "none" : "1px solid #1d4ed8",
                    cursor: "pointer",
                    padding: 0,
                    transition: "all 0.2s",
                    flexShrink: 0,
                  }}
                />
              );
            })
          ) : (
            // For more than 20 pages: show mini progress bar + page selector dropdown
            <div style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", maxWidth: 500 }}>
              <span style={{ color: "#64748b", fontSize: 11, whiteSpace: "nowrap" }}>
                Page
              </span>
              <select
                value={currentPreviewPage}
                onChange={(e) => {
                  const newPage = parseInt(e.target.value);
                  setCurrentPreviewPage(newPage);
                  setCurrentPage(newPage);
                }}
                style={{
                  background: "#0f3460",
                  border: "1px solid #1d4ed8",
                  borderRadius: 6,
                  padding: "4px 8px",
                  color: "#93c5fd",
                  fontSize: 12,
                  fontWeight: 600,
                  cursor: "pointer",
                  outline: "none",
                }}
              >
                {Array.from({ length: totalPreviewPages }, (_, i) => i + 1).map(pageNum => (
                  <option key={pageNum} value={pageNum} style={{ background: "#16213e", color: "#fff" }}>
                    {pageNum}
                  </option>
                ))}
              </select>
              <span style={{ color: "#64748b", fontSize: 11, whiteSpace: "nowrap" }}>
                of {totalPreviewPages}
              </span>

              <div style={{ flex: 1, height: 4, background: "#0f3460", borderRadius: 2, overflow: "hidden" }}>
                <div style={{
                  height: "100%",
                  width: `${(currentPreviewPage / totalPreviewPages) * 100}%`,
                  background: "#3b82f6",
                  borderRadius: 2,
                  transition: "width 0.3s ease",
                }} />
              </div>

              <span style={{ color: "#334155", fontSize: 11, whiteSpace: "nowrap", minWidth: 45 }}>
                {Math.round((currentPreviewPage / totalPreviewPages) * 100)}%
              </span>
            </div>
          )}
        </div>
      </div>
    );
  };
  // ============================================================
  // MERGED VALUE HELPER - Same answer = show once, Different = italic+small
  // ============================================================

  const getMergedCellValue = useCallback((questionId: string): {
    value: string;
    isMerged: boolean;
    count: number;
  } => {
    if (!sameFormatResponses || sameFormatResponses.length === 0) {
      return { value: "", isMerged: false, count: 0 };
    }

    const values: string[] = [];

    sameFormatResponses.forEach(resp => {
      const raw = resp?.answers?.[questionId];
      if (raw === null || raw === undefined || raw === "") return;
      let strVal = "";
      if (typeof raw === "object") {
        if (raw.status) strVal = String(raw.status);
        else if (raw.chassisNumber) strVal = String(raw.chassisNumber);
        else { try { strVal = JSON.stringify(raw); } catch { strVal = String(raw); } }
      } else {
        strVal = String(raw);
      }
      if (strVal.trim()) values.push(strVal.trim());
    });

    if (values.length === 0) return { value: "", isMerged: false, count: 0 };

    const uniqueValues = Array.from(new Set(values));

    if (uniqueValues.length === 1) {
      // All responses have the same answer — show once, normal style
      return { value: uniqueValues[0], isMerged: false, count: values.length };
    }

    // Different answers across responses — combine with comma
    return { value: uniqueValues.join(", "), isMerged: true, count: values.length };
  }, [sameFormatResponses]);


  // ============================================================
  // DYNAMIC SECTION & QUESTION RESOLUTION
  // ============================================================

  const allSections = useMemo(() => form?.sections || [], [form]);

  const getQuestionsFromSectionByIndex = useCallback((sectionIndex: number) => {
    const section = allSections[sectionIndex];
    if (!section) return [];
    return (section.questions || []).filter((q: any) => !q.showWhen?.questionId);
  }, [allSections]);

  const getQuestionsFromSectionById = useCallback((sectionId: string, fallbackIndex?: number) => {
    const byId = allSections.find((s: any) => s.id === sectionId || s._id === sectionId);
    if (byId) return (byId.questions || []).filter((q: any) => !q.showWhen?.questionId);
    if (fallbackIndex !== undefined) return getQuestionsFromSectionByIndex(fallbackIndex);
    return [];
  }, [allSections, getQuestionsFromSectionByIndex]);

  // Find process steps section - the one with most questions (should have 14 columns)
  const processStepsSection = useMemo(() => {
    // Look for section with "process" or "step" in name first
    const byName = allSections.find((s: any) => {
      const id = (s.id || s._id || "").toLowerCase();
      const title = (s.title || "").toLowerCase();
      return id.includes("process") || id.includes("step") ||
        title.includes("process") || title.includes("step");
    });
    if (byName && byName.questions?.length >= 10) return byName;

    // Fallback: section with most questions (should be the process steps section)
    return allSections.reduce((best: any, s: any) => {
      return (s.questions?.length || 0) > (best?.questions?.length || 0) ? s : best;
    }, null);
  }, [allSections]);

  // Find illustrations section - separate section for images
  const illustrationsSection = useMemo(() => {
    return allSections.find((s: any) => {
      const id = (s.id || s._id || "").toLowerCase();
      const title = (s.title || "").toLowerCase();
      return id.includes("illust") || id.includes("image") || id.includes("photo") ||
        title.includes("illust") || title.includes("image") || title.includes("photo");
    }) || null;
  }, [allSections]);

  // Find past problems section
  const pastProblemsSection = useMemo(() => {
    return allSections.find((s: any) => {
      const id = (s.id || s._id || "").toLowerCase();
      const title = (s.title || "").toLowerCase();
      return id.includes("past") || id.includes("problem") || id.includes("abnormal") ||
        title.includes("past") || title.includes("problem") || title.includes("abnormal");
    }) || null;
  }, [allSections]);

  // Find doc control section (Format No, Control No)
  const docControlSection = useMemo(() => {
    return allSections.find((s: any) => {
      const id = (s.id || s._id || "").toLowerCase();
      const title = (s.title || "").toLowerCase();
      return id.includes("doc") || id.includes("control") || id.includes("format") ||
        title.includes("doc") || title.includes("control") || title.includes("format");
    }) || null;
  }, [allSections]);

  // Header section = first section that isn't special
  const headerSection = useMemo(() => {
    const specialSections = new Set([
      processStepsSection?.id,
      illustrationsSection?.id,
      pastProblemsSection?.id,
      docControlSection?.id,
    ].filter(Boolean));

    return allSections.find((s: any) => !specialSections.has(s.id || s._id)) || allSections[0] || null;
  }, [allSections, processStepsSection, illustrationsSection, pastProblemsSection, docControlSection]);



  // Final question arrays
  const headerQuestions = useMemo(() =>
    (headerSection?.questions || []).filter((q: any) => !q.showWhen?.questionId),
    [headerSection]);

  const docControlQuestions = useMemo(() =>
    (docControlSection?.questions || []).filter((q: any) => !q.showWhen?.questionId),
    [docControlSection]);

  // PROCESS QUESTIONS - THESE ARE THE 14 COLUMNS (Importance, Activity, Method, etc.)
  const processQuestions = useMemo(() =>
    (processStepsSection?.questions || []).filter((q: any) => !q.showWhen?.questionId),
    [processStepsSection]);

  const pastProbsQuestions = useMemo(() =>
    (pastProblemsSection?.questions || []).filter((q: any) => !q.showWhen?.questionId),
    [pastProblemsSection]);

  // ILLUSTRATION QUESTIONS - ONLY FOR IMAGES (separate from process steps)
  const illustrationQuestions = useMemo(() =>
    (illustrationsSection?.questions || []).filter((q: any) => !q.showWhen?.questionId),
    [illustrationsSection]);
  // ============================================================
  // ALL ILLUSTRATION IMAGES - collect thumbnails from every response
  // ============================================================

  const allIllustrationImages = useMemo((): Array<{ url: string; respIdx: number }> => {
    const result: Array<{ url: string; respIdx: number }> = [];

    const parseImages = (imgVal: any): string[] => {
      if (!imgVal) return [];
      if (Array.isArray(imgVal)) {
        return imgVal.map((item: any) => typeof item === "string" ? item : item?.url).filter(Boolean);
      }
      if (typeof imgVal === "string") {
        try {
          const parsed = JSON.parse(imgVal);
          if (Array.isArray(parsed)) return parsed.map((item: any) => typeof item === "string" ? item : item?.url).filter(Boolean);
          if (parsed?.url) return [parsed.url];
        } catch { /* not JSON */ }
        return [imgVal];
      }
      if (typeof imgVal === "object" && imgVal?.url) return [imgVal.url];
      return [];
    };

    // Find illustration question ID
    const imgQ = illustrationQuestions.find(
      (q: any) => q.type === "file" || q.type === "image" ||
        (q.id || "").toLowerCase().includes("image") ||
        (q.id || "").toLowerCase().includes("illust")
    ) || illustrationQuestions[0];

    if (!imgQ) return result;
    const qId = imgQ.id || imgQ._id;

    sameFormatResponses.forEach((resp, idx) => {
      const urls = parseImages(resp?.answers?.[qId]);
      urls.forEach(url => result.push({ url, respIdx: idx }));
    });

    return result;
  }, [sameFormatResponses, illustrationQuestions]);
  // ============================================================
  // GENERIC ANSWER HELPERS
  // ============================================================

  const getAnswerByQuestionId = useCallback((questionId: string, resp: any = response): string => {
    if (!resp?.answers || !questionId) return "";
    const answer = resp.answers[questionId];
    if (answer === null || answer === undefined || answer === "") return "";
    if (typeof answer === "object") {
      if (answer.status) return String(answer.status);
      if (answer.chassisNumber) return String(answer.chassisNumber);
      try { return JSON.stringify(answer); } catch { return String(answer); }
    }
    return String(answer);
  }, [response]);

  const getAnswerByIndex = useCallback((questions: any[], index: number, resp: any = response): string => {
    const q = questions[index];
    if (!q) return "";
    return getAnswerByQuestionId(q.id || q._id, resp);
  }, [getAnswerByQuestionId, response]);

  const getCombinedHeaderAnswer = useCallback((questions: any[], index: number, responsesOverride?: any[]): string => {
    const responses = (responsesOverride ?? (sameFormatResponses.length > 0 ? sameFormatResponses : [response]));
    const uniqueValues = new Set<string>();
    responses.forEach(resp => {
      const answer = getAnswerByIndex(questions, index, resp);
      if (answer) uniqueValues.add(answer);
    });
    return Array.from(uniqueValues).join(", ") || "—";
  }, [sameFormatResponses, response, getAnswerByIndex]);

  const label = useCallback((questions: any[], index: number, fallback: string): string => {
    const q = questions[index];
    return q?.text || q?.label || fallback;
  }, []);

  // ============================================================
  // PROCESS STEP COLUMNS - FROM processQuestions (14 columns)
  // ============================================================
  const processStepColumns = useMemo(() => {
    if (processQuestions.length === 0) {
      // Fallback hardcoded columns if no process questions found
      return [
        { field: 'importance', defaultLabel: 'Item Importance', width: '5%' },
        { field: 'activity', defaultLabel: 'What / Activity', width: '8%' },
        { field: 'method', defaultLabel: 'Method (How)', width: '12%' },
        { field: 'frequency', defaultLabel: 'Frequency / When', width: '5%' },
        { field: 'standard', defaultLabel: 'Standard (Spec./Criteria)', width: '12%' },
        { field: 'responsibility', defaultLabel: 'Responsibility', width: '5%' },
        { field: 'equipment', defaultLabel: 'Equipment / Measuring Eq.', width: '5%' },
        { field: 'abnormalities', defaultLabel: 'Possible Abnormalities', width: '6%' },
        { field: 'reactionPlan', defaultLabel: 'Reaction Plan', width: '6%' },
        { field: 'partNameQty', defaultLabel: 'Part Name & QTY', width: '6%' },
        { field: 'ppe', defaultLabel: 'PPEs required', width: '5%' },
        { field: 'document', defaultLabel: 'Required/Document', width: '6%' },
        { field: 'remarks', defaultLabel: 'Remarks', width: '6%' },
      ];
    }

    // Map each question in process section to a column
    return processQuestions.map((q: any, idx: number) => ({
      questionId: q.id || q._id,
      label: q.text || q.label || q.id,
      width: q.columnWidth || `${Math.max(6, Math.floor(80 / Math.max(processQuestions.length, 1)))}%`,
      originalIndex: idx, // Store original index for translation
    }));
  }, [processQuestions]);

  // ============================================================
  // ILLUSTRATION IMAGES (separate from process steps)
  // ============================================================

  const getResponseIllustrationImages = (resp: any): string[] => {
    // Find image/file question in the illustrations section
    const imgQ = illustrationQuestions.find(
      (q: any) => q.type === "file" || q.type === "image" ||
        (q.id || "").toLowerCase().includes("image") ||
        (q.id || "").toLowerCase().includes("illust")
    ) || illustrationQuestions[0];

    if (!imgQ || !resp?.answers) return [];
    const imgVal = resp.answers[imgQ.id || imgQ._id];
    if (!imgVal) return [];

    if (Array.isArray(imgVal)) {
      return imgVal.map((item: any) => typeof item === "string" ? item : item?.url).filter(Boolean);
    }
    if (typeof imgVal === "string") {
      try {
        const parsed = JSON.parse(imgVal);
        if (Array.isArray(parsed)) return parsed.map((item: any) => typeof item === "string" ? item : item?.url).filter(Boolean);
        if (parsed?.url) return [parsed.url];
      } catch { /* not JSON */ }
      return [imgVal];
    }
    if (typeof imgVal === "object" && imgVal?.url) return [imgVal.url];
    return [];
  };


  const isImportanceColumn = (col: { label: string }) => {
    const l = col.label.toLowerCase();
    // Still identify importance column for bold font styling
    return l.includes("importance") || l.includes("priority") || l.includes("item importance");
  };

  // ============================================================
  // HEADER VALUES
  // ============================================================

  const deptCombined = getCombinedHeaderAnswer(headerQuestions, 0, paginatedResponses);
  const lineCombined = getCombinedHeaderAnswer(headerQuestions, 1, paginatedResponses);
  const modelCombined = getCombinedHeaderAnswer(headerQuestions, 2, paginatedResponses);
  const stationCombined = getCombinedHeaderAnswer(headerQuestions, 3, paginatedResponses);
  const formatNoCombined = getCombinedHeaderAnswer(docControlQuestions, 0, paginatedResponses);
  const controlNoCombined = getCombinedHeaderAnswer(docControlQuestions, 1, paginatedResponses);
  const pastProblemMerged = useMemo(() => {
    if (!pastProbsQuestions[0]) return { value: "", isMerged: false };
    const qId = pastProbsQuestions[0]?.id || pastProbsQuestions[0]?._id;
    const values: string[] = [];
    paginatedResponses.forEach(resp => {          // <-- was sameFormatResponses
      const raw = resp?.answers?.[qId];
      if (raw) values.push(String(raw));
    });
    const uniqueValues = Array.from(new Set(values));
    return {
      value: uniqueValues.join(", "),
      isMerged: uniqueValues.length > 1 && paginatedResponses.length > 1
    };
  }, [pastProbsQuestions, paginatedResponses]);   // <-- dependency updated
  // ============================================================
  // STATIC CONTENT (General Instructions - can also be made dynamic)
  // ============================================================
  const DEF_FIFO = `1. Bin/trolley must be changed only after complete usage of all material in it.\n2. Empty bin/trolley should be replaced with new one.\n3. Don't top up partially filled bin.\n4. Follow FIFO on line during Process.\n5. Do not use next bin / Trolley material until running not consumed.`;
  const DEF_NONLUB = "Do not use any lubrication if not specified in OPS / Process Sheet.";
  const DEF_ENV = `1. Do waste segregation.\n2. Switch off idle lights & machines.\n3. Ensure 3R Principal in daily activities.\n4. If there was any leakage, communicate to Sub Leader.`;
  const DEF_SAFE = `1. Follow POS sheet in case of any Chemical.\n2. Follow MSDS/SDS in case of any emergency regarding chemical.\n3. Follow your PPE's.`;
  const DEF_PROC_INS = [
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
  const TROUBLE_ROWS = [
    "Equipment Trouble / Machine Break Down",
    "A Trouble You Are Responsible For",
    "Empty Marshal Carrier",
    "Stock Out / Material Shortage",
    "A Trouble From Different Section",
  ];

  const cellBase = {
    border: "1px solid #999", padding: "1px 2px", fontSize: "6pt",
    lineHeight: 1.2, overflowWrap: "break-word", overflow: "hidden",
  };
  const hdrCell = { ...cellBase, background: "#d9d9d9", fontWeight: 700, textAlign: "center", verticalAlign: "middle" };
  const valCell = { ...cellBase, background: "#fff", verticalAlign: "middle" };
  const lblCell = { ...cellBase, background: "#e8e8e8", fontWeight: 700, verticalAlign: "middle" };
  const boldBorder = { border: "2px solid #000" };
  const live = (v: string) => (v && v !== "—") ? "#15803d" : "#999";

  const BORDER = "1px solid #999";
  const BORDER2 = "2px solid #000";

  const C: React.CSSProperties = { border: BORDER, padding: "1px 1.5px", fontSize: "7pt", verticalAlign: "top", wordBreak: "break-word", lineHeight: "1.4" };
  const H: React.CSSProperties = { ...C, background: "#d9d9d9", fontWeight: 700, textAlign: "center", verticalAlign: "middle", fontSize: "6.5pt" };
  const L: React.CSSProperties = { ...C, background: "#e8e8e8", fontWeight: 700, fontSize: "6.5pt", verticalAlign: "middle", lineHeight: "1" };
  const V: React.CSSProperties = { ...C, background: "#fff", verticalAlign: "middle" };
  const T: React.CSSProperties = { width: "100%", borderCollapse: "collapse" as const, tableLayout: "fixed" as const, fontSize: "7pt", lineHeight: "1" };


  return (
    <div style={{ width: "100%", minWidth: 900, fontFamily: "Arial,sans-serif", fontSize: "7pt", background: "#fff", color: "#000", overflowX: "auto" }}>
      <div data-no-print="true" style={{ marginBottom: "10px", display: "flex", justifyContent: "flex-end", gap: "8px", padding: "5px" }}>
        <select
          value={selectedLang}
          onChange={(e) => handleTranslateOPS(e.target.value)}
          disabled={translating}
          style={{
            padding: "4px 8px",
            fontSize: "10px",
            fontWeight: "bold",
            border: "1px solid #ccc",
            borderRadius: "4px",
            backgroundColor: "#fff",
            cursor: "pointer",
          }}
        >
          {INDIAN_LANGUAGES.map(lang => (
            <option key={lang.code} value={lang.code}>{lang.label}</option>
          ))}
        </select>
        {translating && (
          <span style={{ fontSize: "10px", color: "#2563eb", display: "flex", alignItems: "center", gap: "4px" }}>
            <span className="w-3 h-3 border-2 border-blue-500 border-t-transparent rounded-full animate-spin inline-block"></span>
            Translating...
          </span>
        )}
      </div>

      {/* RETENTION BAR */}
      <table style={{ width: "100%", borderCollapse: "collapse", border: "2px solid #000" }}>
        <tbody>
          <tr>
            <td style={{ padding: "1px 6px", textAlign: "right", fontWeight: 700, fontSize: "6.5pt" }}>
              Retention Period : 20 years after Model is discontinued
            </td>
          </tr>
        </tbody>
      </table>

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
            <td rowSpan={8} style={{ border: BORDER2, padding: 0, background: "#ffffffff" }}> {/* Reduced padding */}
              <img src={ASSETS.logo} alt="Logo" style={{ width: "35.93mm", height: "16.5mm", verticalAlign: "top" }} /> {/* Reduced from 80 */}
            </td>
            <td style={{ ...L, marginBottom: 0, lineHeight: "1.5", width: "2%" }}>{label(headerQuestions, 0, "Dept. / Section")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(deptCombined), marginBottom: 0, lineHeight: "1.2", width: "10%" }}>
              {deptCombined || "—"}
            </td>
            <td style={{ ...L, marginBottom: 0, lineHeight: "1.5", width: "2.5%" }}>{label(headerQuestions, 1, "Line / Zones")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(lineCombined), marginBottom: 0, lineHeight: "1.2", width: "12%" }}>
              {lineCombined || "—"}
            </td>
            <td colSpan={4} style={{ border: BORDER2, textAlign: "center", verticalAlign: "middle", padding: 2 }}> {/* Reduced padding */}
              <div style={{ fontSize: "9pt", fontWeight: 700, letterSpacing: 1 }}>Operation Standard</div> {/* Reduced from 14pt */}
            </td>
            {/* Prepared, Checked, Approved — plain empty cells (no ruled lines) */}
            {Array.from({ length: 3 }).map((_, idx) => (
              <td key={`empty-${idx}`} rowSpan={7} style={{ border: BORDER2, verticalAlign: "top", padding: 0 }} />
            ))}
            {/* SUBMISSION HISTORY - Paginated */}
            {Array.from({ length: 3 }).map((_, idx) => {
              const colIdx = idx + 4;

              // Determine how many rows to render — at least 5, or all entries if more
              const rowCount = Math.max(5, paginatedSubmissionHistory.length);

              return (
                <td key={idx} rowSpan={7} style={{ border: BORDER2, verticalAlign: "top", padding: 0, background: "#fff" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}> {/* ← Remove height: "100%" */}
                    {Array.from({ length: rowCount }).map((_, i) => {
                      const emptyRows = rowCount - paginatedSubmissionHistory.length;

                      const entry =
                        i >= emptyRows
                          ? paginatedSubmissionHistory[
                          paginatedSubmissionHistory.length - 1 - (i - emptyRows)
                          ]
                          : null;

                      let display: string = "\u00A0";

                      if (entry) {
                        if (colIdx === 4) display = String(entry.no);
                        else if (colIdx === 5) {
                          try {
                            display = new Date(entry.date).toLocaleDateString("en-GB", {
                              day: "2-digit",
                              month: "2-digit",
                              year: "2-digit",
                            });
                          } catch {
                            display = entry.date || "\u00A0";
                          }
                        } else {
                          display = entry.issuanceDetails || "\u00A0";
                        }
                      }

                      return (
                        <tr key={i}>
                          <td
                            style={{
                              borderBottom: "1px solid #ccc",
                              height: 30,           // ← Each row stays 36px
                              minHeight: 30,        // ← Enforce minimum
                              padding: "0 2px",
                              fontSize: "5pt",
                              textAlign: "center",
                              color: entry ? "#15803d" : "transparent",
                              fontWeight: entry ? 700 : 400,
                              whiteSpace: "nowrap",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                            }}
                          >
                            {display}
                            {entry?.responseIndex !== undefined &&
                              sameFormatResponses.length > 1 && (
                                <span style={{ fontSize: "4pt", marginLeft: 1, color: "#b45309" }}>
                                  R{entry.responseIndex + 1}
                                </span>
                              )}
                          </td>
                        </tr>
                      );
                    })}
                  </table>
                </td>
              );
            })}
            <td rowSpan={8} style={{ border: BORDER2, verticalAlign: "top", padding: "2px 3px", fontSize: "5.5pt" }}>
              {/* Format No. - Use actual question text from form */}

              <div style={{ fontWeight: 700, color: "#c00", lineHeight: "1", marginBottom: 2, }}>
                {label(docControlQuestions, 4, "Format No. AP")} :
              </div>
              <div style={{ fontWeight: 700, fontSize: "7pt", lineHeight: "1.2", marginBottom: 1, color: live(formatNoCombined) }}>
                {formatNoCombined || "—"}
              </div>


              <div style={{ borderTop: "0.5px solid #999", margin: "2px 0" }} />

              {/* Control No. - Use actual question text from form */}
              <div style={{ fontWeight: 700, color: "#c00", lineHeight: "1", marginBottom: 1, }}>
                {label(docControlQuestions, 5, "Control No. AP")} :
              </div>
              <div style={{ fontWeight: 700, fontSize: "7pt", marginBottom: 1, lineHeight: "1.2", color: live(controlNoCombined) }}>
                {controlNoCombined || "—"}
              </div>

              <div style={{ borderTop: "0.5px solid #999", margin: "2px 0" }} />

              {/* QR Code - This might be a static label or from form */}
              <div style={{ fontWeight: 700, marginBottom: 1 }}>QR Code :</div>
              <img src={ASSETS.qr} alt="QR" style={{ width: 55, height: 30, objectFit: "contain" }} />
            </td>
          </tr>

          <tr style={{ height: 4 }}> {/* Reduced from 13 */}
            <td style={{ ...L, lineHeight: "1.5" }}>{label(headerQuestions, 2, "Model AP")}</td>
            <td style={{ ...V, fontWeight: 700, color: live(modelCombined) }}>
              {modelCombined || "—"}
            </td>
            <td style={{ ...L, lineHeight: "1.5" }}>{label(headerQuestions, 3, "Process / Station AP")} :</td>
            <td style={{ ...V, fontWeight: 700, color: live(stationCombined), lineHeight: "1.2" }}>
              {stationCombined || "—"}
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
              <img src={ASSETS.stop} alt="Stop Call Wait" style={{ width: "100%", height: "100%", objectFit: "contain" }} /> {/* Reduced from 180 */}
            </td>
            <td style={{ ...H, }}>S. No.</td>
            <td style={H}>Trouble</td>
            <td style={H}>Your task</td>
          </tr>

          <tr style={{ height: 4 }}> {/* Reduced from 12 */}
            <td style={{ ...C, textAlign: "center" }}>1</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{TROUBLE_ROWS[0]}</td> {/* Reduced from 6pt */}
            <td rowSpan={5} style={{ fontSize: "5.5pt", textAlign: "center", verticalAlign: "middle", lineHeight: "2" }}> {/* Reduced from 5.5pt */}
              Stop The Line<br />Inform the Zone Leader<br />Write on card if mentioned in OPS
            </td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>2</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{TROUBLE_ROWS[1]}</td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>3</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{TROUBLE_ROWS[2]}</td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>4</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{TROUBLE_ROWS[3]}</td>
          </tr>
          <tr style={{ height: 7 }}>
            <td style={{ ...C, textAlign: "center" }}>5</td>
            <td style={{ ...C, fontSize: "5.5pt" }}>{TROUBLE_ROWS[4]}</td>
            <td style={{ ...H, border: BORDER2, fontSize: "6pt" }}>Prepared</td>
            <td style={{ ...H, fontSize: "6pt" }}>Checked</td>
            <td style={{ ...H, fontSize: "6pt" }}>Approved</td>
            <td style={{ ...H, fontSize: "6pt" }}>No.</td>
            <td style={{ ...H, fontSize: "6pt" }}>DD/MM/YY</td>
            <td style={{ ...H, border: BORDER2, fontSize: "6pt" }}>Issuance / Revision details</td>
          </tr>
        </tbody>
      </table>

      {/* General Instructions Body - COMPACT */}
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
          <col style={{ width: "8%" }} />
          <col style={{ width: "17%" }} />
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
            <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
              <div style={{ fontWeight: 800, marginBottom: 0, fontSize: "5.5pt" }}>FIFO System</div>
              {DEF_FIFO.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, lineHeight: "1.4", fontSize: "5.5pt" }}>{l}</div>)}
            </td>
            <td style={{ ...C, verticalAlign: "top", padding: 0 }}>
              <div style={{ padding: "1px 2px", borderBottom: "0.5px solid #ccc", fontSize: "5pt", lineHeight: "1.4" }}>{DEF_NONLUB}</div>
              <div style={{ display: "flex", borderBottom: "0.5px solid #ccc" }}>
                <div style={{ flex: 1, borderRight: "0.5px solid #ccc", padding: "1px 2px", textAlign: "center", fontWeight: 700, fontSize: "4.5pt", background: "#d9d9d9", lineHeight: "1.4" }}>No mobile on shopfloor</div>
                <div style={{ flex: 1, padding: "1px 2px", textAlign: "center", fontWeight: 700, fontSize: "4.5pt", background: "#d9d9d9", lineHeight: "1.4" }}>Do not run on shopfloor</div>
              </div>
              <div style={{ display: "flex" }}>
                <div style={{ flex: 1, borderRight: "0.5px solid #ccc", padding: 1, textAlign: "center" }}>
                  <img src={ASSETS.noMob} alt="No Mobile" style={{ width: 44, height: 55, objectFit: "contain" }} />
                </div>
                <div style={{ flex: 1, padding: 1, textAlign: "center" }}>
                  <img src={ASSETS.noRun} alt="No Run" style={{ width: 44, height: 55, objectFit: "contain" }} />
                </div>
              </div>
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 1 }}>
              <img src={ASSETS.ppeG} alt="Full PPE Uniform" style={{ width: "100%", height: 95, objectFit: "contain" }} />
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 1 }}>
              <img src={ASSETS.ppeGl} alt="Station PPE" style={{ width: "100%", maxHeight: 95, objectFit: "contain" }} />
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "top", padding: 1 }}>
              <img src={ASSETS.shift} alt="Shift Timings" style={{ width: "100%", height: 90, objectFit: "contain" }} />
            </td>
            <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
              <div style={{ fontWeight: 700, color: "#166534", marginBottom: 1, fontSize: "5pt" }}>Environmental Issues</div>
              {DEF_ENV.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, fontSize: "5pt", lineHeight: "1.4" }}>{l}</div>)}
            </td>
            <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
              <div style={{ fontWeight: 700, color: "#991b1b", marginBottom: 1, fontSize: "5pt" }}>Safety Issues</div>
              {DEF_SAFE.split("\n").map((l, i) => <div key={i} style={{ marginBottom: 0, lineHeight: "1.4", fontSize: "5pt" }}>{l}</div>)}
            </td>
            <td style={{ ...C, textAlign: "center", verticalAlign: "middle", padding: 1 }}>
              <img src={ASSETS.fiveS} alt="5S Guidelines" style={{ width: "100%", height: 95, objectFit: "fill" }} />
            </td>
            <td style={{ ...C, verticalAlign: "top", fontSize: "5pt" }}>
              {DEF_PROC_INS.map((l, i) => (
                <div key={i} style={{ marginBottom: 0, lineHeight: "1.4", fontSize: "5pt" }}>
                  {l}
                </div>
              ))}
            </td>
          </tr>
        </tbody>
      </table>

      {/* PROCESS STEPS TABLE - WITH PAGINATION */}
      <table style={{ width: "100%", borderCollapse: "collapse", borderLeft: "2px solid #000", borderRight: "2px solid #000", borderBottom: "2px solid #000", tableLayout: "fixed" }}>
        <colgroup>
          <col style={{ width: "8%" }} />
          <col style={{ width: "2.5%" }} />
          {processStepColumns.map((col, i) => <col key={i} style={{ width: col.width }} />)}
        </colgroup>
        <thead>
          <tr>
            <th style={{ ...hdrCell, background: "#ffffffff", color: "#000" }}>
              {illustrationQuestions[0]?.text || illustrationQuestions[0]?.label || "Illustrations & Process Details"}
            </th>
            <th style={hdrCell}>SN</th>
            {processStepColumns.map((col) => (
              <th key={col.questionId} style={{ ...hdrCell, whiteSpace: "pre-line", lineHeight: 1.2 }}>
                {col.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {isLoadingSameFormatResponses && (
            <tr>
              <td colSpan={2 + processStepColumns.length}
                style={{ ...cellBase, textAlign: 'center', background: '#f0f9ff', color: '#3b82f6', fontStyle: 'italic', height: 50 }}>
                ⏳ Loading responses…
              </td>
            </tr>
          )}

          {/* EACH RESPONSE GETS ITS OWN ROW - USING PAGINATED RESPONSES */}
          {!isLoadingSameFormatResponses && paginatedResponses?.map((resp, respIdx) => {
            const globalIdx = (currentPage - 1) * responsesPerPage + respIdx;
            const illustrationImages = getResponseIllustrationImages(resp);

            return (
              <tr key={`resp-${globalIdx}`} style={{ minHeight: 50 }}>
                {/* Illustration column */}
                <td style={{ ...cellBase, background: "#ffffffff", textAlign: "center", verticalAlign: "middle", padding: 2, height: 50 }}>
                  {illustrationImages.length > 0 ? (
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', justifyContent: 'center' }}>
                      {illustrationImages.map((url, idx) => (
                        <img key={idx} src={url} alt={`Illustration ${idx + 1}`}
                          style={{ width: 50, height: 40, objectFit: 'contain', cursor: 'pointer', border: '0.5px solid #ccc', borderRadius: 2 }}
                          onClick={() => window.open(url, '_blank')} />
                      ))}
                    </div>
                  ) : (
                    <div style={{ width: "100%", height: 40, display: "flex", alignItems: "center", justifyContent: "center", color: "#bbb", fontSize: 20 }}>📷</div>
                  )}
                </td>

                {/* SN column - global response number */}
                <td style={{ ...cellBase, textAlign: "center", fontWeight: 700, fontSize: "9pt", verticalAlign: "middle" }}>
                  {globalIdx + 1}
                </td>

                {/* Process step answer cells */}
                {processStepColumns.map((col, colIdx) => {
                  const rawVal = resp?.answers?.[col.questionId];
                  let cellVal = '';
                  if (rawVal !== null && rawVal !== undefined && rawVal !== '') {
                    if (typeof rawVal === 'object') {
                      try { cellVal = JSON.stringify(rawVal); } catch { cellVal = String(rawVal); }
                    } else {
                      cellVal = String(rawVal);
                    }
                  }

                  // IMPORTANT: Use globalIdx (not respIdx) for consistent key when pagination changes
                  // Also use col.originalIndex to match the column's actual position in the original process questions
                  const translationKey = `${globalIdx}:${col.originalIndex ?? colIdx}`;
                  const translated = cellTranslations.get(translationKey);
                  const shouldShowTranslation = selectedLang !== 'en' &&
                    TRANSLATE_COL_INDICES.has(col.originalIndex ?? colIdx) &&
                    translated &&
                    translated !== cellVal;

                  return (
                    <td key={col.questionId} style={{
                      ...cellBase,
                      height: 50, minHeight: 50,
                      padding: 0,
                      verticalAlign: 'top',
                      background: '#fff',
                    }}>
                      {cellVal ? (
                        <>
                          {/* English top */}
                          <div style={{
                            padding: '2px 3px',
                            fontSize: '6pt',
                            lineHeight: 1.3,
                            borderBottom: shouldShowTranslation ? '1px dashed #bbb' : 'none',
                            minHeight: shouldShowTranslation ? 28 : undefined,
                          }}>
                            {cellVal}
                          </div>
                          {/* Translated bottom */}
                          {shouldShowTranslation && (
                            <div style={{
                              padding: '2px 3px',
                              fontSize: '5.5pt',
                              color: '#1a1a8c',
                              lineHeight: 1.3,
                              fontFamily: "'Noto Sans Devanagari', Arial, sans-serif",
                            }}>
                              {translated}
                            </div>
                          )}
                        </>
                      ) : '\u00A0'}
                    </td>
                  );
                })}
              </tr>
            );
          })}

          {/* Blank filler rows - adjusted for pagination */}
          {!isLoadingSameFormatResponses && paginatedResponses && paginatedResponses.length < responsesPerPage &&
            Array.from({ length: responsesPerPage - paginatedResponses.length }).map((_, i) => (
              <tr key={`blank-${i}`} style={{ minHeight: 50 }}>
                <td style={{ ...cellBase, background: "#ffffffff", height: 50 }} />
                <td style={{ ...cellBase, textAlign: 'center', fontWeight: 700, fontSize: '9pt', verticalAlign: 'middle', color: '#ccc' }}>
                  {(currentPage - 1) * responsesPerPage + paginatedResponses.length + i + 1}
                </td>
                {processStepColumns.map(col => (
                  <td key={col.questionId} style={{ ...cellBase, height: 50, background: "#fff" }}>&nbsp;</td>
                ))}
              </tr>
            ))
          }
        </tbody>
      </table>

      {/* ABNORMALITY + PAST PROBLEMS */}
      < table style={{ width: "100%", borderCollapse: "collapse", borderLeft: "2px solid #000", borderRight: "2px solid #000", borderBottom: "2px solid #000", tableLayout: "fixed" }}>
        <tbody>
          <tr>
            <td rowSpan={2} style={{ ...cellBase, padding: 4, verticalAlign: "top", fontSize: "6.5pt", width: "12%" }}>
              <div style={{ fontWeight: 700, marginBottom: 2 }}>
                {label(pastProbsQuestions, 0, "Abnormality handling route")} :
              </div>
              <div>In case of any abnormality inform the Zone In-Charge</div>
              <div style={{ marginTop: 2 }}>Flow of Communication :-</div>
              <div>Operator ▶ Team Member ▶ Section Mgr ▶ As required</div>
            </td>
            <td style={{ ...lblCell, padding: "2px", textAlign: "center", fontSize: "7pt" }}>
              {label(pastProbsQuestions, 1, "Past Problem Details")}
            </td>
          </tr>
          <tr>
            <td style={{
              ...cellBase, padding: 4, verticalAlign: "top", minHeight: 60, height: 60, fontSize: "7pt", background: "#fff",
              fontStyle: pastProblemMerged.isMerged ? "italic" : "normal",
              color: pastProblemMerged.value ? (pastProblemMerged.isMerged ? "#b45309" : "#000") : "#bbb"
            }}>
              {pastProblemMerged.value || <span style={{ color: "#ccc", fontStyle: "italic" }}>No data recorded</span>}
              {pastProblemMerged.isMerged && (
                <div style={{ fontSize: "5.5pt", color: "#9a6700", marginTop: 2 }}>
                  ({sameFormatResponses.length} responses merged)
                </div>
              )}
            </td>
          </tr>
        </tbody>
      </table>

      {/* ASSOCIATE NAME & SIGN */}
      <table style={{ width: "100%", borderCollapse: "collapse", borderLeft: "2px solid #000", borderRight: "2px solid #000", borderBottom: "2px solid #000", tableLayout: "fixed" }}>
        <colgroup>
          <col style={{ width: "5%" }} />
          {Array.from({ length: 22 }).map((_, i) => <col key={i} style={{ width: `${(95 / 22).toFixed(2)}%` }} />)}
        </colgroup>
        <tbody>
          <tr>
            <td style={{ ...lblCell, textAlign: "center", fontSize: "5.5pt", border: boldBorder, verticalAlign: "middle" }}>
              Associate Name &amp; Emp. Code
            </td>
            {Array.from({ length: 22 }).map((_, i) => <td key={i} style={{ border: "1px solid #999", height: 22 }} />)}
          </tr>
          <tr>
            <td style={{ ...lblCell, textAlign: "center", fontSize: "5.5pt", border: boldBorder, verticalAlign: "middle" }}>
              Sign &amp; Date
            </td>
            {Array.from({ length: 22 }).map((_, i) => <td key={i} style={{ border: "1px solid #999", height: 26 }} />)}
          </tr>
        </tbody>
      </table>

      {/* PAGE NUMBER */}
      <table style={{ width: "100%", borderCollapse: "collapse", borderLeft: "2px solid #000", borderRight: "2px solid #000", borderBottom: "2px solid #000", tableLayout: "fixed" }}>
        <tbody>
          <tr>
            <td style={{ ...cellBase, padding: 2 }}>{form?.title || "Operation Standard"}</td>
            <td style={{ ...hdrCell, padding: 3, fontSize: "8.5pt" }}>Page Number : XX / XX</td>
          </tr>
        </tbody>
      </table>
      {/* PAGINATION CONTROLS */}
      {!isLoadingSameFormatResponses && sameFormatResponses.length > responsesPerPage && (
        <div style={{
          marginTop: "20px",
          padding: "10px",
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          gap: "8px",
          borderTop: "2px solid #000",
          borderLeft: "2px solid #000",
          borderRight: "2px solid #000",
          borderBottom: "2px solid #000",
          backgroundColor: "#f5f5f5"
        }}>
          <button
            onClick={() => handlePageChange(currentPage - 1)}
            disabled={currentPage === 1}
            style={{
              padding: "6px 12px",
              backgroundColor: currentPage === 1 ? "#ccc" : "#1d4ed8",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: currentPage === 1 ? "not-allowed" : "pointer",
              fontSize: "10pt",
              fontWeight: "bold"
            }}
          >
            ← Previous
          </button>

          <div style={{ display: "flex", gap: "5px" }}>
            {Array.from({ length: totalPages }, (_, i) => i + 1).map(page => {
              // Show limited page numbers for better UX
              if (totalPages <= 10) {
                return (
                  <button
                    key={page}
                    onClick={() => handlePageChange(page)}
                    style={{
                      padding: "6px 12px",
                      backgroundColor: currentPage === page ? "#1d4ed8" : "#e5e7eb",
                      color: currentPage === page ? "white" : "#333",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer",
                      fontSize: "10pt",
                      fontWeight: currentPage === page ? "bold" : "normal"
                    }}
                  >
                    {page}
                  </button>
                );
              } else {
                // Smart pagination for many pages
                if (page === 1 || page === totalPages || (page >= currentPage - 2 && page <= currentPage + 2)) {
                  return (
                    <button
                      key={page}
                      onClick={() => handlePageChange(page)}
                      style={{
                        padding: "6px 12px",
                        backgroundColor: currentPage === page ? "#1d4ed8" : "#e5e7eb",
                        color: currentPage === page ? "white" : "#333",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "10pt"
                      }}
                    >
                      {page}
                    </button>
                  );
                }
                if (page === currentPage - 3 || page === currentPage + 3) {
                  return <span key={`ellipsis-${page}`} style={{ padding: "6px 8px" }}>...</span>;
                }
                return null;
              }
            })}
          </div>

          <button
            onClick={() => handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages}
            style={{
              padding: "6px 12px",
              backgroundColor: currentPage === totalPages ? "#ccc" : "#1d4ed8",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: currentPage === totalPages ? "not-allowed" : "pointer",
              fontSize: "10pt",
              fontWeight: "bold"
            }}
          >
            Next →
          </button>

          <span style={{ marginLeft: "15px", fontSize: "9pt", color: "#666" }}>
            Page {currentPage} of {totalPages} |
            Total {sameFormatResponses.length} response{sameFormatResponses.length !== 1 ? 's' : ''}
          </span>
        </div>
      )}

      {/* FULLSCREEN PREVIEW BUTTON - Add near the print button area */}
      <div data-no-print="true" style={{ marginTop: "10px", marginBottom: "10px", textAlign: "center" }}>
        <button
          onClick={() => setShowFullscreenPreview(true)}
          style={{
            padding: "8px 20px",
            backgroundColor: "#6b21a5",
            color: "white",
            border: "none",
            borderRadius: "6px",
            cursor: "pointer",
            fontSize: "11pt",
            fontWeight: "bold",
            display: "inline-flex",
            alignItems: "center",
            gap: "8px"
          }}
        >
          <Eye size={16} />
          Fullscreen Preview (Slideshow)
        </button>
      </div>


      {/* Render fullscreen preview modal */}
      {showFullscreenPreview && <FullscreenPreview />}
    </div>
  );
}

export default function ResponseDetailsPage() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const { showSuccess, showError, showConfirm } = useNotification();
  const { logo } = useLogo();


  const [response, setResponse] = useState<Response | null>(null);
  const [form, setForm] = useState<Form | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<"dashboard" | "responses" | "ops">(
    "dashboard"
  );
  const [updatingStatus, setUpdatingStatus] = useState(false);
  const [showStatusUpdate, setShowStatusUpdate] = useState(false);
  const [exportingExcel, setExportingExcel] = useState(false);
  const [generatingPDF, setGeneratingPDF] = useState(false);
  const [exportingZip, setExportingZip] = useState(false);
  const [selectedPDFType, setSelectedPDFType] = useState<
    "no-only" | "yes-only" | "both" | "na-only" | "section" | "default" | "responses-view" | null
  >(null);
  const [editingResponse, setEditingResponse] = useState<Response | null>(null);
  const [editingForm, setEditingForm] = useState<Form | null>(null);
  const [savingEdit, setSavingEdit] = useState(false);
  const [editingFormLoading, setEditingFormLoading] = useState(false);
  const [expandResponseRateBreakdown, setExpandResponseRateBreakdown] =
    useState(false);
  const [showPDFTypeSelector, setShowPDFTypeSelector] = useState(false);
  const sectionRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const isCancelledRef = useRef<boolean>(false);
  const [pendingSectionId, setPendingSectionId] = useState<string | null>(null);
  const [sectionChartTypes, setSectionChartTypes] = useState<Record<string, "pie" | "bar">>({});
  const [showMainParamsImages, setShowMainParamsImages] = useState<Record<string, boolean>>({});
  const [showSectionsPDFModal, setShowSectionsPDFModal] = useState(false);
  const [downloadingSectionsPDF, setDownloadingSectionsPDF] = useState(false);
  const [autoOpenSectionId, setAutoOpenSectionId] = useState<string | null>(null);

  const [showUpdatePopup, setShowUpdatePopup] = useState(false);
  const [popupDate, setPopupDate] = useState("");
  const [popupIssuanceDetails, setPopupIssuanceDetails] = useState("");
  const [pendingUpdate, setPendingUpdate] = useState<any>(null);
  const [submissionHistory, setSubmissionHistory] = useState<any[]>([]);
  const opsPrintRef = useRef<HTMLDivElement>(null);

  // Add these state variables near your other state declarations (around line 140)
  const [allResponses, setAllResponses] = useState<any[]>([]);
  const [sameFormatResponses, setSameFormatResponses] = useState<any[]>([]);
  const [isLoadingSameFormatResponses, setIsLoadingSameFormatResponses] = useState(false);
  // Translation state


  // Add this after setting form state (around line 200-250)
  const effectiveOpsMapping = useMemo(() => {
    if (!form?.sections) return null;

    const mapping = {
      headerSectionId: "sec_basic_info",
      generalInstructionsSectionId: "sec_doc_control",
      pastProblemsSectionId: "sec_abnormalities_handling",
      processStepsSectionId: "sec_process_steps",
      associateSignSectionId: "sec_associate_sign",
      illustrationsSectionId: "sec_illustrations",
    };

    return mapping;
  }, [form]);

  // Helper to get questions from a section
  const getQuestionsFromSection = useCallback((sectionId: string | undefined) => {
    if (!form || !sectionId) return [];
    const section = form.sections?.find((s: any) => s.id === sectionId || s._id === sectionId);
    if (!section) return [];
    return (section.questions || []).filter((q: any) => !q.showWhen?.questionId);
  }, [form]);

  // Get the instructions questions (from sec_doc_control section)
  const instructionsQuestions = useMemo(() => {
    return getQuestionsFromSection("sec_doc_control");
  }, [getQuestionsFromSection]);

  // Get the header questions (from sec_basic_info section)
  const headerQuestionsForFormat = useMemo(() => {
    return getQuestionsFromSection("sec_basic_info");
  }, [getQuestionsFromSection]);


  // Add this function to fetch all responses for the same form
  const fetchAllResponsesForForm = useCallback(async () => {
    if (!form) return [];

    try {
      const responsesData = await apiClient.getResponses();
      const formIdentifier = response?.questionId || response?.formId;

      // Filter responses for the same form
      const formResponses = responsesData.responses.filter((r: any) => {
        const rFormId = r.questionId || r.formId;
        return rFormId === formIdentifier || String(rFormId) === String(formIdentifier);
      });

      setAllResponses(formResponses);
      return formResponses;
    } catch (err) {
      console.error("Failed to fetch all responses:", err);
      return [];
    }
  }, [form, response]);

  // Add this function to get the Format No value from a response
  // Add this function to get the Format No value from a response
  const getFormatNoFromResponse = useCallback((resp: any): string => {
    if (!form) return "";

    // Get the instructions section (sec_doc_control)
    const instructionsSection = form.sections?.find(
      (s: any) => s.id === "sec_doc_control"
    );

    // Format No is typically the first question in this section
    const formatQuestion = instructionsSection?.questions?.[0];

    if (!formatQuestion) return "";

    const formatAnswer = resp.answers?.[formatQuestion.id];
    if (formatAnswer === undefined || formatAnswer === null || formatAnswer === "") return "";
    return String(formatAnswer);
  }, [form]);
  const [isGroupedView, setIsGroupedView] = useState(false);
  const [groupedResponses, setGroupedResponses] = useState<any[]>([]);



  // Update the filterSameFormatResponses function
  const filterSameFormatResponses = useCallback(async () => {
    if (!response || !form) return;

    setIsLoadingSameFormatResponses(true);

    try {
      // If we're in grouped view, use the grouped responses directly
      if (isGroupedView && groupedResponses.length > 0) {
        // Sort by creation date (oldest first)
        const sorted = [...groupedResponses].sort((a: any, b: any) => {
          return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
        });
        setSameFormatResponses(sorted);
        setIsLoadingSameFormatResponses(false);
        return;
      }

      // Original logic for single response
      let formResponses = allResponses;
      if (formResponses.length === 0) {
        formResponses = await fetchAllResponsesForForm();
      }

      const currentFormatNo = getFormatNoFromResponse(response);

      if (!currentFormatNo) {
        setSameFormatResponses([response]);
        return;
      }

      const filtered = formResponses.filter((resp: any) => {
        const formatNo = getFormatNoFromResponse(resp);
        return formatNo === currentFormatNo;
      });

      const sorted = filtered.sort((a: any, b: any) => {
        return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
      });

      setSameFormatResponses(sorted);
    } catch (err) {
      console.error("Failed to filter same format responses:", err);
      setSameFormatResponses([response]);
    } finally {
      setIsLoadingSameFormatResponses(false);
    }
  }, [response, form, allResponses, fetchAllResponsesForForm, getFormatNoFromResponse, isGroupedView, groupedResponses]);
  // Call this after fetching the response and form
  useEffect(() => {
    if (response && form) {
      filterSameFormatResponses();
    }
  }, [response, form]);

  const [currentPage, setCurrentPage] = useState(1);




  const [pdfProgress, setPdfProgress] = useState<{
    stage: 'uploading' | 'generating' | 'downloading' | 'complete' | 'error';
    percentage: number;
    message?: string;
  } | null>(null);

  const [pdfDownloadProgress, setPdfDownloadProgress] = useState<number | null>(null);

  useEffect(() => {
    fetchResponseDetails();
  }, [id]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as HTMLElement;
      if (!target.closest('.pdf-type-selector')) {
        setShowPDFTypeSelector(false);
      }
    };

    document.addEventListener('click', handleClickOutside);
    return () => {
      document.removeEventListener('click', handleClickOutside);
    };
  }, []);

  useEffect(() => {
    if (location.state?.viewMode) {
      setViewMode(location.state.viewMode);
    }
  }, [location.state]);

  useEffect(() => {
    if (autoOpenSectionId) {
      const element = document.getElementById(`section-detail-${autoOpenSectionId}`);
      if (element) {
        element.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    }
  }, [autoOpenSectionId]);

  useEffect(() => {
    if (form?.sections) {
      const initialShowImages: Record<string, boolean> = {};
      form.sections.forEach((section: any) => {
        initialShowImages[section.id] = true;
      });
      setShowMainParamsImages(initialShowImages);
    }
  }, [form]);

  const complianceLabels = useMemo(() => {
    const defaultLabels = { yes: "Yes", no: "No", na: "N/A", correct: "Correct", wrong: "Wrong" };
    let labels = { ...defaultLabels };

    // Check if any question in the form is a zone/chassis type
    const hasSpecialTypes = form?.sections?.some((s) =>
      s.questions?.some((q: any) =>
        [
          "chassis-with-zone",
          "zone-in",
          "zone-out",
          "chassis-without-zone",
          "chassisNumber",
        ].includes(q.type),
      ),
    );

    if (hasSpecialTypes) {
      return { yes: "Accepted", no: "Rejected", na: "Rework", correct: "Accepted", wrong: "Rejected" };
    }

    // Check if form is a quiz/accuracy type
    const hasAccuracyQuestions = form?.sections?.some((s: any) =>
      s.questions?.some((q: any) =>
        !["yesNoNA", "chassisNumber", "chassis-with-zone", "chassis-without-zone", "zone-in", "zone-out"].includes(q.type)
      )
    );

    if (hasAccuracyQuestions && !form?.sections?.some((s: any) => s.questions?.some((q: any) => q.type === "yesNoNA"))) {
      return { yes: "Correct", no: "Wrong", na: "N/A", correct: "Correct", wrong: "Wrong" };
    }

    if (form?.sections) {
      // First pass: look for any question that has non-default labels
      for (const section of form.sections) {
        if (section.questions) {
          for (const q of section.questions) {
            // Check for yesNoNA type and at least 2 options
            if (q.type === "yesNoNA" && q.options && q.options.length >= 2) {
              const hasCustomLabels =
                q.options[0] !== "Yes" ||
                q.options[1] !== "No" ||
                (q.options[2] && q.options[2] !== "N/A");

              if (hasCustomLabels) {
                return {
                  yes: q.options[0] || "Yes",
                  no: q.options[1] || "No",
                  na: q.options[2] || "N/A",
                  correct: q.options[0] || "Correct",
                  wrong: q.options[1] || "Wrong"
                };
              }

              // If we haven't found custom labels yet, store the first yesNoNA labels we find as fallback
              if (labels.yes === "Yes") {
                labels.yes = q.options[0] || "Yes";
                labels.no = q.options[1] || "No";
                labels.na = q.options[2] || "N/A";
                labels.correct = q.options[0] || "Correct";
                labels.wrong = q.options[1] || "Wrong";
              }
            }
          }
        }
      }
    }
    return labels;
  }, [form]);
  // Add these state variables near other state declarations (around line 140)
  const [groupModelNo, setGroupModelNo] = useState("");
  // In ResponseDetailsPage.tsx, update the fetchResponseDetails function:

  const fetchResponseDetails = async () => {
    try {
      setLoading(true);
      if (!id) {
        throw new Error("Response ID is required");
      }

      // Try to get state from multiple sources
      let navigationState = location.state;

      // Also check window.history.state
      if (!navigationState && window.history.state?.usr) {
        navigationState = window.history.state.usr;
        console.log("Got state from window.history.state:", navigationState);
      }

      // If still no state, check sessionStorage
      let groupedData = null;
      if (!navigationState?.groupedResponses) {
        const storageKey = `grouped_${id}`;
        const stored = sessionStorage.getItem(storageKey);
        if (stored) {
          groupedData = JSON.parse(stored);
          console.log("Got grouped data from sessionStorage:", groupedData);
          sessionStorage.removeItem(storageKey); // Clean up
        }
      }

      const hasGroupedResponses = navigationState?.groupedResponses || groupedData?.groupedResponses;

      console.log("=== RESPONSE DETAILS PAGE ===");
      console.log("Navigation state:", navigationState);
      console.log("Grouped data from storage:", groupedData);
      console.log("Has groupedResponses:", !!hasGroupedResponses);

      if (hasGroupedResponses && hasGroupedResponses.length > 0) {
        // Use the grouped responses
        const grouped = navigationState?.groupedResponses || groupedData?.groupedResponses;
        console.log("Setting grouped responses:", grouped.length);
        setGroupedResponses(grouped);
        setIsGroupedView(true);
        setGroupModelNo(navigationState?.modelNo || groupedData?.modelNo || "");

        // Use the first response as the main response for metadata
        const mainResponse = grouped[0];

        const formIdentifier = mainResponse.questionId || mainResponse.formId;
        if (!formIdentifier) {
          throw new Error("Missing form identifier for response");
        }

        const formData = await apiClient.getForm(formIdentifier);
        const selectedForm = formData.form;

        // Get submissionHistory from first response's answers
        const history = mainResponse.answers?.__submissionHistory || [];

        setResponse(mainResponse);
        setForm(selectedForm);
        setSubmissionHistory(history);

        // Set sameFormatResponses directly from grouped responses
        const sorted = [...grouped].sort((a: any, b: any) => {
          return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
        });
        setSameFormatResponses(sorted);

        console.log("Set sameFormatResponses with", sorted.length, "responses");

      } else {
        // Original single response logic
        console.log("No grouped responses found, loading single response");
        const responsesData = await apiClient.getResponses();
        const selectedResponse = responsesData.responses.find(
          (r: any) => r._id === id || r.id === id || String(r._id) === id || String(r.id) === id
        );

        if (!selectedResponse) {
          throw new Error(`Response with ID "${id}" not found.`);
        }

        const formIdentifier = selectedResponse.questionId || selectedResponse.formId;
        const formData = await apiClient.getForm(formIdentifier);
        const selectedForm = formData.form;
        const history = selectedResponse.answers?.__submissionHistory || [];

        setResponse(selectedResponse);
        setForm(selectedForm);
        setSubmissionHistory(history);
        setGroupedResponses([selectedResponse]);
        setIsGroupedView(false);
        setGroupModelNo("");
        setSameFormatResponses([selectedResponse]);
      }

    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load response");
      console.error("Error loading response details:", err);
    } finally {
      setLoading(false);
    }
  };

  const handleStatusUpdate = async (newStatus: string) => {
    if (!response) return;

    setUpdatingStatus(true);
    try {
      await apiClient.updateResponse(response._id, { status: newStatus });
      setResponse({ ...response, status: newStatus });
      setShowStatusUpdate(false);
      showSuccess(`Status updated to ${getStatusInfo(newStatus).label}`);
    } catch (err) {
      console.error("Failed to update status:", err);
      showError("Failed to update status. Please try again.");
    } finally {
      setUpdatingStatus(false);
    }
  };

  const handleEditResponse = async () => {
    if (!response) return;

    setEditingResponse(response);
    setEditingForm(null);
    setEditingFormLoading(true);
    try {
      const formIdentifier = response.questionId || response.formId;
      if (!formIdentifier) {
        throw new Error("Missing form identifier for response");
      }
      const formData = await apiClient.getForm(formIdentifier);
      const loadedForm = formData.form;

      if (loadedForm?.sections) {
        loadedForm.sections.forEach((section: any) => {
          if (section.questions) {
            section.questions.forEach((question: any) => {
              if (!Array.isArray(question.followUpQuestions)) {
                question.followUpQuestions = [];
              }
            });
          }
        });
      }

      if (!Array.isArray(loadedForm.followUpQuestions)) {
        loadedForm.followUpQuestions = [];
      }

      setEditingForm(loadedForm);
    } catch (err) {
      console.error("Failed to load form for editing:", err);
      showError("Failed to load form for editing. Please try again.");
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
    if (savingEdit || !response) return;

    setPendingUpdate(updated);
    setPopupDate(new Date().toLocaleDateString("en-GB"));
    setPopupIssuanceDetails(`Update of response`);
    setShowUpdatePopup(true);
  };

  const handleConfirmUpdate = async () => {
    if (!pendingUpdate || !response) return;

    setSavingEdit(true);
    try {
      // Get the ID from different possible sources
      const responseId = response._id || response.id;
      console.log("=== UPDATE DEBUG ===");
      console.log("response._id:", response._id);
      console.log("response.id:", response.id);
      console.log("URL param id (from useParams):", id);
      console.log("Using responseId:", responseId);

      // Try using the URL param directly instead
      const urlId = id; // This comes from useParams
      console.log("Would URL param work?:", urlId);

      const newHistoryEntry = {
        date: popupDate,
        issuanceDetails: popupIssuanceDetails
      };

      const existingHistory = response.answers?.__submissionHistory || submissionHistory || [];
      const newHistory = [...existingHistory, newHistoryEntry];

      const updatedAnswers = {
        ...pendingUpdate.answers,
        __submissionHistory: newHistory
      };

      const updateData = {
        answers: updatedAnswers
      };

      // Try with URL param ID instead
      console.log("Attempting update with URL param ID:", urlId);
      await apiClient.updateResponse(urlId!, updateData);

      setResponse({
        ...pendingUpdate,
        answers: updatedAnswers,
        submissionHistory: newHistory
      });
      setSubmissionHistory(newHistory);

      handleCloseEdit();
      showSuccess("Response updated successfully.");
      setShowUpdatePopup(false);
    } catch (err) {
      console.error("Failed to save response:", err);
      showError("Failed to save response. Please try again.");
    } finally {
      setSavingEdit(false);
    }
  };
  const handleExportExcel = async (type?: 'yes-only' | 'no-only' | 'na-only' | 'both' | 'default') => {
    if (!response || !form) return;

    setExportingExcel(true);
    try {
      // If type is not provided, default to 'default' (full report)
      const exportType = type || 'default';

      const fileName = `${form.title}_${exportType !== 'default' ? exportType + '_' : ''}${formatTimestamp(
        response.createdAt
      )}.xlsx`;

      await generateResponseExcelReport([response], form, fileName, exportType);
      showSuccess("Excel file downloaded successfully.");
    } catch (err) {
      console.error("Failed to export Excel:", err);
      showError("Failed to export Excel. Please try again.");
    } finally {
      setExportingExcel(false);
    }
  };

  const handleDeleteResponse = async () => {
    if (
      !window.confirm(
        "Are you sure you want to delete this response? This action cannot be undone."
      )
    ) {
      return;
    }

    try {
      await apiClient.deleteResponse(id!);
      showSuccess("Response deleted successfully");
      navigate("/responses/all");
    } catch (err) {
      console.error("Failed to delete response:", err);
      showError("Failed to delete response");
    }
  };
  const handlePrintOPS = useCallback(() => {

    const printContent = () => {
      const el = opsPrintRef.current;
      if (!el) return;

      const pw = window.open("", "_blank", "width=1600,height=1000");
      if (!pw) {
        alert("Please allow popups to save as PDF.");
        return;
      }

      // Serialize all inline styles from computed styles for every element
      const cloneWithComputedStyles = (source: Element, target: Element) => {
        const computed = window.getComputedStyle(source);
        const el = target as HTMLElement;
        const src = source as HTMLElement;

        // Copy critical layout properties explicitly
        const props = [
          'width', 'min-width', 'max-width',
          'height', 'min-height',
          'display', 'position',
          'font-size', 'font-family', 'font-weight', 'line-height',
          'padding', 'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
          'margin', 'margin-top', 'margin-right', 'margin-bottom', 'margin-left',
          'border', 'border-top', 'border-right', 'border-bottom', 'border-left',
          'border-collapse', 'border-spacing',
          'background-color', 'color',
          'text-align', 'vertical-align',
          'white-space', 'overflow', 'overflow-wrap', 'word-break',
          'flex-direction', 'flex-wrap', 'align-items', 'justify-content', 'gap',
          'box-sizing', 'table-layout',
        ];

        props.forEach(prop => {
          try {
            const val = computed.getPropertyValue(prop);
            if (val) el.style.setProperty(prop, val, 'important');
          } catch { }
        });

        // For tables, also capture exact pixel widths of cells
        if (source.tagName === 'TD' || source.tagName === 'TH' || source.tagName === 'COL') {
          const rect = source.getBoundingClientRect();
          el.style.setProperty('width', `${rect.width}px`, 'important');
          el.style.setProperty('min-width', `${rect.width}px`, 'important');
          el.style.setProperty('max-width', `${rect.width}px`, 'important');
        }

        if (source.tagName === 'IMG') {
          const imgSrc = source as HTMLImageElement;
          const imgTarget = target as HTMLImageElement;
          imgTarget.src = imgSrc.currentSrc || imgSrc.src;
          const rect = imgSrc.getBoundingClientRect();
          el.style.setProperty('width', `${rect.width}px`, 'important');
          el.style.setProperty('height', `${rect.height}px`, 'important');
        }

        // Recurse children
        Array.from(source.children).forEach((child, i) => {
          if (target.children[i]) {
            cloneWithComputedStyles(child, target.children[i]);
          }
        });
      };

      const clone = el.cloneNode(true) as HTMLElement;

      // Apply styles FIRST (while clone structure matches source)
      cloneWithComputedStyles(el, clone);

      // THEN remove UI elements (after styles, so index mismatch doesn't matter)
      clone.querySelectorAll('[data-no-print="true"]').forEach(node => (node as HTMLElement).remove());

      // Get total rendered dimensions
      const rect = el.getBoundingClientRect();

      pw.document.write(`<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>OPS - ${form?.title || 'Operation Standard'}</title>
<style>
  @page {
    size: A3 landscape;
    margin: 3mm;
  }
  * {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
    color-adjust: exact !important;
  }
  html {
    zoom: ${(410 / (rect.width / (96 / 25.4))).toFixed(4)};
  }
  body {
    margin: 0 !important;
    padding: 0 !important;
    font-family: Arial, Helvetica, sans-serif !important;
    background: #fff !important;
  }
  table {
    border-collapse: collapse !important;
    table-layout: fixed !important;
  }
  img {
    -webkit-print-color-adjust: exact !important;
  }
</style>
</head>
<body>
${clone.outerHTML}
</body>
</html>`);

      pw.document.close();

      // Wait for all images to load
      const images = Array.from(pw.document.querySelectorAll('img'));
      let loaded = 0;

      const doPrint = () => {
        pw.focus();
        setTimeout(() => pw.print(), 500);
      };

      if (images.length === 0) {
        setTimeout(doPrint, 400);
      } else {
        let triggered = false;
        const checkDone = () => {
          loaded++;
          if (loaded >= images.length && !triggered) {
            triggered = true;
            doPrint();
          }
        };
        images.forEach(img => {
          if (img.complete) {
            checkDone();
          } else {
            img.onload = checkDone;
            img.onerror = checkDone;
          }
        });
        // Safety fallback
        setTimeout(() => {
          if (!triggered) { triggered = true; doPrint(); }
        }, 4000);
      }
    };

    if (viewMode !== "ops") {
      setViewMode("ops");
      setTimeout(printContent, 700);
    } else {
      printContent();
    }
  }, [viewMode, form]);

  const handleDownloadPDF = async (type?: 'yes-only' | 'no-only' | 'na-only' | 'both' | 'section' | 'default' | 'responses-view') => {
    if (!response || !form) return;

    setShowPDFTypeSelector(false);

    if (type === 'section') {
      setShowSectionsPDFModal(true);
      return;
    }

    await handleDownloadPDFNow(type);
  };

  const handleDownloadSectionsPDF = async () => {
    if (!response || !form) return;

    setDownloadingSectionsPDF(true);
    await new Promise(resolve => setTimeout(resolve, 500));

    try {
      await handleDownloadPDFNow('section');
      setShowSectionsPDFModal(false);
    } finally {
      setDownloadingSectionsPDF(false);
    }
  };

  const handleDownloadPDFNow = async (type?: 'yes-only' | 'no-only' | 'na-only' | 'both' | 'section' | 'default' | 'responses-view') => {
    if (!response || !form) return;

    setPdfDownloadProgress(0);
    setGeneratingPDF(true);

    try {
      const availableSections = form.sections || [];

      // Prepare section question stats
      const questionStats: Record<string, any[]> = {};
      availableSections.forEach((section: any) => {
        questionStats[section.id] = getSectionYesNoQuestionStats(section.id);
      });

      // Create PDF options
      const pdfOptions = {
        filename: `${form.title}_Report_${formatTimestamp(response.createdAt, 'file')}_${type || 'default'}.pdf`,
        formTitle: form.title,
        submittedDate: formatTimestamp(response.createdAt),
        sectionStats: sectionStats,
        sectionQuestionStats: questionStats,
        form: form,
        response: response,
        availableSections: availableSections,
        type: type // Add the type parameter
      };

      // Create progress callback
      const onProgress = (progress: {
        stage: 'uploading' | 'generating' | 'downloading' | 'complete';
        percentage: number;
        message?: string;
      }) => {
        console.log('📊 PDF Progress:', progress);
        setPdfDownloadProgress(Math.round(progress.percentage));
      };

      // Call generateAndDownloadPDF with progress callback ONCE
      await generateAndDownloadPDF(pdfOptions, type, onProgress);

      showSuccess("PDF downloaded successfully.");
    } catch (err: any) {
      console.error("Failed to generate PDF:", err);
      showError(err.message || "Failed to generate PDF. Please try again.");
    } finally {
      setGeneratingPDF(false);
      setPdfDownloadProgress(null);
    }
  };

  const handleBulkDownloadZip = async () => {
    if (!response || !form) return;

    try {
      setExportingZip(true);
      isCancelledRef.current = false;

      const formIdentifier = response.questionId || response.formId;
      if (!formIdentifier) {
        throw new Error("Form identifier not found for this response.");
      }

      // Fetch all responses for this form
      const responsesData = await apiClient.getResponses();
      const filteredResponses = responsesData.responses.filter(
        (r: Response) => {
          const rFormId = r.questionId || r.formId || (r as any).formIdentifier;
          return rFormId === formIdentifier || String(rFormId) === String(formIdentifier);
        }
      );

      if (filteredResponses.length === 0) {
        throw new Error("No responses found for this form.");
      }

      // Get the full form data to ensure we have everything needed
      const formData = await apiClient.getForm(formIdentifier);
      const fullForm = formData.form;

      await exportAllResponsesToZip(
        filteredResponses,
        fullForm,
        (progress) => {
          setPdfProgress({
            stage: 'generating',
            percentage: (progress.current / progress.total) * 100,
            message: progress.message
          });
        },
        () => isCancelledRef.current
      );

      if (isCancelledRef.current) {
        showSuccess("Bulk download cancelled.");
        setPdfProgress(null);
        return;
      }

      showSuccess(`Bulk download of ${filteredResponses.length} responses completed.`);
      setPdfProgress({
        stage: 'complete',
        percentage: 100,
        message: 'Download complete'
      });

      setTimeout(() => setPdfProgress(null), 3000);

    } catch (err: any) {
      console.error("Bulk download failed:", err);
      showError(err.message || "Bulk download failed.");
      setPdfProgress({
        stage: 'error',
        percentage: 0,
        message: err.message
      });
    } finally {
      setExportingZip(false);
    }
  };

  const getStatusInfo = (status: string) => {
    switch (status.toLowerCase()) {
      case "pending":
        return {
          color: "text-yellow-600",
          bgColor: "bg-yellow-50",
          icon: Clock,
          label: "Pending",
        };
      case "confirmed":
        return {
          color: "text-blue-600",
          bgColor: "bg-blue-50",
          icon: CheckCircle,
          label: "Confirmed",
        };
      case "verified":
        return {
          color: "text-green-600",
          bgColor: "bg-green-50",
          icon: CheckCircle,
          label: "Verified",
        };
      case "rejected":
        return {
          color: "text-red-600",
          bgColor: "bg-red-50",
          icon: XCircle,
          label: "Rejected",
        };
      default:
        return {
          color: "text-gray-600",
          bgColor: "bg-gray-50",
          icon: Clock,
          label: "Unknown",
        };
    }
  };

  const getAllQuestions = (form: Form) => {
    const questions: Record<string, any> = {};

    form.sections?.forEach((section) => {
      section.questions?.forEach((question: any) => {
        questions[question.id] = question;
        question.followUpQuestions?.forEach((followUp: any) => {
          questions[followUp.id] = followUp;
        });
      });
    });

    form.followUpQuestions?.forEach((question: any) => {
      questions[question.id] = question;
    });

    return questions;
  };

  function collectYesNoQuestionIds(form: Form): string[] {
    const ids = new Set<string>();

    const processQuestion = (question: any) => {
      if (!question) {
        return;
      }
      const supportedTypes = ["yesNoNA", "radio", "checkbox", "select", "search-select", "radio-image", "rating", "scale", "chassisNumber", "chassis-with-zone", "chassis-without-zone", "zone-in", "zone-out"];
      if (supportedTypes.includes(question.type) && question.id) {
        ids.add(question.id);
      }
      question.followUpQuestions?.forEach(processQuestion);
    };

    form.sections?.forEach((section) => {
      section.questions?.forEach(processQuestion);
    });

    form.followUpQuestions?.forEach(processQuestion);

    return Array.from(ids);
  }

  function extractYesNoValues(value: any): string[] {
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
  }

  const renderHighlightedAnswer = (value: any, question?: any, compact: boolean = false) => {
    const isArray = Array.isArray(value);

    // Helper to get string representation for comparison
    const getStringValue = (val: any): string => {
      if (Array.isArray(val)) return val.map(v => getStringValue(v)).join(", ");
      if (typeof val === 'object' && val !== null) {
        if (val.url) return val.url;
        if (val.answer !== undefined) return String(val.answer);
        // Handle chassis-type objects - don't return string here, let renderInnerValue handle display
        return String(val); // Return placeholder, actual display handled elsewhere
      }
      return String(val || "");
    };

    const strValue = getStringValue(value);
    const normalized = strValue.trim().toLowerCase();

    let bgColor = "bg-white dark:bg-gray-700";
    let textColor = "text-gray-900 dark:text-gray-100";
    let borderColor = "border-gray-200 dark:border-gray-600";
    let Icon = null;

    let isYes = normalized === "yes";
    let isNo = normalized === "no";
    let isNA = normalized === "n/a" || normalized === "na" || normalized === "not applicable";

    // For yesNoNA type, we should use the option position if available
    if (question && question.type === "yesNoNA" && question.options && question.options.length >= 2) {
      isYes = normalized === String(question.options[0]).toLowerCase();
      isNo = normalized === String(question.options[1]).toLowerCase();
      if (question.options.length >= 3) {
        isNA = normalized === String(question.options[2]).toLowerCase();
      }
    }

    // Quiz logic
    const isQuiz = question && (question.correctAnswer || (question.correctAnswers && question.correctAnswers.length > 0));
    let isCorrect = false;

    if (isQuiz) {
      if (question.correctAnswers && question.correctAnswers.length > 0) {
        if (isArray) {
          isCorrect = value.length === question.correctAnswers.length &&
            value.every((a: any) => question.correctAnswers!.some((ca: any) => String(ca).toLowerCase() === getStringValue(a).toLowerCase()));
        } else {
          isCorrect = question.correctAnswers.some((ca: any) => String(ca).toLowerCase() === normalized);
        }
      } else if (question.correctAnswer) {
        isCorrect = String(question.correctAnswer).toLowerCase() === normalized;
      }
    }

    if (isQuiz) {
      if (isCorrect) {
        bgColor = "bg-green-100 dark:bg-green-900/30";
        textColor = "text-green-800 dark:text-green-300";
        borderColor = "border-green-200 dark:border-green-800";
        Icon = CheckCircle;
      } else {
        bgColor = "bg-red-100 dark:bg-red-900/30";
        textColor = "text-red-800 dark:text-red-300";
        borderColor = "border-red-200 dark:border-red-800";
        Icon = XCircle;
      }
    } else if (isYes) {
      bgColor = "bg-green-100 dark:bg-green-900/30";
      textColor = "text-green-800 dark:text-green-300";
      borderColor = "border-green-200 dark:border-green-800";
      Icon = CheckCircle;
    } else if (isNo) {
      bgColor = "bg-red-100 dark:bg-red-900/30";
      textColor = "text-red-800 dark:text-red-300";
      borderColor = "border-red-200 dark:border-red-800";
      Icon = XCircle;
    } else if (isNA) {
      bgColor = "bg-yellow-100 dark:bg-yellow-900/30";
      textColor = "text-yellow-800 dark:text-yellow-300";
      borderColor = "border-yellow-200 dark:border-yellow-800";
      Icon = AlertTriangle;
    }

    const answerBox = (
      <div
        className={`${compact ? 'w-full px-4 py-2' : 'flex-1 p-3'} ${bgColor} ${textColor} ${borderColor} rounded-lg border text-sm break-words font-medium flex items-center shadow-sm`}
        style={{
          boxShadow: isQuiz
            ? (isCorrect ? '0 4px 12px rgba(34, 197, 94, 0.4)' : '0 4px 12px rgba(239, 68, 68, 0.4)')
            : (isYes ? '0 4px 12px rgba(34, 197, 94, 0.4)' : (isNo ? '0 4px 12px rgba(239, 68, 68, 0.4)' : '0 1px 2px 0 rgba(0, 0, 0, 0.05)'))
        }}
      >
        <div className="w-full">
          {!compact && isQuiz && (
            <div className="text-[10px] font-bold opacity-70 uppercase tracking-wider mb-1">
              Customer Filled Answer
            </div>
          )}
          <div className={`flex items-center gap-2 ${compact ? 'justify-center' : ''}`}>
            {Icon && <Icon className="w-4 h-4 flex-shrink-0" />}
            <div className="flex-1">
              {(() => {
                const renderInnerValue = (val: any): React.ReactNode => {
                  if (val === null || val === undefined || val === "") return null;

                  if (Array.isArray(val)) {
                    return (
                      <div className="flex flex-col gap-1">
                        {val.map((item, i) => (
                          <div key={i}>{renderInnerValue(item)}</div>
                        ))}
                      </div>
                    );
                  }

                  if (typeof val === "object" && val !== null) {
                    if (val.url && isImageUrl(String(val.url))) {
                      return <ImageLink text={String(val.url)} />;
                    }
                    if (val.answer && isImageUrl(String(val.answer))) {
                      return <ImageLink text={String(val.answer)} />;
                    }

                    const entries = Object.entries(val);

                    // Check if this is a chassis-type object
                    const isChassisType = val.chassisNumber !== undefined || val.status !== undefined || val.zone !== undefined || val.categories !== undefined;

                    if (isChassisType) {
                      // Get color for zone
                      const getZoneColor = (zoneName: string): string => {
                        const z = zoneName.toLowerCase().trim();
                        if (z.includes('zone a') || z === 'a') return 'blue';
                        if (z.includes('zone b') || z === 'b') return 'green';
                        if (z.includes('zone c') || z === 'c') return 'purple';
                        if (z.includes('zone d') || z === 'd') return 'orange';
                        if (z.includes('zone e') || z === 'e') return 'pink';
                        if (z.includes('zone f') || z === 'f') return 'cyan';
                        return 'indigo';
                      };

                      const colorMap: Record<string, string> = {
                        blue: 'bg-blue-50 dark:bg-blue-900/30 text-blue-800 dark:text-blue-200',
                        green: 'bg-green-50 dark:bg-green-900/30 text-green-800 dark:text-green-200',
                        purple: 'bg-purple-50 dark:bg-purple-900/30 text-purple-800 dark:text-purple-200',
                        orange: 'bg-orange-50 dark:bg-orange-900/30 text-orange-800 dark:text-orange-200',
                        pink: 'bg-pink-50 dark:bg-pink-900/30 text-pink-800 dark:text-pink-200',
                        cyan: 'bg-cyan-50 dark:bg-cyan-900/30 text-cyan-800 dark:text-cyan-200',
                        red: 'bg-red-50 dark:bg-red-900/30 text-red-800 dark:text-red-200',
                        amber: 'bg-amber-50 dark:bg-amber-900/30 text-amber-800 dark:text-amber-200',
                        indigo: 'bg-indigo-50 dark:bg-indigo-900/30 text-indigo-800 dark:text-indigo-200',
                      };

                      const parts: { label: string; value: any; zoneColor?: string; isImage?: boolean }[] = [];

                      if (val.chassisNumber && String(val.chassisNumber).trim() && String(val.chassisNumber).toLowerCase() !== 'no response') {
                        parts.push({ label: 'Chassis', value: String(val.chassisNumber), zoneColor: 'blue' });
                      }
                      if (val.status && String(val.status).trim() && String(val.status).toLowerCase() !== 'no response') {
                        parts.push({ label: 'Status', value: String(val.status), zoneColor: 'red' });
                      }
                      if (val.zone) {
                        const zoneVal = Array.isArray(val.zone) ? val.zone.join(', ') : String(val.zone);
                        if (zoneVal.trim()) {
                          // If multiple zones, show each with its own color and sort alphabetically
                          if (zoneVal.includes(',')) {
                            const zones = zoneVal.split(',').map(z => z.trim()).sort((a, b) => a.localeCompare(b));
                            zones.forEach(z => {
                              parts.push({ label: 'Zone', value: z, zoneColor: getZoneColor(z) });
                            });
                          } else {
                            parts.push({ label: 'Zone', value: zoneVal, zoneColor: getZoneColor(zoneVal) });
                          }
                        }
                      }

                      // Handle zonesData with colors
                      if (val.zonesData && typeof val.zonesData === 'object') {
                        const zoneEntries = Object.entries(val.zonesData);
                        for (const [zoneName, zoneVal] of zoneEntries) {
                          const zoneColor = getZoneColor(zoneName);
                          const colorClass = colorMap[zoneColor] || colorMap.indigo;

                          parts.push({ label: 'Zone', value: zoneName, zoneColor });

                          const categories = (zoneVal as any)?.categories;
                          if (categories && Array.isArray(categories)) {
                            for (const cat of categories) {
                              const catName = typeof cat === 'string' ? cat : (cat?.name || cat?.category || '-');
                              parts.push({ label: 'Category', value: String(catName), zoneColor });

                              const defects = cat?.defects;
                              if (defects && Array.isArray(defects)) {
                                for (const defect of defects) {
                                  const defectName = typeof defect === 'string' ? defect : (defect?.name || defect?.defect || '-');
                                  const defectDetails = typeof defect === 'object' ? (defect?.details || {}) : {};
                                  const remark = defectDetails?.remark || defectDetails?.remarks || '-';
                                  parts.push({ label: 'Defect', value: String(defectName), zoneColor });
                                  if (remark && String(remark).trim() && String(remark).toLowerCase() !== '-') {
                                    parts.push({ label: 'Remark', value: String(remark), zoneColor });
                                  }
                                  const fileUrl = defectDetails?.fileUrl || defectDetails?.file || defect?.fileUrl || defect?.file || defect?.imageUrl || '';
                                  if (fileUrl && String(fileUrl).toLowerCase() !== 'no response' && String(fileUrl).trim()) {
                                    parts.push({ label: 'Evidence', value: String(fileUrl), zoneColor, isImage: true });
                                  }
                                }
                              }
                            }
                          }
                        }
                      }

                      // Handle categories (direct property) - both object and array formats
                      if (val.categories) {
                        if (Array.isArray(val.categories)) {
                          // ChassisWithoutZone format: array of category objects
                          for (const cat of val.categories) {
                            const catName = cat?.name || cat?.category || '-';
                            if (catName !== '-') {
                              parts.push({ label: 'Category', value: String(catName), zoneColor: 'purple' });

                              const defects = cat?.defects;
                              if (defects && Array.isArray(defects)) {
                                for (const defect of defects) {
                                  const defectName = typeof defect === 'string' ? defect : (defect?.name || defect?.defect || '-');
                                  const defectDetails = typeof defect === 'object' ? (defect?.details || {}) : {};
                                  const remark = defectDetails?.remark || defectDetails?.remarks || '-';
                                  parts.push({ label: 'Defect', value: String(defectName), zoneColor: 'purple' });
                                  if (remark && String(remark).trim() && String(remark).toLowerCase() !== '-') {
                                    parts.push({ label: 'Remark', value: String(remark), zoneColor: 'purple' });
                                  }
                                  const fileUrl = defectDetails?.fileUrl || defectDetails?.file || defect?.fileUrl || defect?.file || defect?.imageUrl || '';
                                  if (fileUrl && String(fileUrl).toLowerCase() !== 'no response' && String(fileUrl).trim()) {
                                    parts.push({ label: 'Evidence', value: String(fileUrl), zoneColor: 'purple', isImage: true });
                                  }
                                }
                              }
                            }
                          }
                        } else if (typeof val.categories === 'object') {
                          // Object format: key-value pairs
                          const catEntries = Object.entries(val.categories);
                          for (const [catKey, catVal] of catEntries) {
                            parts.push({ label: String(catKey), value: String(catVal), zoneColor: 'amber' });
                          }
                        }
                      }

                      // Handle evidenceUrl
                      if (val.evidenceUrl && String(val.evidenceUrl).toLowerCase() !== 'no response' && String(val.evidenceUrl).trim()) {
                        parts.push({ label: 'Evidence', value: String(val.evidenceUrl), zoneColor: 'indigo', isImage: true });
                      }

                      if (parts.length > 0) {
                        return (
                          <div className="flex flex-col gap-1">
                            {parts.map((part, idx) => {
                              const pc = colorMap[part.zoneColor || 'indigo'] || colorMap.indigo;
                              return (
                                <div key={idx} className="flex items-start gap-2">
                                  <span className={`px-2 py-0.5 ${pc} text-xs rounded font-medium min-w-[70px]`}>
                                    {part.label}
                                  </span>
                                  {part.isImage ? (
                                    <ImageLink text={part.value} showImage={true} />
                                  ) : (
                                    <span className={`px-2 py-0.5 ${pc} text-xs rounded font-medium`}>
                                      {String(part.value)}
                                    </span>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        );
                      }

                      return <span className="text-gray-400 italic">No response</span>;
                    }

                    if (entries.length > 0) {
                      // Filter out empty values and "No response" strings
                      const filteredEntries = entries.filter(([k, v]) => {
                        if (v === null || v === undefined) return false;
                        if (typeof v === 'string') {
                          const lowerV = v.toLowerCase().trim();
                          return lowerV !== '' && lowerV !== 'no response' && lowerV !== 'n/a' && lowerV !== 'na';
                        }
                        if (typeof v === 'object') {
                          if (Array.isArray(v)) return v.length > 0;
                          const objEntries = Object.entries(v).filter(([ok, ov]) => {
                            if (ov === null || ov === undefined) return false;
                            if (typeof ov === 'string') return ov.toLowerCase().trim() !== '' && ov.toLowerCase().trim() !== 'no response';
                            return true;
                          });
                          return objEntries.length > 0;
                        }
                        return true;
                      });

                      if (filteredEntries.length === 0) {
                        return <span className="text-gray-400 italic">No response</span>;
                      }

                      return (
                        <div className="flex flex-col gap-1">
                          {filteredEntries.map(([k, v], i) => {
                            const displayKey = k.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase()).trim();
                            return (
                              <div key={i} className="flex flex-col gap-0.5">
                                <span className="text-[10px] font-bold opacity-70 uppercase tracking-tighter text-indigo-800 dark:text-indigo-300">
                                  {displayKey}
                                </span>
                                {renderInnerValue(v)}
                              </div>
                            );
                          })}
                        </div>
                      );
                    }
                    return JSON.stringify(val);
                  }

                  const textVal = String(val);
                  if (isImageUrl(textVal)) {
                    return <ImageLink text={textVal} />;
                  }
                  return textVal;
                };
                return renderInnerValue(value);
              })()}
            </div>
          </div>
        </div>
      </div>
    );

    if (isQuiz && !compact) {
      const correctAnswerDisplay = question.correctAnswers && question.correctAnswers.length > 0
        ? question.correctAnswers.join(", ")
        : String(question.correctAnswer || "");

      return (
        <div className="flex flex-row gap-3 w-full">
          {/* Given Correct Answer - No color (Neutral) */}
          <div className="flex-1 bg-gray-50 dark:bg-gray-800/50 border border-gray-200 dark:border-gray-700 rounded-lg p-3 text-sm">
            <div className="text-[10px] font-bold text-gray-500 dark:text-gray-400 uppercase tracking-wider mb-1">
              Expected Answer
            </div>
            <div className="font-medium text-gray-700 dark:text-gray-300">
              {correctAnswerDisplay}
            </div>
          </div>

          {/* Customer Filled Answer - Colored based on correctness */}
          {answerBox}
        </div>
      );
    }

    return answerBox;
  };

  function getSectionYesNoStats(
    form: Form,
    answers: Record<string, any>
  ): SectionStat[] {
    const stats =
      form.sections?.map((section: any) => {
        const counts = { yes: 0, no: 0, na: 0, total: 0, correct: 0, wrong: 0, answeredCount: 0 };
        let hasYesNo = false;

        const processQuestion = (question: any) => {
          if (!question) {
            return;
          }
          const supportedTypes = ["yesNoNA", "radio", "checkbox", "select", "search-select", "radio-image", "rating", "scale", "chassisNumber", "chassis-with-zone", "chassis-without-zone", "zone-in", "zone-out"];
          if (!supportedTypes.includes(question.type) || !question.id) {
            question.followUpQuestions?.forEach(processQuestion);
            return;
          }

          const isCompliance = question.type === "yesNoNA" || ["chassisNumber", "chassis-with-zone", "chassis-without-zone", "zone-in", "zone-out"].includes(question.type);
          const isAccuracy = !isCompliance;

          if (isCompliance) {
            hasYesNo = true;
          }

          const rawValue = answers?.[question.id];
          const normalizedValues = extractYesNoValues(rawValue);

          // Check if it has any value (not null/undefined/empty string/empty array)
          const hasValue = rawValue !== null && rawValue !== undefined && rawValue !== "" &&
            (!Array.isArray(rawValue) || rawValue.length > 0) &&
            (typeof rawValue !== 'object' || Object.keys(rawValue).length > 0);

          counts.total += 1;

          if (!hasValue) {
            if (question.required) {
              counts.answeredCount += 1;
              if (isCompliance) {
                counts.no += 1;
              } else {
                counts.wrong += 1;
              }
            }
            // For unanswered questions, we still check follow-ups 
            question.followUpQuestions?.forEach(processQuestion);
            return;
          }

          counts.answeredCount += 1;

          if (isAccuracy) {
            const isNA = normalizedValues.some(v => ["n/a", "na", "not applicable"].includes(v));
            if (isNA) {
              counts.na += 1;
            } else {
              let isCorrect = false;
              const isArray = Array.isArray(rawValue);
              const strValue = isArray ? rawValue.join(", ") : String(rawValue || "");
              const normalized = strValue.trim().toLowerCase();

              if (question.correctAnswers && question.correctAnswers.length > 0) {
                if (isArray) {
                  isCorrect = rawValue.length === question.correctAnswers.length &&
                    rawValue.every((a: any) => question.correctAnswers!.some((ca: any) => String(ca).toLowerCase() === String(a).toLowerCase()));
                } else {
                  isCorrect = question.correctAnswers.some((ca: any) => String(ca).toLowerCase() === normalized);
                }
              } else if (question.correctAnswer) {
                isCorrect = String(question.correctAnswer).toLowerCase() === normalized;
              } else {
                // Fallback for accuracy questions without explicit correct answers:
                // If it has a value and it's not "N/A", it's considered "Correct" (Answered)
                isCorrect = true;

                // Special logic for chassis/zone types: if rejected or has defects, it's "Wrong"
                if (["chassis-with-zone", "zone-in", "zone-out", "chassis-without-zone"].includes(question.type)) {
                  if (rawValue && typeof rawValue === 'object') {
                    const hasDefects = rawValue.status === 'Rejected' ||
                      (rawValue.zonesData && Object.keys(rawValue.zonesData).length > 0) ||
                      (rawValue.categories && Object.keys(rawValue.categories).length > 0);
                    if (hasDefects) {
                      isCorrect = false;
                    }
                  }
                }
              }

              if (isCorrect) {
                counts.correct += 1;
              } else {
                counts.wrong += 1;
              }
            }
          } else if (isCompliance) {
            // Special logic for chassis/zone types within compliance
            if (["chassis-with-zone", "zone-in", "zone-out", "chassis-without-zone"].includes(question.type)) {
              if (rawValue && typeof rawValue === 'object') {
                const hasDefects = rawValue.status === 'Rejected' ||
                  (rawValue.zonesData && Object.keys(rawValue.zonesData).length > 0) ||
                  (rawValue.categories && Object.keys(rawValue.categories).length > 0);
                if (hasDefects) {
                  counts.no = 1;
                } else {
                  counts.yes = 1;
                }
              } else {
                counts.yes = 1;
              }
            } else {
              const yesLabel = question.options?.[0]?.toLowerCase() || "yes";
              const noLabel = question.options?.[1]?.toLowerCase() || "no";
              const naLabel = question.options?.[2]?.toLowerCase() || "n/a";

              if (normalizedValues.includes(yesLabel)) {
                counts.yes = 1;
              } else if (normalizedValues.includes(noLabel)) {
                counts.no = 1;
              } else if (
                normalizedValues.includes(naLabel) ||
                normalizedValues.includes("n/a") ||
                normalizedValues.includes("na") ||
                normalizedValues.includes("not applicable")
              ) {
                counts.na = 1;
              } else {
                counts.yes = 1;
              }
            }
          }

          question.followUpQuestions?.forEach(processQuestion);
        };

        section.questions?.forEach(processQuestion);

        if (!counts.answeredCount) {
          return null;
        }

        return {
          id: section.id,
          title: section.title || "Untitled Section",
          yes: counts.yes,
          no: counts.no,
          na: counts.na,
          correct: counts.correct,
          wrong: counts.wrong,
          total: counts.total,
          answeredCount: counts.answeredCount,
          hasYesNo,
        };
      }) ?? [];

    return stats.filter((stat): stat is SectionStat => Boolean(stat));
  }

  function getSectionYesNoQuestionStats(sectionId: string) {
    if (!form || !response) return [];

    const section = form.sections?.find((s: any) => s.id === sectionId);
    if (!section) return [];

    const questionStats: Array<{
      id: string;
      title: string;
      subParam1?: string;
      yes: number;
      no: number;
      na: number;
      correct: number;
      wrong: number;
      total: number;
      hasYesNo: boolean;
      isQuiz: boolean;
    }> = [];

    const processQuestion = (question: any) => {
      if (!question) return;

      const supportedTypes = ["yesNoNA", "radio", "checkbox", "select", "search-select", "radio-image", "rating", "scale", "chassisNumber", "chassis-with-zone", "chassis-without-zone", "zone-in", "zone-out"];
      if (supportedTypes.includes(question.type) && question.id) {
        const rawValue = response.answers?.[question.id];
        const normalizedValues = extractYesNoValues(rawValue);
        const counts = { yes: 0, no: 0, na: 0, total: 1, correct: 0, wrong: 0, answeredCount: 0 };
        const isCompliance = question.type === "yesNoNA" || ["chassisNumber", "chassis-with-zone", "chassis-without-zone", "zone-in", "zone-out"].includes(question.type);
        const isAccuracy = !isCompliance;

        const hasValue = rawValue !== null && rawValue !== undefined && rawValue !== "" &&
          (!Array.isArray(rawValue) || rawValue.length > 0) &&
          (typeof rawValue !== 'object' || Object.keys(rawValue).length > 0);

        if (!hasValue) {
          if (question.required) {
            counts.answeredCount = 1;
            if (isCompliance) {
              counts.no = 1;
            } else {
              counts.wrong = 1;
            }
          } else {
            // Optional unanswered questions are skipped from stats
            return;
          }
        } else {
          counts.answeredCount = 1;
          if (isAccuracy) {
            const isNA = normalizedValues.some(v => ["n/a", "na", "not applicable"].includes(v));
            if (isNA) {
              counts.na = 1;
            } else {
              let isCorrect = false;
              const isArray = Array.isArray(rawValue);
              const strValue = isArray ? rawValue.join(", ") : String(rawValue || "");
              const normalized = strValue.trim().toLowerCase();

              if (question.correctAnswers && question.correctAnswers.length > 0) {
                if (isArray) {
                  isCorrect = rawValue.length === question.correctAnswers.length &&
                    rawValue.every((a: any) => question.correctAnswers!.some((ca: any) => String(ca).toLowerCase() === String(a).toLowerCase()));
                } else {
                  isCorrect = question.correctAnswers.some((ca: any) => String(ca).toLowerCase() === normalized);
                }
              } else if (question.correctAnswer) {
                isCorrect = String(question.correctAnswer).toLowerCase() === normalized;
              } else {
                // Fallback for accuracy questions without explicit correct answers:
                // If it has a value and it's not "N/A", it's considered "Correct" (Answered)
                isCorrect = true;

                // Special logic for chassis/zone types: if rejected or has defects, it's "Wrong"
                if (["chassis-with-zone", "zone-in", "zone-out", "chassis-without-zone"].includes(question.type)) {
                  if (rawValue && typeof rawValue === 'object') {
                    const hasDefects = rawValue.status === 'Rejected' ||
                      (rawValue.zonesData && Object.keys(rawValue.zonesData).length > 0) ||
                      (rawValue.categories && Object.keys(rawValue.categories).length > 0);
                    if (hasDefects) {
                      isCorrect = false;
                    }
                  }
                }
              }

              if (isCorrect) {
                counts.correct = 1;
              } else {
                counts.wrong = 1;
              }
            }
          } else if (isCompliance) {
            const yesLabel = question.options?.[0]?.toLowerCase() || "yes";
            const noLabel = question.options?.[1]?.toLowerCase() || "no";
            const naLabel = question.options?.[2]?.toLowerCase() || "n/a";

            if (normalizedValues.includes(yesLabel)) {
              counts.yes = 1;
            } else if (normalizedValues.includes(noLabel)) {
              counts.no = 1;
            } else if (
              normalizedValues.includes(naLabel) ||
              normalizedValues.includes("n/a") ||
              normalizedValues.includes("na") ||
              normalizedValues.includes("not applicable")
            ) {
              counts.na = 1;
            } else {
              counts.yes = 1;
            }
          }
        }

        questionStats.push({
          id: question.id,
          title:
            question.title ||
            question.label ||
            question.text ||
            `Question ${question.id}`,
          subParam1: question.subParam1,
          hasYesNo: isCompliance,
          isQuiz: isAccuracy,
          ...counts,
        });

        // For zone types, unroll categories into individual stats if defects exist
        // This allows the "category" to show up in the bar charts as requested
        if (["chassis-with-zone", "zone-in", "zone-out"].includes(question.type) && rawValue && typeof rawValue === "object") {
          const zonesData = rawValue.zonesData || {};
          Object.entries(zonesData).forEach(([zoneName, zoneVal]: [string, any]) => {
            const categories = zoneVal?.categories;
            if (Array.isArray(categories)) {
              categories.forEach((cat: any) => {
                const catName = cat?.name || cat?.category || "-";
                questionStats.push({
                  id: `${question.id}-${zoneName}-${catName}`,
                  title: `${question.title} - ${zoneName} - ${catName}`,
                  subParam1: catName,
                  hasYesNo: true,
                  isQuiz: false,
                  yes: 0,
                  no: 1,
                  na: 0,
                  total: 1,
                  correct: 0,
                  wrong: 0,
                  answeredCount: 1,
                });
              });
            }
          });

          // Also check for flat categories structure
          const flatCategories = rawValue.categories;
          if (Array.isArray(flatCategories)) {
            flatCategories.forEach((cat: any) => {
              const catName = cat?.name || cat?.category || "-";
              questionStats.push({
                id: `${question.id}-${catName}`,
                title: `${question.title} - ${catName}`,
                subParam1: catName,
                hasYesNo: true,
                isQuiz: false,
                yes: 0,
                no: 1,
                na: 0,
                total: 1,
                correct: 0,
                wrong: 0,
                answeredCount: 1,
              });
            });
          }
        }
      }

      question.followUpQuestions?.forEach(processQuestion);
    };

    section.questions?.forEach(processQuestion);

    const groupedStats: Map<
      string,
      {
        id: string;
        title: string;
        subParam1?: string;
        yes: number;
        no: number;
        na: number;
        correct: number;
        wrong: number;
        total: number;
        hasYesNo: boolean;
        isQuiz: boolean;
      }
    > = new Map();

    questionStats.forEach((stat) => {
      const key = stat.subParam1 || "No parameter";
      if (groupedStats.has(key)) {
        const existing = groupedStats.get(key)!;
        existing.yes += stat.yes;
        existing.no += stat.no;
        existing.na += stat.na;
        existing.correct += stat.correct;
        existing.wrong += stat.wrong;
        existing.total += stat.total;
        if (stat.hasYesNo) existing.hasYesNo = true;
        if (stat.isQuiz) existing.isQuiz = true;
      } else {
        groupedStats.set(key, { ...stat });
      }
    });

    return Array.from(groupedStats.values());
  }

  const getSectionQuestionsWithFollowUps = (sectionId: string) => {
    if (!form || !response) return [];

    const section = form.sections?.find((s: any) => s.id === sectionId);
    if (!section) return [];

    const mainQuestionsWithFollowUps: any[] = [];
    const questionIds = collectYesNoQuestionIds({
      ...form,
      sections: [section],
    });

    const mainQuestions: any[] = [];
    const followUpMap = new Map<string, any[]>();

    section.questions?.forEach((question: any) => {
      if (question.showWhen && question.showWhen.questionId) {
        const parentId = question.showWhen.questionId;
        if (!followUpMap.has(parentId)) {
          followUpMap.set(parentId, []);
        }
        followUpMap.get(parentId)!.push(question);
      } else {
        mainQuestions.push(question);
      }
    });

    mainQuestions.forEach((question: any) => {
      // INCLUDE ALL MAIN QUESTIONS, not just those from collectYesNoQuestionIds
      // This ensures text-based main parameters (like in "Basic Information") are shown
      const answers = response.answers?.[question.id];
      const yesNoValues = extractYesNoValues(answers);

      const followUpQuestionsForThis = [
        ...(form.followUpQuestions?.filter(
          (fq: any) => fq.parentId === question.id
        ) || []),
        ...(question.followUpQuestions || []),
        ...(followUpMap.get(question.id) || []),
      ];

      // Show question if it has an answer OR has follow-ups
      if ((answers !== undefined && answers !== null && answers !== "") || followUpQuestionsForThis.length > 0) {
        const mainQuestion = {
          id: question.id,
          title: question.title || question.label || question.text,
          subParam1: question.subParam1,
          yesNoValues,
          followUpQuestions: followUpQuestionsForThis.map((fq: any) => ({
            id: fq.id || fq._id,
            title: fq.title || fq.label || fq.text,
            subParam1: fq.subParam1,
            answer: response.answers?.[fq.id || fq._id],
          })),
        };

        mainQuestionsWithFollowUps.push(mainQuestion);
      }
    });

    return mainQuestionsWithFollowUps;
  };

  const sectionStats = useMemo(() => {
    if (!form || !response) {
      return [] as SectionStat[];
    }
    return getSectionYesNoStats(form, response.answers);
  }, [form, response]);

  const filteredSectionStats = useMemo(
    () =>
      sectionStats.filter(
        (stat) =>
          stat.total > 0
      ),
    [sectionStats]
  );

  const sectionChartData = useMemo(() => {
    const calculatePercentage = (value: number, total: number) =>
      total ? parseFloat(((value / total) * 100).toFixed(1)) : 0;

    const datasets = [
      {
        label: complianceLabels.correct,
        data: filteredSectionStats.map((stat) =>
          calculatePercentage(stat.correct, stat.total)
        ),
        borderColor: "#10b981",
        backgroundColor: "rgba(16, 185, 129, 0.25)",
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: "#10b981",
        pointBorderColor: "#fff",
        pointBorderWidth: 2,
        pointHoverRadius: 6,
        tension: 0.4,
      },
      {
        label: complianceLabels.wrong,
        data: filteredSectionStats.map((stat) =>
          calculatePercentage(stat.wrong, stat.total)
        ),
        borderColor: "#ef4444",
        backgroundColor: "rgba(239, 68, 68, 0.25)",
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: "#ef4444",
        pointBorderColor: "#fff",
        pointBorderWidth: 2,
        pointHoverRadius: 6,
        tension: 0.4,
      },
      {
        label: complianceLabels.yes,
        data: filteredSectionStats.map((stat) =>
          calculatePercentage(stat.yes, stat.total)
        ),
        borderColor: "#1d4ed8",
        backgroundColor: "rgba(29, 78, 216, 0.25)",
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: "#1d4ed8",
        pointBorderColor: "#fff",
        pointBorderWidth: 2,
        pointHoverRadius: 6,
        tension: 0.4,
      },
      {
        label: complianceLabels.no,
        data: filteredSectionStats.map((stat) =>
          calculatePercentage(stat.no, stat.total)
        ),
        borderColor: "#3b82f6",
        backgroundColor: "rgba(59, 130, 246, 0.25)",
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: "#3b82f6",
        pointBorderColor: "#fff",
        pointBorderWidth: 2,
        pointHoverRadius: 6,
        tension: 0.4,
      },
    ];

    if (filteredSectionStats.some(stat => stat.na > 0)) {
      datasets.push({
        label: complianceLabels.na,
        data: filteredSectionStats.map((stat) =>
          calculatePercentage(stat.na, stat.total)
        ),
        borderColor: "#93c5fd",
        backgroundColor: "rgba(147, 197, 253, 0.25)",
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: "#93c5fd",
        pointBorderColor: "#fff",
        pointBorderWidth: 2,
        pointHoverRadius: 6,
        tension: 0.4,
      });
    }

    return {
      labels: filteredSectionStats.map((stat) =>
        formatSectionLabel(stat.title)
      ),
      datasets,
    };
  }, [filteredSectionStats]);

  const sectionChartOptions = useMemo(
    () => ({
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          callbacks: {
            title: (items: any[]) => {
              const index = items?.[0]?.dataIndex;
              if (index === undefined) {
                return "";
              }
              return filteredSectionStats[index]?.title || "";
            },
            label: (context: any) => {
              const value = context.parsed?.r ?? 0;
              return `${context.dataset.label}: ${value.toFixed(1)}%`;
            },
          },
        },
      },
      scales: {
        r: {
          beginAtZero: true,
          max: 100,
          pointLabels: {
            display: true,
            font: {
              size: 10,
              weight: 500,
            },
            color: document.documentElement.classList.contains("dark")
              ? "#d1d5db"
              : "#374151",
          },
          ticks: {
            callback: (value: any) => `${value}%`,
            color: document.documentElement.classList.contains("dark")
              ? "#d1d5db"
              : "#374151",
            font: {
              size: 11,
            },
          },
          grid: {
            color: document.documentElement.classList.contains("dark")
              ? "rgba(147, 197, 253, 0.3)"
              : "rgba(59, 130, 246, 0.3)",
            lineWidth: 1.5,
          },
          angleLines: {
            display: true,
            color: document.documentElement.classList.contains("dark")
              ? "rgba(147, 197, 253, 0.4)"
              : "rgba(59, 130, 246, 0.4)",
            lineWidth: 1.5,
          },
        },
      },
    }),
    [filteredSectionStats]
  );

  const sectionChartHeight = 450;

  const sectionSummaryRows = useMemo(
    () =>
      filteredSectionStats.map((stat) => {
        const yesPercent = stat.total ? (stat.yes / stat.total) * 100 : 0;
        const noPercent = stat.total ? (stat.no / stat.total) * 100 : 0;
        const naPercent = stat.total ? (stat.na / stat.total) * 100 : 0;
        const correctPercent = stat.total ? (stat.correct / stat.total) * 100 : 0;
        const wrongPercent = stat.total ? (stat.wrong / stat.total) * 100 : 0;

        return {
          id: stat.id,
          title: stat.title,
          total: stat.total,
          yes: stat.yes,
          no: stat.no,
          na: stat.na,
          correct: stat.correct,
          wrong: stat.wrong,
          yesPercent,
          noPercent,
          naPercent,
          correctPercent,
          wrongPercent,
          hasYesNo: stat.hasYesNo,
          hasQuiz: stat.correct > 0 || (stat.total > 0 && !stat.hasYesNo), // If total > 0 and no yes/no, it might be quiz
        };
      }),
    [filteredSectionStats]
  );

  const summaryTotals = useMemo(() => {
    return sectionSummaryRows.reduce(
      (acc, row) => ({
        total: acc.total + row.total,
        yes: acc.yes + (row.yes || 0),
        no: acc.no + (row.no || 0),
        na: acc.na + (row.na || 0),
        correct: acc.correct + (row.correct || 0),
        wrong: acc.wrong + (row.wrong || 0),
        hasAnyYesNo: acc.hasAnyYesNo || row.hasYesNo,
        hasAnyQuiz: acc.hasAnyQuiz || row.hasQuiz || row.correct > 0,
      }),
      { total: 0, yes: 0, no: 0, na: 0, correct: 0, wrong: 0, hasAnyYesNo: false, hasAnyQuiz: false }
    );
  }, [sectionSummaryRows]);

  const renderFormContent = () => {
    if (!response || !form) return null;

    const questions = getAllQuestions(form);
    return (
      <div className="space-y-4">
        {Object.entries(response.answers).map(([key, value]) => {
          const question = questions[key];
          return (
            <div
              key={key}
              className="border-b border-primary-100 dark:border-gray-700 pb-4 last:border-b-0"
            >
              <div className="font-semibold text-primary-900 dark:text-gray-100 mb-2">
                {question?.text || key}
              </div>
              {question?.description && (
                <p className="text-sm text-primary-600 dark:text-gray-400 mb-2">
                  {question.description}
                </p>
              )}
              <div className="bg-primary-50 dark:bg-gray-800 rounded-lg p-4">
                {Array.isArray(value) ? (
                  <div className="space-y-1">
                    {value.map((v, idx) => (
                      <div
                        key={idx}
                        className="text-primary-700 dark:text-gray-200"
                      >
                        {isImageUrl(typeof v === 'object' && v !== null && v.url ? String(v.url) : String(v)) ? (
                          <ImageLink text={typeof v === 'object' && v !== null && v.url ? String(v.url) : String(v)} />
                        ) : (
                          String(v)
                        )}
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="text-primary-700 dark:text-gray-200">
                    {isImageUrl(typeof value === 'object' && value !== null && value.url ? String(value.url) : String(value)) ? (
                      <ImageLink text={typeof value === 'object' && value !== null && value.url ? String(value.url) : String(value)} />
                    ) : typeof value === "object" && value !== null ? (
                      <div className="flex flex-col gap-2">
                        {Object.entries(value).map(([k, v], i) => (
                          <div key={i} className="flex flex-col gap-0.5 border-l-2 border-primary-100 pl-2">
                            <span className="text-[10px] font-bold opacity-70 uppercase tracking-tighter text-primary-600 dark:text-primary-400">
                              {k}
                            </span>
                            {isImageUrl(typeof v === 'object' && v !== null && (v as any).url ? String((v as any).url) : String(v)) ? (
                              <ImageLink text={typeof v === 'object' && v !== null && (v as any).url ? String((v as any).url) : String(v)} />
                            ) : (
                              String(v)
                            )}
                          </div>
                        ))}
                      </div>
                    ) : (
                      String(value)
                    )}
                  </div>
                )}
              </div>

              {/* ADDITION: Render synthetic follow-ups for this question */}
              {(() => {
                const syntheticKey = `synthetic_${key}`;
                const syntheticData = response.answers[syntheticKey];
                if (syntheticData && typeof syntheticData === 'object') {
                  return Object.entries(syntheticData).map(([fuText, fuData]: [string, any], idx) => (
                    <div key={idx} className="mt-3 ml-6 border-l-2 border-red-200 pl-4 py-2 bg-red-50/30 dark:bg-red-900/10 rounded-r-lg">
                      <div className="text-xs font-bold text-red-600 dark:text-red-400 mb-1 flex items-center gap-1">
                        <span className="bg-red-100 dark:bg-red-900/30 px-1.5 py-0.5 rounded text-[10px]">FU.S</span>
                        {fuText}
                      </div>
                      <div className="text-sm text-gray-700 dark:text-gray-300">
                        {isImageUrl(fuData.answer && typeof fuData.answer === 'object' && fuData.answer.url ? String(fuData.answer.url) : String(fuData.answer)) ? (
                          <ImageLink text={fuData.answer && typeof fuData.answer === 'object' && fuData.answer.url ? String(fuData.answer.url) : String(fuData.answer)} />
                        ) : (
                          String(fuData.answer)
                        )}
                      </div>
                    </div>
                  ));
                }
                return null;
              })()}
            </div>
          );
        })}
      </div>
    );
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary-600"></div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <div className="text-red-600 mb-4 text-lg">{error}</div>
          <button
            onClick={() => navigate("/responses/all")}
            className="btn-secondary flex items-center gap-2 mx-auto"
          >
            <ArrowLeft className="w-4 h-4" />
            Back to Responses
          </button>
        </div>
      </div>
    );
  }

  if (!response || !form) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <div className="text-gray-600 mb-4">Response not found</div>
          <button
            onClick={() => navigate("/responses/all")}
            className="btn-secondary flex items-center gap-2 mx-auto"
          >
            <ArrowLeft className="w-4 h-4" />
            Back to Responses
          </button>
        </div>
      </div>
    );
  }

  const statusInfo = getStatusInfo(response.status || "pending");
  const StatusIcon = statusInfo.icon;
  const questions = getAllQuestions(form);

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100 dark:from-gray-900 dark:to-gray-800">
      {/* Main Content */}
      <div className="px-6 md:px-8 py-6">

        {/* View Mode Tabs & Back Button */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 mb-6">
          <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4">
            <div className="flex gap-1 bg-white dark:bg-gray-700 rounded-lg p-1 w-fit border border-gray-200 dark:border-gray-600">
              <button
                onClick={() => setViewMode("dashboard")}
                className={`flex items-center gap-2 px-3.5 py-1.5 text-sm font-medium rounded-md transition-all duration-200 ${viewMode === "dashboard"
                  ? "text-white"
                  : "text-gray-900 dark:text-gray-100 hover:text-black dark:hover:text-white"
                  }`}
                style={{ backgroundColor: viewMode === "dashboard" ? "#1e3a8a" : "transparent" }}
              >
                <BarChart3 className="w-4 h-4" />
                Dashboard
              </button>
              <button
                onClick={() => setViewMode("responses")}
                className={`flex items-center gap-2 px-3.5 py-1.5 text-sm font-medium rounded-md transition-all duration-200 ${viewMode === "responses"
                  ? "text-white"
                  : "text-gray-900 dark:text-gray-100 hover:text-black dark:hover:text-white"
                  }`}
                style={{ backgroundColor: viewMode === "responses" ? "#1e3a8a" : "transparent" }}
              >
                <FileText className="w-4 h-4" />
                Responses
              </button>
              <button
                onClick={() => setViewMode("ops")}
                className={`flex items-center gap-2 px-3.5 py-1.5 text-sm font-medium rounded-md transition-all duration-200 ${viewMode === "ops"
                  ? "text-white"
                  : "text-gray-900 dark:text-gray-100 hover:text-black dark:hover:text-white"
                  }`}
                style={{ backgroundColor: viewMode === "ops" ? "#1e3a8a" : "transparent" }}
              >
                <Printer className="w-4 h-4" />
                OPS Template
              </button>
            </div>

            <div className="flex items-center gap-4">
              <div className="text-left">
                <h2 className="text-lg font-bold text-gray-900 dark:text-white leading-tight">
                  {form.title}
                </h2>
                <p className="text-xs text-gray-600 dark:text-gray-400">
                  Comprehensive analysis and insights
                </p>
              </div>
              <div className="h-8 w-px bg-gray-300 dark:bg-gray-600 hidden sm:block"></div>
              <div className="text-left hidden sm:block">
                <p className="text-xs text-gray-500 dark:text-gray-400">
                  Submitted
                </p>
                <p className="text-xs font-semibold text-gray-900 dark:text-white">
                  {formatTimestamp(response.createdAt)}
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={handlePrintOPS}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[9px] font-bold uppercase tracking-wider transition-all"
            >
              <Printer className="h-3 w-3" /> PDF(A3)
            </button>

            <button
              onClick={handleEditResponse}
              className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-white rounded-lg transition-all duration-200 hover:opacity-90"
              style={{ backgroundColor: "#2563eb" }}
              title="Edit Response"
            >
              <Edit2 className="w-4 h-4" />
              <span className="hidden sm:inline">Edit</span>
            </button>

            <button
              onClick={handleDeleteResponse}
              className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-white rounded-lg transition-all duration-200 hover:opacity-90"
              style={{ backgroundColor: "#dc2626" }}
              title="Delete Response"
            >
              <Trash2 className="w-4 h-4" />
              <span className="hidden sm:inline">Delete</span>
            </button>

            {viewMode === "responses" && (
              <button
                onClick={() => handleExportExcel()}
                disabled={exportingExcel}
                className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-white rounded-lg transition-all duration-200 hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed"
                style={{ backgroundColor: "#16a34a" }}
                title="Export to Excel"
              >
                {exportingExcel ? (
                  <div className="w-4 h-4 border-2 border-green-600 border-t-transparent rounded-full animate-spin" />
                ) : (
                  <Download className="w-4 h-4" />
                )}
                <span className="hidden sm:inline">
                  {exportingExcel ? "Exporting..." : "Excel"}
                </span>
              </button>
            )}

            {viewMode === "dashboard" && (
              <div className="relative pdf-type-selector">
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setShowPDFTypeSelector(!showPDFTypeSelector);
                  }}
                  disabled={generatingPDF}
                  className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-white rounded-lg transition-all duration-200 hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed"
                  style={{ backgroundColor: "#0891b2" }}
                  title="Download PDF"
                >
                  {generatingPDF ? (
                    <div className="flex items-center gap-2">
                      <div className="relative w-5 h-5">
                        {/* Spinner */}
                        <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin" />

                        {/* Progress text overlay */}
                        {pdfDownloadProgress !== null && pdfDownloadProgress > 0 && (
                          <div className="absolute inset-0 flex items-center justify-center">

                          </div>
                        )}
                      </div>
                      <span className="hidden sm:inline whitespace-nowrap">
                        {pdfDownloadProgress !== null ? `Downloading..${pdfDownloadProgress}%` : 'Generating...'}
                      </span>
                    </div>
                  ) : (
                    <>
                      <FileText className="w-4 h-4" />
                      <span className="hidden sm:inline">PDF</span>
                      {showPDFTypeSelector && (
                        <ChevronDown className="w-4 h-4 ml-1 transition-transform" />
                      )}
                    </>
                  )}
                </button>

                {showPDFTypeSelector && !generatingPDF && (
                  <div className="absolute top-full mt-2 right-0 w-56 bg-white dark:bg-gray-800 rounded-xl shadow-xl border border-gray-200 dark:border-gray-700 z-50 overflow-hidden animate-in fade-in slide-in-from-top-2 duration-200">
                    <div className="py-1">
                      <div className="px-3 py-2 text-xs font-semibold text-gray-500 dark:text-gray-400 uppercase tracking-wide">
                        {complianceLabels.yes === "Accepted" ? "Compliance Status" : "Response Types"}
                      </div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleDownloadPDF('yes-only');
                        }}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-green-50 dark:hover:bg-green-900/20 transition-colors duration-150"
                      >
                        <CheckCircle className="w-4 h-4 text-green-600 dark:text-green-400 mr-2 flex-shrink-0" />
                        <span>{complianceLabels.yes} Responses (Type 1)</span>
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleDownloadPDF('no-only');
                        }}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors duration-150"
                      >
                        <XCircle className="w-4 h-4 text-red-600 dark:text-red-400 mr-2 flex-shrink-0" />
                        <span>{complianceLabels.no} Responses (Type 2)</span>
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleDownloadPDF('na-only');
                        }}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-yellow-50 dark:hover:bg-yellow-900/20 transition-colors duration-150"
                      >
                        <AlertTriangle className="w-4 h-4 text-yellow-600 dark:text-yellow-400 mr-2 flex-shrink-0" />
                        <span>{complianceLabels.na} Responses (Type 3)</span>
                      </button>
                      <div className="border-t border-gray-200 dark:border-gray-700 my-1"></div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleDownloadPDF('both');
                        }}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition-colors duration-150"
                      >
                        <FileCheck className="w-4 h-4 text-blue-600 dark:text-blue-400 mr-2 flex-shrink-0" />
                        <span>All Response Types (Type 4)</span>
                      </button>
                      <div className="border-t border-gray-200 dark:border-gray-700 my-1"></div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleDownloadPDF('section');
                        }}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-purple-50 dark:hover:bg-purple-900/20 transition-colors duration-150"
                      >
                        <BarChart3 className="w-4 h-4 text-purple-600 dark:text-purple-400 mr-2 flex-shrink-0" />
                        <span>View Sections</span>
                      </button>
                      <div className="border-t border-gray-200 dark:border-gray-700 my-1"></div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleDownloadPDF('responses-view');
                        }}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-cyan-50 dark:hover:bg-cyan-900/20 transition-colors duration-150"
                      >
                        <FileText className="w-4 h-4 text-cyan-600 dark:text-cyan-400 mr-2 flex-shrink-0" />
                        <span>Responses Detail</span>
                      </button>
                      <div className="border-t border-gray-200 dark:border-gray-700 my-1"></div>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setShowPDFTypeSelector(false);
                          handleBulkDownloadZip();
                        }}
                        disabled={exportingZip || generatingPDF}
                        className="flex items-center w-full px-3 py-2 text-left text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors duration-150"
                      >
                        <Download className="w-4 h-4 text-indigo-600 dark:text-indigo-400 mr-2 flex-shrink-0" />
                        <span>Bulk Download (ZIP)</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            <div className="h-6 w-px bg-gray-300 dark:bg-gray-700 mx-1"></div>

            <button
              onClick={() => navigate("/responses/all")}
              className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-gray-600 dark:text-gray-300 hover:text-gray-900 dark:hover:text-white transition-colors"
            >
              <ArrowLeft className="w-4 h-4" />
              <span className="hidden sm:inline">Back</span>
            </button>
          </div>
        </div>

        {/* Content Area */}
        {viewMode === "dashboard" ? (
          filteredSectionStats.length > 0 ? (
            <div className="space-y-5 flex flex-col" style={{ gap: "1.25rem" }}>
              {/* Two-Column Layout: Stats (25%) and Basic Information (75%) */}
              <div className="flex flex-col lg:flex-row gap-5 items-stretch">
                {/* Stats Cards - 25% */}
                <div className="w-full lg:w-1/4 flex flex-col gap-2">
                  <div className="bg-blue-50 dark:bg-blue-900/20 p-2.5 rounded-lg border border-blue-200 dark:border-blue-700 transition-shadow duration-300">
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <p className="text-xs font-semibold text-blue-700 dark:text-blue-300 mb-0.5">
                          {summaryTotals.correct + summaryTotals.wrong > 0
                            ? (complianceLabels.yes === "Accepted" ? "Inspection Score" : "Accuracy Score")
                            : "Overall Score"}
                        </p>
                        <p className="text-lg font-bold text-blue-900 dark:text-blue-100">
                          {(() => {
                            const scoringTotal = summaryTotals.yes + summaryTotals.no + summaryTotals.correct + summaryTotals.wrong;
                            const totalSuccess = summaryTotals.yes + summaryTotals.correct;

                            return scoringTotal > 0 ? ((totalSuccess / scoringTotal) * 100).toFixed(1) : "0.0";
                          })()}%
                        </p>
                      </div>
                      <div className="p-1 bg-blue-100 dark:bg-blue-900/40 rounded-full flex-shrink-0">
                        <Award className="w-3 h-3 text-blue-600 dark:text-blue-400" />
                      </div>
                    </div>
                  </div>

                  <div className="bg-indigo-50 dark:bg-indigo-900/20 p-2.5 rounded-lg border border-indigo-200 dark:border-indigo-700 transition-shadow duration-300 cursor-pointer" onClick={() => setViewMode("responses")}>
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <p className="text-xs font-semibold text-indigo-700 dark:text-indigo-300 mb-0.5">
                          Total Sections
                        </p>
                        <p className="text-lg font-bold text-indigo-900 dark:text-indigo-100">
                          {filteredSectionStats.length}
                        </p>
                      </div>
                      <div className="p-1 bg-indigo-100 dark:bg-indigo-900/40 rounded-full flex-shrink-0">
                        <Target className="w-3 h-3 text-indigo-600 dark:text-indigo-400" />
                      </div>
                    </div>
                  </div>

                  <div className="bg-green-50 dark:bg-green-900/20 p-2.5 rounded-lg border border-green-200 dark:border-green-700 transition-shadow duration-300 cursor-pointer" onClick={() => setExpandResponseRateBreakdown(!expandResponseRateBreakdown)}>
                    <div className="flex items-start justify-between gap-2">
                      <div>
                        <div className="flex items-center gap-1">
                          <p className="text-xs font-semibold text-green-700 dark:text-green-300 mb-0.5">
                            Response Rate
                          </p>
                          <ChevronDown className={`w-3 h-3 text-green-700 dark:text-green-300 transition-transform duration-300 ${expandResponseRateBreakdown ? "rotate-180" : ""}`} />
                        </div>
                        <p className="text-lg font-bold text-green-900 dark:text-green-100">
                          {(() => {
                            const totalQuestions = filteredSectionStats.reduce((sum, stat) => sum + stat.total, 0);
                            const totalAnswered = filteredSectionStats.reduce((sum, stat) => sum + stat.answeredCount, 0);
                            return totalQuestions > 0 ? ((totalAnswered / totalQuestions) * 100).toFixed(1) : "0.0";
                          })()}%
                        </p>
                      </div>
                      <div className="p-1 bg-green-100 dark:bg-green-900/30 rounded-full flex-shrink-0">
                        <Activity className="w-3 h-3 text-green-600 dark:text-green-400" />
                      </div>
                    </div>

                    {expandResponseRateBreakdown && (
                      <div className="mt-2 pt-2 border-t border-green-300 dark:border-green-800">
                        <div className="grid grid-cols-3 gap-2 mt-2">
                          {(() => {
                            const totalYes = filteredSectionStats.reduce((sum, stat) => sum + stat.yes, 0);
                            const totalNo = filteredSectionStats.reduce((sum, stat) => sum + stat.no, 0);
                            const totalNA = filteredSectionStats.reduce((sum, stat) => sum + stat.na, 0);
                            const totalCorrect = filteredSectionStats.reduce((sum, stat) => sum + stat.correct, 0);
                            const totalWrong = filteredSectionStats.reduce((sum, stat) => sum + stat.wrong, 0);
                            const totalAnswered = filteredSectionStats.reduce((sum, stat) => sum + stat.answeredCount, 0);
                            const yesPercent = totalAnswered > 0 ? ((totalYes / totalAnswered) * 100).toFixed(1) : "0.0";
                            const noPercent = totalAnswered > 0 ? ((totalNo / totalAnswered) * 100).toFixed(1) : "0.0";
                            const naPercent = totalAnswered > 0 ? ((totalNA / totalAnswered) * 100).toFixed(1) : "0.0";
                            const correctPercent = totalAnswered > 0 ? ((totalCorrect / totalAnswered) * 100).toFixed(1) : "0.0";
                            const wrongPercent = totalAnswered > 0 ? ((totalWrong / totalAnswered) * 100).toFixed(1) : "0.0";

                            return (
                              <>
                                {summaryTotals.hasAnyYesNo && (
                                  <div className="grid grid-cols-3 gap-2">
                                    <div className="text-center p-1.5 bg-green-100/60 dark:bg-green-900/20 rounded-md">
                                      <p className="text-[10px] font-bold text-green-700 dark:text-green-400 uppercase tracking-tight">{complianceLabels.yes}</p>
                                      <p className="text-sm font-bold text-green-800 dark:text-green-300">{totalYes}</p>
                                      <p className="text-[10px] text-green-700 dark:text-green-400 font-medium">{yesPercent}%</p>
                                    </div>
                                    <div className="text-center p-1.5 bg-red-100/60 dark:bg-red-900/20 rounded-md">
                                      <p className="text-[10px] font-bold text-red-700 dark:text-red-400 uppercase tracking-tight">{complianceLabels.no}</p>
                                      <p className="text-sm font-bold text-red-800 dark:text-red-300">{totalNo}</p>
                                      <p className="text-[10px] text-red-700 dark:text-red-400 font-medium">{noPercent}%</p>
                                    </div>
                                    <div className="text-center p-1.5 bg-yellow-100/60 dark:bg-yellow-900/20 rounded-md">
                                      <p className="text-[10px] font-bold text-yellow-700 dark:text-yellow-400 uppercase tracking-tight">{complianceLabels.na}</p>
                                      <p className="text-sm font-bold text-yellow-800 dark:text-yellow-300">{totalNA}</p>
                                      <p className="text-[10px] text-yellow-700 dark:text-yellow-400 font-medium">{naPercent}%</p>
                                    </div>
                                  </div>
                                )}

                                {/* Accuracy Section */}
                                <div className="col-span-3 mt-1.5 pt-1.5 border-t border-green-200 dark:border-green-800/50">
                                  <p className="text-[10px] font-bold text-green-800 dark:text-green-300 uppercase mb-1 flex items-center gap-1">
                                    <Zap className="w-2.5 h-2.5" /> {complianceLabels.yes === "Accepted" ? "Inspection Statistics" : "Accuracy Statistics"}
                                  </p>
                                  <div className="grid grid-cols-2 gap-2">
                                    <div className="flex items-center justify-between p-1.5 bg-emerald-100/50 dark:bg-emerald-900/20 rounded border border-emerald-200 dark:border-emerald-800/40">
                                      <span className="text-[10px] font-bold text-emerald-700 dark:text-emerald-400 uppercase">{complianceLabels.correct}</span>
                                      <div className="text-right">
                                        <p className="text-xs font-bold text-emerald-800 dark:text-emerald-200 leading-none">{totalCorrect}</p>
                                        <p className="text-[9px] text-emerald-600 dark:text-emerald-400 font-bold">{correctPercent}%</p>
                                      </div>
                                    </div>
                                    <div className="flex items-center justify-between p-1.5 bg-rose-100/50 dark:bg-rose-900/20 rounded border border-rose-200 dark:border-rose-800/40">
                                      <span className="text-[10px] font-bold text-rose-700 dark:text-rose-400 uppercase">{complianceLabels.wrong}</span>
                                      <div className="text-right">
                                        <p className="text-xs font-bold text-rose-800 dark:text-rose-200 leading-none">{totalWrong}</p>
                                        <p className="text-[9px] text-rose-600 dark:text-rose-400 font-bold">{wrongPercent}%</p>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </>
                            );
                          })()}
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="bg-purple-50 dark:bg-purple-900/20 p-2.5 rounded-lg border border-purple-200 dark:border-purple-700 transition-shadow duration-300">
                    <div className="flex items-start justify-between gap-2">
                      <div className="min-w-0 flex-1">
                        <p className="text-xs font-semibold text-purple-700 dark:text-purple-300 mb-0.5">
                          Location
                        </p>
                        <p className="text-xs font-medium text-purple-900 dark:text-purple-100 truncate">
                          {response.submissionMetadata?.location
                            ? (() => {
                              const loc = response.submissionMetadata.location;
                              const parts = [];
                              if (loc.city) parts.push(loc.city);
                              if (loc.region) parts.push(loc.region);
                              if (loc.country) parts.push(loc.country);
                              return parts.length > 0 ? parts.join(", ") : "Unavailable";
                            })()
                            : "Disabled"
                          }
                        </p>
                      </div>
                      <div className="p-1 bg-purple-100 dark:bg-purple-900/30 rounded-full flex-shrink-0">
                        <MapPin className="w-3 h-3 text-purple-600 dark:text-purple-400" />
                      </div>
                    </div>
                  </div>
                </div>

                {/* Basic Information - 75% */}
                <div className="w-full lg:w-3/4 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5">
                  {form?.sections && form.sections.length > 0 ? (
                    (() => {
                      const section = form.sections[0];
                      return (
                        <div key={section.id || 0}>
                          <div className="mb-4 pb-4 border-b border-gray-200 dark:border-gray-700">
                            <h4 className="text-sm font-semibold text-indigo-900 dark:text-indigo-100 mb-1">
                              {section.title || "Section 1"}
                            </h4>
                          </div>

                          {section.questions && section.questions.length > 0 ? (
                            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                              {section.questions.map((question: any) => {
                                const answer = response.answers?.[question.id];
                                const isMainQuestion = question && !question.parentId && !question.showWhen?.questionId;
                                return (
                                  <div
                                    key={question.id}
                                    className={`p-3 rounded-lg border transition-shadow ${isMainQuestion
                                      ? "bg-gradient-to-br from-blue-50 to-blue-100 dark:from-blue-900/30 dark:to-blue-800/30 border-blue-200 dark:border-blue-700"
                                      : "bg-gray-50 dark:bg-gray-800/50 border-gray-200 dark:border-gray-700"
                                      }`}
                                  >
                                    <div className="flex flex-col gap-1 mb-1">
                                      {question.subParam1 && (
                                        <span className="inline-block bg-blue-100/60 dark:bg-blue-900/30 text-blue-900 dark:text-blue-200 px-2 py-0.5 rounded font-semibold text-xs w-fit">
                                          {question.subParam1}
                                        </span>
                                      )}
                                      <p className="text-xs font-semibold text-gray-600 dark:text-gray-400 uppercase tracking-wide">
                                        {question.text || question.label || question.id}
                                      </p>
                                    </div>
                                    <div className="mt-1 flex flex-col gap-1">
                                      {answer !== undefined && answer !== null && answer !== ''
                                        ? (
                                          <>
                                            {renderHighlightedAnswer(answer, question)}
                                            {question.trackResponseRank && response.responseRanks?.[question.id] && (
                                              <div className="flex flex-col gap-1 mt-2">
                                                {question.trackResponseRankLabel && (
                                                  <span className="text-[10px] font-bold text-blue-600 dark:text-blue-400 uppercase tracking-tight leading-none">
                                                    {question.trackResponseRankLabel}
                                                  </span>
                                                )}
                                                <span className={`text-[10px] font-bold min-w-[24px] h-6 px-1.5 rounded-full flex items-center justify-center border shadow-sm w-fit ${getRankStyle(answer, document.documentElement.classList.contains("dark"))}`}>
                                                  #{response.responseRanks[question.id]}
                                                </span>
                                              </div>
                                            )}
                                          </>
                                        )
                                        : <span className="text-gray-400 italic text-xs">No answer</span>
                                      }
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          ) : (
                            <div className="text-center py-4 text-xs text-gray-500 dark:text-gray-400">
                              No questions in this section
                            </div>
                          )}
                        </div>
                      );
                    })()
                  ) : (
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                      {response.dealerName && (
                        <div className="p-3 rounded-lg bg-gray-50 dark:bg-gray-800/50 border border-gray-200 dark:border-gray-700">
                          <p className="text-xs font-semibold text-gray-600 dark:text-gray-400 uppercase tracking-wide mb-1">
                            Dealer Name
                          </p>
                          <p className="text-sm font-medium text-gray-900 dark:text-white break-words">
                            {response.dealerName}
                          </p>
                        </div>
                      )}
                      {response.answers?.dealerCode && (
                        <div className="p-3 rounded-lg bg-gray-50 dark:bg-gray-800/50 border border-gray-200 dark:border-gray-700">
                          <p className="text-xs font-semibold text-gray-600 dark:text-gray-400 uppercase tracking-wide mb-1">
                            Dealer Code
                          </p>
                          <p className="text-sm font-medium text-gray-900 dark:text-white break-words">
                            {renderHighlightedAnswer(response.answers.dealerCode)}
                          </p>
                        </div>
                      )}
                      {response.answers?.location && (
                        <div className="p-3 rounded-lg bg-gray-50 dark:bg-gray-800/50 border border-gray-200 dark:border-gray-700">
                          <p className="text-xs font-semibold text-gray-600 dark:text-gray-400 uppercase tracking-wide mb-1">
                            Location
                          </p>
                          <p className="text-sm font-medium text-gray-900 dark:text-white break-words">
                            {renderHighlightedAnswer(response.answers.location)}
                          </p>
                        </div>
                      )}
                      {response.answers?.auditorDate && (
                        <div className="p-3 rounded-lg bg-gray-50 dark:bg-gray-800/50 border border-gray-200 dark:border-gray-700">
                          <p className="text-xs font-semibold text-gray-600 dark:text-gray-400 uppercase tracking-wide mb-1">
                            Auditor Date
                          </p>
                          <p className="text-sm font-medium text-gray-900 dark:text-white break-words">
                            {renderHighlightedAnswer(response.answers.auditorDate)}
                          </p>
                        </div>
                      )}
                      {response.answers?.auditorName && (
                        <div className="p-3 rounded-lg bg-gray-50 dark:bg-gray-800/50 border border-gray-200 dark:border-gray-700">
                          <p className="text-xs font-semibold text-gray-600 dark:text-gray-400 uppercase tracking-wide mb-1">
                            Auditor Name
                          </p>
                          <p className="text-sm font-medium text-gray-900 dark:text-white break-words">
                            {renderHighlightedAnswer(response.answers.auditorName)}
                          </p>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>

              {/* Charts Section */}
              <div className="flex flex-col lg:flex-row gap-4 items-stretch">
                <div className="flex-shrink-0 lg:w-[70%]">

                  {/* Section-wise Breakdown Table */}
                  <div className="bg-white dark:bg-gray-900 rounded-lg border border-gray-200 dark:border-gray-700 overflow-hidden">
                    <div className="bg-primary-600 p-6">
                      <div className="flex items-center justify-between">
                        <div>
                          <h3 className="text-2xl font-bold text-white flex items-center">
                            <BarChart3 className="w-7 h-7 mr-3" />
                            Section-wise Breakdown
                          </h3>
                          <p className="text-blue-100 mt-1">
                            Detailed performance analysis by section
                          </p>
                        </div>
                        <div className="flex items-center gap-2">
                        </div>
                      </div>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full divide-y divide-gray-200 dark:divide-gray-700 text-sm">
                        <thead className="bg-gradient-to-r from-gray-50 to-gray-100 dark:from-gray-700 dark:to-gray-800 sticky top-0">
                          <tr>
                            <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-48">
                              Section
                            </th>
                            <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-20">
                              Total
                            </th>
                            {summaryTotals.hasAnyQuiz && (
                              <>
                                <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-24 text-green-600 dark:text-green-400">
                                  {complianceLabels.correct}
                                </th>
                                <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-24 text-red-600 dark:text-red-400">
                                  {complianceLabels.wrong}
                                </th>
                              </>
                            )}
                            {summaryTotals.hasAnyYesNo && (
                              <>
                                <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-24">
                                  {complianceLabels.yes}
                                </th>
                                <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-24">
                                  {complianceLabels.no}
                                </th>
                                <th className="px-6 py-5 text-left font-bold text-gray-900 dark:text-gray-100 uppercase tracking-wider min-w-24">
                                  {complianceLabels.na}
                                </th>
                              </>
                            )}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200 dark:divide-gray-700 bg-white dark:bg-gray-900">
                          {sectionSummaryRows.map((row) => (
                            <tr
                              key={row.id}
                              onClick={() => {
                                setAutoOpenSectionId(null);
                                setTimeout(() => setAutoOpenSectionId(row.id), 10);
                              }}
                              className="group hover:bg-gradient-to-r hover:from-blue-50 hover:to-purple-50 dark:hover:from-gray-700 dark:hover:to-gray-600 transition-all duration-300 bg-white dark:bg-gray-900 cursor-pointer"
                            >
                              <td className="px-6 py-5 font-bold text-gray-900 dark:text-gray-100 flex items-center">
                                <div className="w-3 h-3 bg-blue-500 rounded-full mr-3"></div>
                                <span className="text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300 transition-colors">
                                  {row.title}
                                </span>
                              </td>
                              <td className="px-6 py-5 text-gray-700 dark:text-gray-300 font-medium">
                                {row.total}
                              </td>
                              {summaryTotals.hasAnyQuiz && (
                                <>
                                  <td className="px-6 py-5 text-green-600 dark:text-green-400 font-bold">
                                    {row.hasQuiz ? `${row.correct} (${row.correctPercent.toFixed(1)}%)` : "-"}
                                  </td>
                                  <td className="px-6 py-5 text-red-600 dark:text-red-400 font-bold">
                                    {row.hasQuiz ? `${row.wrong} (${row.wrongPercent.toFixed(1)}%)` : "-"}
                                  </td>
                                </>
                              )}
                              {summaryTotals.hasAnyYesNo && (
                                <>
                                  <td className="px-6 py-5 text-gray-700 dark:text-gray-300 font-medium">
                                    {row.hasYesNo ? `${row.yes} (${row.yesPercent.toFixed(1)}%)` : "-"}
                                  </td>
                                  <td className="px-6 py-5 text-gray-700 dark:text-gray-300 font-medium">
                                    {row.hasYesNo ? `${row.no} (${row.noPercent.toFixed(1)}%)` : "-"}
                                  </td>
                                  <td className="px-6 py-5 text-gray-700 dark:text-gray-300 font-medium">
                                    {row.hasYesNo ? `${row.na} (${row.naPercent.toFixed(1)}%)` : "-"}
                                  </td>
                                </>
                              )}
                            </tr>
                          ))}

                          {/* Total Row */}
                          <tr className="bg-gray-50 dark:bg-gray-800/50 font-bold border-t-2 border-gray-300 dark:border-gray-600">
                            <td className="px-6 py-5 text-gray-900 dark:text-gray-100 flex items-center">
                              <div className="w-3 h-3 bg-indigo-600 rounded-full mr-3"></div>
                              <span>TOTAL</span>
                            </td>
                            <td className="px-6 py-5 text-gray-900 dark:text-gray-100 font-bold">
                              {summaryTotals.total}
                            </td>
                            {summaryTotals.hasAnyQuiz && (
                              <>
                                <td className="px-6 py-5 text-green-600 dark:text-green-400 font-bold">
                                  {summaryTotals.correct} ({summaryTotals.total > 0 ? ((summaryTotals.correct / summaryTotals.total) * 100).toFixed(1) : 0}%)
                                </td>
                                <td className="px-6 py-5 text-red-600 dark:text-red-400 font-bold">
                                  {summaryTotals.wrong} ({summaryTotals.total > 0 ? ((summaryTotals.wrong / summaryTotals.total) * 100).toFixed(1) : 0}%)
                                </td>
                              </>
                            )}
                            {summaryTotals.hasAnyYesNo && (
                              <>
                                <td className="px-6 py-5 text-gray-900 dark:text-gray-100 font-bold">
                                  {summaryTotals.yes} ({summaryTotals.total > 0 ? ((summaryTotals.yes / summaryTotals.total) * 100).toFixed(1) : 0}%)
                                </td>
                                <td className="px-6 py-5 text-gray-900 dark:text-gray-100 font-bold">
                                  {summaryTotals.no} ({summaryTotals.total > 0 ? ((summaryTotals.no / summaryTotals.total) * 100).toFixed(1) : 0}%)
                                </td>
                                <td className="px-6 py-5 text-gray-900 dark:text-gray-100 font-bold">
                                  {summaryTotals.na} ({summaryTotals.total > 0 ? ((summaryTotals.na / summaryTotals.total) * 100).toFixed(1) : 0}%)
                                </td>
                              </>
                            )}
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
                <div className="lg:w-[30%] flex flex-col gap-4">
                  <div className="bg-gradient-to-br from-white via-blue-50/30 to-indigo-50/30 dark:from-gray-800 dark:via-blue-900/10 dark:to-indigo-900/10 p-2 rounded-2xl border border-blue-200/50 dark:border-blue-700/50 transform hover:scale-[1.02] transition-all duration-500 backdrop-blur-sm w-full flex flex-col">
                    <div className="flex flex-col items-center justify-center mb-1 gap-1">
                      <h3 className="text-xs font-bold text-gray-900 dark:text-white flex flex-col items-center text-center">
                        <span>Section Performance</span>
                      </h3>
                      <div className="flex gap-1 bg-white/50 dark:bg-gray-700/50 rounded-md px-1.5 py-1 flex-row">
                        {summaryTotals.correct > 0 && (
                          <div className="flex items-center space-x-0.5">
                            <div className="w-1.5 h-1.5 bg-green-500 rounded-full"></div>
                            <span className="text-[8px] font-medium text-gray-700 dark:text-gray-300">{complianceLabels.correct}</span>
                          </div>
                        )}
                        {summaryTotals.wrong > 0 && summaryTotals.correct > 0 && (
                          <div className="flex items-center space-x-0.5">
                            <div className="w-1.5 h-1.5 bg-red-500 rounded-full"></div>
                            <span className="text-[8px] font-medium text-gray-700 dark:text-gray-300">{complianceLabels.wrong}</span>
                          </div>
                        )}
                        <div className="flex items-center space-x-0.5">
                          <div className="w-1.5 h-1.5 bg-blue-500 rounded-full"></div>
                          <span className="text-[8px] font-medium text-gray-700 dark:text-gray-300">{complianceLabels.yes}</span>
                        </div>
                        <div className="flex items-center space-x-0.5">
                          <div className="w-1.5 h-1.5 bg-blue-400 rounded-full"></div>
                          <span className="text-[8px] font-medium text-gray-700 dark:text-gray-300">{complianceLabels.no}</span>
                        </div>
                        <div className="flex items-center space-x-0.5">
                          <div className="w-1.5 h-1.5 bg-blue-300 rounded-full"></div>
                          <span className="text-[8px] font-medium text-gray-700 dark:text-gray-300">{complianceLabels.na}</span>
                        </div>
                      </div>
                    </div>
                    <div id="section-performance-chart" className="w-full h-48 flex items-center justify-center">
                      <Radar data={sectionChartData} options={{ ...sectionChartOptions, maintainAspectRatio: false }} />
                    </div>
                  </div>
                </div>
              </div>










              {/* Section - Yes/No/N/A Analysis for ALL Sections */}
              {form?.sections?.map((section: any) => {
                if (!section) return null;

                const questionStats = getSectionYesNoQuestionStats(section.id);
                if (questionStats.length === 0) return null;

                const hasAnyYesNoInSection = questionStats.some(q => q.hasYesNo);
                const yesNoStats = questionStats.filter(q => q.hasYesNo && !q.isQuiz);
                const quizStats = questionStats.filter(q => q.isQuiz);
                // Also handle other types that might not be either but have stats
                const otherStats = questionStats.filter(q => !q.hasYesNo && !q.isQuiz);

                const sectionTotals = questionStats.reduce(
                  (totals, stat) => ({
                    yes: totals.yes + stat.yes,
                    no: totals.no + stat.no,
                    na: totals.na + stat.na,
                    correct: totals.correct + stat.correct,
                    wrong: totals.wrong + stat.wrong,
                    total: totals.total + stat.total,
                  }),
                  { yes: 0, no: 0, na: 0, total: 0, correct: 0, wrong: 0 }
                );

                const sectionPercentages = {
                  yes:
                    sectionTotals.total > 0
                      ? ((sectionTotals.yes / sectionTotals.total) * 100).toFixed(1)
                      : "0.0",
                  no:
                    sectionTotals.total > 0
                      ? ((sectionTotals.no / sectionTotals.total) * 100).toFixed(1)
                      : "0.0",
                  na:
                    sectionTotals.total > 0
                      ? ((sectionTotals.na / sectionTotals.total) * 100).toFixed(1)
                      : "0.0",
                  correct:
                    sectionTotals.total > 0
                      ? ((sectionTotals.correct / sectionTotals.total) * 100).toFixed(1)
                      : "0.0",
                  wrong:
                    sectionTotals.total > 0
                      ? ((sectionTotals.wrong / sectionTotals.total) * 100).toFixed(1)
                      : "0.0",
                };

                const hasYesNo = yesNoStats.length > 0;
                const hasQuiz = quizStats.length > 0;

                const labels: string[] = [];
                const data: number[] = [];
                const colors: string[] = [];

                if (hasYesNo && hasQuiz && sectionTotals.correct > 0) {
                  labels.push(complianceLabels.correct, complianceLabels.wrong, complianceLabels.yes, complianceLabels.no);
                  data.push(sectionTotals.correct, sectionTotals.wrong, sectionTotals.yes, sectionTotals.no);
                  colors.push("#10b981", "#ef4444", "#1e40af", "#3b82f6");
                  if (sectionTotals.na > 0) {
                    labels.push(complianceLabels.na);
                    data.push(sectionTotals.na);
                    colors.push("#93c5fd");
                  }
                } else if (hasYesNo) {
                  labels.push(complianceLabels.yes, complianceLabels.no);
                  data.push(sectionTotals.yes, sectionTotals.no);
                  colors.push("#1e40af", "#3b82f6");
                  if (sectionTotals.na > 0) {
                    labels.push(complianceLabels.na);
                    data.push(sectionTotals.na);
                    colors.push("#93c5fd");
                  }
                } else {
                  labels.push(complianceLabels.correct, complianceLabels.wrong);
                  data.push(sectionTotals.correct, sectionTotals.wrong);
                  colors.push("#16a34a", "#dc2626");
                  if (sectionTotals.na > 0) {
                    labels.push(complianceLabels.na);
                    data.push(sectionTotals.na);
                    colors.push("#94a3b8");
                  }
                }

                const chartData = {
                  labels,
                  datasets: [
                    {
                      data,
                      backgroundColor: colors,
                      borderColor: colors,
                      borderWidth: 2,
                    },
                  ],
                };

                const chartOptions = {
                  responsive: true,
                  maintainAspectRatio: false,
                  plugins: {
                    legend: {
                      position: "bottom" as const,
                      labels: {
                        color: document.documentElement.classList.contains("dark")
                          ? "#d1d5db"
                          : "#374151",
                        boxWidth: 10,
                        padding: 10,
                        font: { size: 10 }
                      },
                    },
                    tooltip: {
                      callbacks: {
                        label: (context: any) => {
                          const total = sectionTotals.total;
                          const value =
                            typeof context.parsed === "number"
                              ? context.parsed
                              : context.parsed?.y || 0;
                          const percentage =
                            total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                          return `${context.label}: ${value} (${percentage}%)`;
                        },
                      },
                    },
                  },
                };

                return (
                  <div key={section.id} id={`section-detail-${section.id}`} className="bg-white dark:bg-gray-900 p-5 rounded-xl border border-gray-200 dark:border-gray-800">
                    <div className="mb-4">
                      <h3 className="text-lg font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        {section.title || "Section"} - {(hasAnyYesNoInSection && quizStats.length > 0) ? "Compliance & Accuracy Analysis" : hasAnyYesNoInSection ? "Compliance Analysis" : "Accuracy Analysis"}
                      </h3>
                      <p className="text-xs text-gray-600 dark:text-gray-400 mt-1">
                        {(hasAnyYesNoInSection && quizStats.length > 0)
                          ? `Question-wise breakdown of compliance (${complianceLabels.yes}/${complianceLabels.no}) and accuracy (Correct/Wrong) with section summary`
                          : hasAnyYesNoInSection
                            ? `Question-wise breakdown of compliance (${complianceLabels.yes}/${complianceLabels.no}/${complianceLabels.na}) with overall section summary`
                            : "Question-wise breakdown of quiz accuracy with overall section summary"}
                      </p>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {/* Chart */}
                      <div className="bg-gray-50 dark:bg-gray-800 p-4 rounded-lg border border-gray-200 dark:border-gray-800">
                        <div className="flex items-center justify-between mb-3">
                          <h4 className="text-sm font-semibold text-gray-900 dark:text-white flex items-center">
                            <PieChart className="w-4 h-4 mr-2" />
                            Response Distribution
                          </h4>
                          <select
                            value={sectionChartTypes[section.id] || "pie"}
                            onChange={(e) =>
                              setSectionChartTypes((prev) => ({
                                ...prev,
                                [section.id]: e.target.value as "pie" | "bar",
                              }))
                            }
                            className="px-2.5 py-1 text-xs bg-white dark:bg-gray-700 border border-gray-300 dark:border-gray-600 rounded text-gray-900 dark:text-white font-medium focus:outline-none focus:ring-1 focus:ring-indigo-500"
                          >
                            <option value="pie">Pie Chart</option>
                            <option value="bar">Bar Chart</option>
                          </select>
                        </div>
                        <div id={`section-chart-${section.id}`} className="w-full h-56">
                          {sectionChartTypes[section.id] === "bar" ? (
                            <Bar
                              data={{
                                labels: questionStats.map(
                                  (stat) => stat.subParam1 || "No parameter"
                                ),
                                datasets: [
                                  {
                                    label: complianceLabels.correct,
                                    data: questionStats.map((stat) => stat.correct),
                                    backgroundColor: "#10b981",
                                    borderColor: "#10b981",
                                    borderWidth: 1,
                                    hidden: !hasQuiz,
                                  },
                                  {
                                    label: complianceLabels.wrong,
                                    data: questionStats.map((stat) => stat.wrong),
                                    backgroundColor: "#ef4444",
                                    borderColor: "#ef4444",
                                    borderWidth: 1,
                                    hidden: !hasQuiz,
                                  },
                                  {
                                    label: complianceLabels.yes,
                                    data: questionStats.map((stat) => stat.yes),
                                    backgroundColor: "#1e40af",
                                    borderColor: "#1e40af",
                                    borderWidth: 1,
                                    hidden: !hasYesNo,
                                  },
                                  {
                                    label: complianceLabels.no,
                                    data: questionStats.map((stat) => stat.no),
                                    backgroundColor: "#3b82f6",
                                    borderColor: "#3b82f6",
                                    borderWidth: 1,
                                    hidden: !hasYesNo,
                                  },
                                  {
                                    label: complianceLabels.na,
                                    data: questionStats.map((stat) => stat.na),
                                    backgroundColor: "#93c5fd",
                                    borderColor: "#93c5fd",
                                    borderWidth: 1,
                                    hidden: !hasYesNo || sectionTotals.na === 0,
                                  },
                                ],
                              }}
                              options={{
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: {
                                  legend: {
                                    position: "top" as const,
                                    labels: {
                                      color: document.documentElement.classList.contains(
                                        "dark"
                                      )
                                        ? "#d1d5db"
                                        : "#374151",
                                    },
                                  },
                                },
                              }}
                            />
                          ) : (
                            <Pie data={chartData} options={chartOptions} />
                          )}
                        </div>
                      </div>

                      {/* Compliance Analysis Table (Yes/No/NA) */}
                      {hasAnyYesNoInSection && (
                        <div className="bg-white dark:bg-gray-800 rounded-lg border border-gray-200 dark:border-gray-800 overflow-hidden mb-4">
                          <div className="bg-blue-50 dark:bg-blue-900/20 px-4 py-2.5 border-b border-blue-100 dark:border-blue-800">
                            <h4 className="text-sm font-semibold text-blue-900 dark:text-blue-100 flex items-center">
                              <CheckCircle className="w-4 h-4 mr-2" />
                              Compliance Analysis ({complianceLabels.yes}/{complianceLabels.no}/{complianceLabels.na})
                            </h4>
                          </div>
                          <div className="overflow-auto max-h-60">
                            <table className="w-full divide-y divide-gray-200 dark:divide-gray-800 text-xs">
                              <thead className="bg-gray-50 dark:bg-gray-700 sticky top-0 z-10">
                                <tr>
                                  <th className="px-3 py-2 text-left font-semibold text-gray-700 dark:text-gray-200 uppercase tracking-wider w-[40%]">
                                    Parameter
                                  </th>
                                  <th className="px-3 py-2 text-center font-semibold text-blue-700 dark:text-blue-300 uppercase tracking-wider w-[20%]">
                                    {complianceLabels.yes}
                                  </th>
                                  <th className="px-3 py-2 text-center font-semibold text-blue-500 dark:text-blue-400 uppercase tracking-wider w-[20%]">
                                    {complianceLabels.no}
                                  </th>
                                  <th className="px-3 py-2 text-center font-semibold text-gray-500 dark:text-gray-400 uppercase tracking-wider w-[20%]">
                                    {complianceLabels.na}
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-gray-200 dark:divide-gray-800 bg-white dark:bg-gray-900">
                                {questionStats.filter(q => q.hasYesNo).map((stat, index) => {
                                  const total = stat.yes + stat.no + stat.na;
                                  const yesPercent = total > 0 ? ((stat.yes / total) * 100).toFixed(1) : 0;
                                  const noPercent = total > 0 ? ((stat.no / total) * 100).toFixed(1) : 0;
                                  const naPercent = total > 0 ? ((stat.na / total) * 100).toFixed(1) : 0;
                                  return (
                                    <tr key={stat.id} className={index % 2 === 0 ? "bg-white dark:bg-gray-900" : "bg-gray-50/50 dark:bg-gray-800/30"}>
                                      <td className="px-3 py-2 font-medium text-gray-900 dark:text-gray-100">
                                        {stat.subParam1 || stat.title}
                                      </td>
                                      <td className="px-3 py-2 text-center">
                                        <div className="flex flex-col items-center">
                                          <span className={`font-bold ${stat.yes > 0 ? "text-blue-700 dark:text-blue-400" : "text-gray-400"}`}>{stat.yes}</span>
                                          {stat.yes > 0 && <span className="text-[10px] text-blue-600/70">{yesPercent}%</span>}
                                        </div>
                                      </td>
                                      <td className="px-3 py-2 text-center">
                                        <div className="flex flex-col items-center">
                                          <span className={`font-bold ${stat.no > 0 ? "text-blue-500 dark:text-blue-400" : "text-gray-400"}`}>{stat.no}</span>
                                          {stat.no > 0 && <span className="text-[10px] text-blue-500/70">{noPercent}%</span>}
                                        </div>
                                      </td>
                                      <td className="px-3 py-2 text-center">
                                        <div className="flex flex-col items-center">
                                          <span className={`font-bold ${stat.na > 0 ? "text-gray-600 dark:text-gray-400" : "text-gray-400"}`}>{stat.na}</span>
                                          {stat.na > 0 && <span className="text-[10px] text-gray-500/70">{naPercent}%</span>}
                                        </div>
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {/* Accuracy Analysis Table (Correct/Wrong) */}
                      {questionStats.some(q => q.isQuiz) && (
                        <div className="bg-white dark:bg-gray-800 rounded-lg border border-gray-200 dark:border-gray-800 overflow-hidden">
                          <div className="bg-green-50 dark:bg-green-900/20 px-4 py-2.5 border-b border-green-100 dark:border-green-800">
                            <h4 className="text-sm font-semibold text-green-900 dark:text-green-100 flex items-center">
                              <Award className="w-4 h-4 mr-2" />
                              {complianceLabels.yes === "Accepted" ? "Inspection Analysis (Accepted/Rejected)" : "Accuracy Analysis (Correct/Wrong)"}
                            </h4>
                          </div>
                          <div className="overflow-auto max-h-60">
                            <table className="w-full divide-y divide-gray-200 dark:divide-gray-800 text-xs">
                              <thead className="bg-gray-50 dark:bg-gray-700 sticky top-0 z-10">
                                <tr>
                                  <th className="px-3 py-2 text-left font-semibold text-gray-700 dark:text-gray-200 uppercase tracking-wider w-[40%]">
                                    Parameter
                                  </th>
                                  <th className="px-3 py-2 text-center font-semibold text-green-700 dark:text-green-400 uppercase tracking-wider w-[30%]">
                                    {complianceLabels.correct}
                                  </th>
                                  <th className="px-3 py-2 text-center font-semibold text-red-600 dark:text-red-400 uppercase tracking-wider w-[30%]">
                                    {complianceLabels.wrong}
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-gray-200 dark:divide-gray-800 bg-white dark:bg-gray-900">
                                {questionStats.filter(q => q.isQuiz).map((stat, index) => {
                                  const total = stat.correct + stat.wrong;
                                  const correctPercent = total > 0 ? ((stat.correct / total) * 100).toFixed(1) : 0;
                                  const wrongPercent = total > 0 ? ((stat.wrong / total) * 100).toFixed(1) : 0;
                                  return (
                                    <tr key={stat.id} className={index % 2 === 0 ? "bg-white dark:bg-gray-900" : "bg-gray-50/50 dark:bg-gray-800/30"}>
                                      <td className="px-3 py-2 font-medium text-gray-900 dark:text-gray-100">
                                        {stat.subParam1 || stat.title}
                                      </td>
                                      <td className="px-3 py-2 text-center">
                                        <div className="flex flex-col items-center">
                                          <span className={`font-bold ${stat.correct > 0 ? "text-green-600 dark:text-green-400" : "text-gray-400"}`}>{stat.correct}</span>
                                          {stat.correct > 0 && <span className="text-[10px] text-green-600/70">{correctPercent}%</span>}
                                        </div>
                                      </td>
                                      <td className="px-3 py-2 text-center">
                                        <div className="flex flex-col items-center">
                                          <span className={`font-bold ${stat.wrong > 0 ? "text-red-600 dark:text-red-400" : "text-gray-400"}`}>{stat.wrong}</span>
                                          {stat.wrong > 0 && <span className="text-[10px] text-red-600/70">{wrongPercent}%</span>}
                                        </div>
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                    </div>

                    {/* Main Parameters Table */}
                    {(() => {
                      const allSectionQuestions = getSectionQuestionsWithFollowUps(section.id);

                      // Filter: Only show main questions that have at least one answered follow-up question
                      // We check if any of the follow-ups for THIS main question have an actual answer
                      const sectionQuestions = allSectionQuestions.filter((q: any) => {
                        return q.followUpQuestions?.some((fq: any) => {
                          const checkIsImage = (val: any): boolean => {
                            if (!val) return false;
                            if (Array.isArray(val)) return val.some(v => checkIsImage(v));
                            if (typeof val === 'object') {
                              if (val.url && isImageUrl(String(val.url))) return true;
                              if (val.answer && isImageUrl(String(val.answer))) return true;
                              return Object.values(val).some(v => checkIsImage(v));
                            }
                            return isImageUrl(String(val));
                          };

                          const hasActualAnswer = fq.answer !== undefined && fq.answer !== null && fq.answer !== "" &&
                            fq.answer !== "N/A" && fq.answer !== "n/a" &&
                            String(fq.answer).toLowerCase() !== complianceLabels.na.toLowerCase();

                          return hasActualAnswer || checkIsImage(fq.answer);
                        });
                      });

                      if (sectionQuestions.length === 0) {
                        return null;
                      }

                      const allFollowUpIds = new Set<string>();
                      const followUpIdAnswerStatus = new Map<string, boolean>();

                      sectionQuestions.forEach((q: any) => {
                        q.followUpQuestions.forEach((fq: any) => {
                          allFollowUpIds.add(fq.id);

                          // Check if it's an image (recursive check)
                          const checkIsImage = (val: any): boolean => {
                            if (!val) return false;
                            if (Array.isArray(val)) return val.some(v => checkIsImage(v));
                            if (typeof val === 'object') {
                              if (val.url && isImageUrl(String(val.url))) return true;
                              if (val.answer && isImageUrl(String(val.answer))) return true;
                              return Object.values(val).some(v => checkIsImage(v));
                            }
                            return isImageUrl(String(val));
                          };

                          const hasActualAnswer = fq.answer !== undefined && fq.answer !== null && fq.answer !== "" &&
                            fq.answer !== "N/A" && fq.answer !== "n/a" &&
                            String(fq.answer).toLowerCase() !== complianceLabels.na.toLowerCase();

                          if (hasActualAnswer || checkIsImage(fq.answer)) {
                            followUpIdAnswerStatus.set(fq.id, true);
                          }
                        });
                      });

                      const followUpIdsWithAnswers = Array.from(allFollowUpIds).filter(
                        (id) => followUpIdAnswerStatus.get(id) === true
                      );

                      const followUpsBySubParam: Map<
                        string,
                        Array<{ id: string; subParam1?: string; answer?: any }>
                      > = new Map();

                      followUpIdsWithAnswers.forEach((followUpId) => {
                        const followUpObj = sectionQuestions
                          .flatMap((q: any) => q.followUpQuestions)
                          .find((fq: any) => fq.id === followUpId);

                        const subParamKey = followUpObj?.subParam1 || followUpId;
                        if (!followUpsBySubParam.has(subParamKey)) {
                          followUpsBySubParam.set(subParamKey, []);
                        }
                        followUpsBySubParam.get(subParamKey)!.push({
                          id: followUpId,
                          subParam1: followUpObj?.subParam1,
                          answer: followUpObj?.answer,
                        });
                      });

                      const uniqueSubParams = Array.from(followUpsBySubParam.keys());

                      const hasImages = Array.from(followUpsBySubParam.values()).some(
                        (items) => items.some((item) => {
                          const checkValue = (val: any): boolean => {
                            if (!val) return false;
                            if (Array.isArray(val)) return val.some(v => checkValue(v));
                            if (typeof val === 'object') {
                              if (val.url && isImageUrl(String(val.url))) return true;
                              if (val.answer && isImageUrl(String(val.answer))) return true;
                              return Object.values(val).some(v => checkValue(v));
                            }
                            return isImageUrl(String(val));
                          };
                          return checkValue(item.answer);
                        })
                      );

                      return (
                        <div className="bg-gradient-to-br from-blue-50 to-blue-50 dark:from-blue-900/20 dark:to-blue-900/20 p-8 rounded-3xl border border-blue-200 dark:border-blue-800 mt-4">
                          <div className="mb-6 flex items-center justify-between">
                            <h3 className="text-2xl font-bold text-blue-900 dark:text-blue-100 flex items-center gap-3">
                              <div className="w-1 h-8 bg-blue-600 rounded-full"></div>
                              {section.title || "Section"} - Main Parameters
                            </h3>
                            {hasImages && (
                              <button
                                onClick={() =>
                                  setShowMainParamsImages((prev) => ({
                                    ...prev,
                                    [section.id]: !prev[section.id],
                                  }))
                                }
                                className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-colors ${showMainParamsImages[section.id]
                                  ? "bg-blue-600 text-white hover:bg-blue-700"
                                  : "bg-gray-300 text-gray-700 hover:bg-gray-400 dark:bg-gray-600 dark:text-gray-200"
                                  }`}
                              >
                                {showMainParamsImages[section.id] ? "Hide Images" : "View Images"}
                              </button>
                            )}
                          </div>

                          {allFollowUpIds.size === 0 && sectionQuestions.length > 0 && (
                            <div className="mt-3 p-3 bg-yellow-100 dark:bg-yellow-900/30 border border-yellow-300 rounded text-sm text-yellow-800 dark:text-yellow-200">
                              <strong>⚠️ No follow-up questions found</strong> for{" "}
                              {sectionQuestions.length} main question(s)
                            </div>
                          )}

                          <div className="overflow-x-auto">
                            <table className="w-full border-collapse">
                              <thead>
                                <tr className="bg-blue-200 dark:bg-blue-800/50">
                                  <th className="px-6 py-3 text-left text-blue-900 dark:text-blue-100 font-semibold border border-blue-300 dark:border-blue-700 min-w-64">
                                    Main Parameters
                                  </th>
                                  {uniqueSubParams.map((subParam) => (
                                    <th
                                      key={subParam}
                                      className="px-4 py-3 text-left text-blue-900 dark:text-blue-100 font-semibold border border-blue-300 dark:border-blue-700 min-w-48 bg-blue-50 dark:bg-blue-900/30"
                                    >
                                      <span className="text-xs font-bold uppercase tracking-wider text-blue-700 dark:text-blue-300">
                                        {subParam}
                                      </span>
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {sectionQuestions.map((mainQuestion, index) => (
                                  <tr
                                    key={mainQuestion.id}
                                    className={`border-b border-blue-200 dark:border-blue-800 ${index % 2 === 0
                                      ? "bg-white dark:bg-gray-800/50"
                                      : "bg-blue-100/30 dark:bg-blue-900/10"
                                      }`}
                                  >
                                    <td className="px-6 py-4 font-medium text-gray-800 dark:text-gray-200 border border-blue-200 dark:border-blue-800 bg-gradient-to-br from-blue-50 to-blue-100 dark:from-blue-900/30 dark:to-blue-800/30">
                                      <div className="flex flex-col gap-3">
                                        {mainQuestion.subParam1 && (
                                          <span
                                            className="inline-block px-3 py-1 bg-blue-100/60 dark:bg-blue-900/30 text-blue-900 dark:text-blue-200 text-xs font-bold rounded-md w-fit"
                                            title={mainQuestion.subParam1}
                                          >
                                            {mainQuestion.subParam1}
                                          </span>
                                        )}
                                        <div
                                          className="text-xs text-gray-600 dark:text-gray-400"
                                          title={mainQuestion.title}
                                        >
                                          {mainQuestion.title}
                                        </div>
                                      </div>
                                    </td>
                                    {uniqueSubParams.map((subParam) => {
                                      const followUpsForParam =
                                        followUpsBySubParam.get(subParam) || [];
                                      const answerQuestionPairs = followUpsForParam
                                        .map((followUp) => {
                                          const followUpFromMain =
                                            mainQuestion.followUpQuestions.find(
                                              (fq: any) => fq.id === followUp.id
                                            );
                                          const answer = followUpFromMain?.answer;

                                          const isNotEmpty = answer !== undefined && answer !== null && answer !== "" &&
                                            (!Array.isArray(answer) || answer.length > 0) &&
                                            (typeof answer !== "object" || Object.keys(answer).length > 0);

                                          const isNA = typeof answer === "string" && (
                                            answer === "N/A" ||
                                            answer === "n/a" ||
                                            answer.toLowerCase() === complianceLabels.na.toLowerCase()
                                          );

                                          if (isNotEmpty && !isNA) {
                                            return {
                                              answer,
                                              question: followUpFromMain,
                                            };
                                          }
                                          return null;
                                        })
                                        .filter((item) => item !== null);

                                      return (
                                        <td
                                          key={subParam}
                                          className="px-4 py-4 border border-blue-200 dark:border-blue-800 text-sm text-gray-700 dark:text-gray-300 bg-blue-50/40 dark:bg-blue-900/20"
                                        >
                                          {answerQuestionPairs.length > 0 ? (
                                            <div className="flex flex-wrap gap-2">
                                              {answerQuestionPairs.map((item: any, idx) => {
                                                const renderValue = (val: any): React.ReactNode => {
                                                  if (val === null || val === undefined || val === "") return null;

                                                  if (Array.isArray(val)) {
                                                    return (
                                                      <div className="flex flex-col gap-1">
                                                        {val.map((v, i) => (
                                                          <div key={i}>{renderValue(v)}</div>
                                                        ))}
                                                      </div>
                                                    );
                                                  }

                                                  if (typeof val === 'object') {
                                                    if (val.url && isImageUrl(String(val.url))) {
                                                      return (
                                                        <ImageLink
                                                          text={String(val.url)}
                                                          showImage={showMainParamsImages[section.id] ?? false}
                                                        />
                                                      );
                                                    }
                                                    if (val.answer && isImageUrl(String(val.answer))) {
                                                      return (
                                                        <ImageLink
                                                          text={String(val.answer)}
                                                          showImage={showMainParamsImages[section.id] ?? false}
                                                        />
                                                      );
                                                    }

                                                    const entries = Object.entries(val);
                                                    if (entries.length > 0) {
                                                      return (
                                                        <div className="flex flex-col gap-1">
                                                          {entries.map(([k, v], i) => (
                                                            <div key={i} className="flex flex-col gap-0.5">
                                                              <span className="text-[10px] font-bold opacity-70 uppercase tracking-tighter text-blue-800 dark:text-blue-300">
                                                                {k}
                                                              </span>
                                                              {renderValue(v)}
                                                            </div>
                                                          ))}
                                                        </div>
                                                      );
                                                    }
                                                    return JSON.stringify(val);
                                                  }

                                                  const textValue = String(val);
                                                  if (isImageUrl(textValue)) {
                                                    return (
                                                      <ImageLink
                                                        text={textValue}
                                                        showImage={showMainParamsImages[section.id] ?? false}
                                                      />
                                                    );
                                                  }

                                                  return textValue;
                                                };

                                                return (
                                                  <div key={idx} className="w-full">
                                                    <div className="font-medium text-gray-800 dark:text-gray-200">
                                                      {renderValue(item.answer)}
                                                    </div>
                                                    {response?.responseRanks?.[item.question?.id] && (
                                                      <div className={`text-[10px] font-bold min-w-[24px] h-6 px-1.5 rounded-full flex items-center justify-center border shadow-sm w-fit mt-1 ${getRankStyle(item.answer, document.documentElement.classList.contains("dark"))}`}>
                                                        #{response.responseRanks[item.question.id]}
                                                      </div>
                                                    )}
                                                  </div>
                                                );
                                              })}
                                            </div>
                                          ) : (
                                            <span className="text-gray-400 italic">
                                              {complianceLabels.na}
                                            </span>
                                          )}
                                        </td>
                                      );
                                    })}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                );
              })}



              {/* Response Summary Card */}
              <div className="bg-white dark:bg-gray-900 rounded-2xl border border-gray-200 dark:border-gray-700 p-6">
                <h3 className="text-lg font-bold text-gray-900 dark:text-white mb-6 flex items-center gap-2">
                  <Users className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                  Response Summary
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-blue-100 dark:bg-blue-900/30 rounded-lg mt-1">
                      <FileText className="w-4 h-4 text-blue-600 dark:text-blue-400" />
                    </div>
                    <div>
                      <p className="text-xs font-semibold text-gray-500 dark:text-gray-400 uppercase">
                        Form Name
                      </p>
                      <p className="text-sm font-semibold text-gray-900 dark:text-white mt-1">
                        {form.title}
                      </p>
                    </div>
                  </div>
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-green-100 dark:bg-green-900/30 rounded-lg mt-1">
                      <Calendar className="w-4 h-4 text-green-600 dark:text-green-400" />
                    </div>
                    <div>
                      <p className="text-xs font-semibold text-gray-500 dark:text-gray-400 uppercase">
                        Submission Date
                      </p>
                      <p className="text-sm font-semibold text-gray-900 dark:text-white mt-1">
                        {formatTimestamp(response.createdAt)}
                      </p>
                    </div>
                  </div>
                  {response.submissionMetadata?.location && (
                    <div className="flex items-start gap-3">
                      <div className="p-2 bg-purple-100 dark:bg-purple-900/30 rounded-lg mt-1">
                        <MapPin className="w-4 h-4 text-purple-600 dark:text-purple-400" />
                      </div>
                      <div>
                        <p className="text-xs font-semibold text-gray-500 dark:text-gray-400 uppercase">
                          Location
                        </p>
                        <p className="text-sm font-semibold text-gray-900 dark:text-white mt-1">
                          {response.submissionMetadata.location.city},{" "}
                          {response.submissionMetadata.location.country}
                        </p>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>
          ) : (
            <div className="bg-white dark:bg-gray-900 rounded-2xl border border-gray-200 dark:border-gray-700 p-8 text-center">
              <p className="text-gray-600 dark:text-gray-400 text-lg">
                No section data available for analysis
              </p>
            </div>
          )
        ) : viewMode === "ops" ? (
          <div ref={opsPrintRef} data-ops-template="true">
            <OPSTemplate
              form={form}
              response={response}
              submissionHistory={submissionHistory}
              sameFormatResponses={sameFormatResponses}
              isLoadingSameFormatResponses={isLoadingSameFormatResponses}
              formatQuestionId={instructionsQuestions?.[0]?.id}
              controlQuestionId={instructionsQuestions?.[1]?.id}

            />
          </div>
        ) : (
          <div className="bg-white dark:bg-gray-900 rounded-2xl border border-gray-200 dark:border-gray-700 p-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
              <h3 className="text-xl font-bold text-gray-900 dark:text-white">
                Form Responses
              </h3>
              <div className="flex flex-wrap items-center gap-3">
                <button
                  onClick={() => handleDownloadPDF('responses-view')}
                  disabled={generatingPDF || exportingZip}
                  className="flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-bold text-white rounded-xl transition-all duration-300 hover:shadow-lg hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed"
                  style={{ backgroundColor: "#0891b2" }}
                >
                  {generatingPDF ? (
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  ) : (
                    <FileText className="w-4 h-4" />
                  )}
                  <span>{generatingPDF ? "Generating PDF..." : "Download as PDF"}</span>
                </button>
                <button
                  onClick={handleBulkDownloadZip}
                  disabled={generatingPDF || exportingZip}
                  className="flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-bold text-white rounded-xl transition-all duration-300 hover:shadow-lg hover:scale-[1.02] active:scale-[0.98] disabled:opacity-50 disabled:cursor-not-allowed"
                  style={{ backgroundColor: "#0e7490" }}
                >
                  {exportingZip ? (
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  ) : (
                    <Download className="w-4 h-4" />
                  )}
                  <span>{exportingZip ? "Preparing ZIP..." : "Bulk Download (ZIP)"}</span>
                </button>
              </div>
            </div>

            {/* Structured Table View */}
            <div className="overflow-x-auto">
              {(() => {
                const allRows: Array<{
                  id: string;
                  questionText: string;
                  subParam1?: string;
                  subParam2?: string;
                  answer: any;
                  sectionTitle: string;
                }> = [];

                const answeredKeys = new Set<string>();

                // 1. Add Submission Metadata first
                if (response?.submissionMetadata) {
                  const meta = response.submissionMetadata;
                  if (meta.ipAddress) allRows.push({ id: 'meta-ip', questionText: 'IP Address', answer: meta.ipAddress, sectionTitle: 'Submission Metadata' });
                  if (meta.submittedAt || response.createdAt) allRows.push({ id: 'meta-time', questionText: 'Submission Time', answer: formatTimestamp(meta.submittedAt || response.createdAt), sectionTitle: 'Submission Metadata' });
                  if (meta.browser) allRows.push({ id: 'meta-browser', questionText: 'Browser', answer: meta.browser, sectionTitle: 'Submission Metadata' });
                  if (meta.os) allRows.push({ id: 'meta-os', questionText: 'Operating System', answer: meta.os, sectionTitle: 'Submission Metadata' });
                  if (meta.device) allRows.push({ id: 'meta-device', questionText: 'Device', answer: meta.device, sectionTitle: 'Submission Metadata' });
                  if (meta.location) {
                    const loc = meta.location;
                    const locStr = [loc.city, loc.region, loc.country].filter(Boolean).join(', ');
                    if (locStr) allRows.push({ id: 'meta-location', questionText: 'IP Location', answer: locStr, sectionTitle: 'Submission Metadata' });
                  }
                }

                // 2. Add specific metadata fields
                if (response?.dealerName) allRows.push({ id: 'meta-dealer', questionText: 'Dealer Name', answer: response.dealerName, sectionTitle: 'Submission Metadata' });
                if (response?.answers?.auditorName) {
                  allRows.push({ id: 'meta-auditor', questionText: 'Auditor Name', answer: response.answers.auditorName, sectionTitle: 'Submission Metadata' });
                  answeredKeys.add('auditorName');
                }

                // 3. Add form questions
                form?.sections?.forEach((section: any) => {
                  section.questions?.forEach((question: any) => {
                    const answer = response?.answers?.[question.id];
                    if (answer !== undefined && answer !== null && answer !== '') {
                      answeredKeys.add(question.id);
                      allRows.push({
                        id: question.id,
                        questionText: question.text || question.label || question.id,
                        subParam1: question.subParam1,
                        subParam2: question.subParam2,
                        answer,
                        sectionTitle: section.title || "Untitled Section"
                      });
                    }
                    question.followUpQuestions?.forEach((followUp: any) => {
                      const followUpAnswer = response?.answers?.[followUp.id];
                      if (followUpAnswer !== undefined && followUpAnswer !== null && followUpAnswer !== '') {
                        answeredKeys.add(followUp.id);
                        allRows.push({
                          id: followUp.id,
                          questionText: followUp.text || followUp.label || followUp.id,
                          subParam1: followUp.subParam1,
                          subParam2: followUp.subParam2,
                          answer: followUpAnswer,
                          sectionTitle: section.title || "Untitled Section"
                        });
                      }
                    });
                  });
                });

                form?.followUpQuestions?.forEach((followUp: any) => {
                  const followUpAnswer = response?.answers?.[followUp.id];
                  if (followUpAnswer !== undefined && followUpAnswer !== null && followUpAnswer !== '') {
                    answeredKeys.add(followUp.id);
                    allRows.push({
                      id: followUp.id,
                      questionText: followUp.text || followUp.label || followUp.id,
                      subParam1: followUp.subParam1,
                      subParam2: followUp.subParam2,
                      answer: followUpAnswer,
                      sectionTitle: "Follow-up Questions"
                    });
                  }
                });

                // 4. Add any other unmapped answers
                Object.entries(response?.answers || {}).forEach(([key, value]) => {
                  if (!answeredKeys.has(key) && !key.startsWith('synthetic_') && value !== undefined && value !== null && value !== '') {
                    allRows.push({
                      id: key,
                      questionText: key,
                      answer: value,
                      sectionTitle: "Additional Data"
                    });
                  }
                });

                const isZoneStructured = (val: any): boolean => {
                  if (!val || typeof val !== 'object') return false;
                  // Only treat as structured if it has hierarchical zone data or explicit defects
                  return ('zonesData' in val && Object.keys(val.zonesData || {}).length > 0) ||
                    ('defects' in val && Array.isArray(val.defects) && val.defects.length > 0);
                };

                const renderSimpleAnswer = (val: any): React.ReactNode => {
                  if (val === null || val === undefined || val === '') return <span className="text-gray-400 italic text-xs">-</span>;
                  if (typeof val === 'string') {
                    if (isImageUrl(val)) return <ImageLink text={val} />;
                    return <span className="text-sm">{val}</span>;
                  }
                  if (Array.isArray(val)) {
                    return (
                      <div className="flex flex-col gap-1">
                        {val.map((v, i) => <div key={i}>{renderSimpleAnswer(v)}</div>)}
                      </div>
                    );
                  }
                  if (typeof val === 'object') {
                    if (val.url && isImageUrl(String(val.url))) return <ImageLink text={val.url} />;
                    if (val.answer !== undefined) return renderSimpleAnswer(val.answer);

                    // Special handling for chassis/zone types in simple view
                    const isChassisType = 'chassisNumber' in val || 'status' in val || 'zone' in val || 'zones' in val;
                    if (isChassisType) {
                      const parts = [];
                      if (val.status) parts.push(<div key="status" className="flex gap-1"><span className="font-bold text-orange-600 dark:text-orange-400 text-[10px] uppercase">Status:</span><span className="text-xs font-semibold">{val.status}</span></div>);
                      if (val.chassisNumber) parts.push(<div key="chassis" className="flex gap-1"><span className="font-bold text-blue-600 dark:text-blue-400 text-[10px] uppercase">Chassis:</span><span className="text-xs font-semibold">{val.chassisNumber}</span></div>);
                      const zones = val.zone || val.zones;
                      if (zones) {
                        const zonesStr = Array.isArray(zones) ? zones.join(', ') : zones;
                        if (zonesStr) parts.push(<div key="zones" className="flex gap-1"><span className="font-bold text-indigo-600 dark:text-indigo-400 text-[10px] uppercase">Zones:</span><span className="text-xs font-semibold">{zonesStr}</span></div>);
                      }
                      if (val.remark) parts.push(<div key="remark" className="flex gap-1 mt-0.5"><span className="font-bold text-gray-500 text-[10px] uppercase">Remark:</span><span className="text-xs italic text-gray-600 dark:text-gray-400">📝 {val.remark}</span></div>);
                      if (val.evidenceUrl) {
                        parts.push(
                          <div key="evidence" className="mt-1">
                            {isImageUrl(val.evidenceUrl) ? <ImageLink text={val.evidenceUrl} /> : <a href={val.evidenceUrl} target="_blank" rel="noopener noreferrer" className="text-[10px] text-blue-600 underline font-bold uppercase tracking-tighter">View Evidence</a>}
                          </div>
                        );
                      }

                      return <div className="flex flex-col gap-0.5 py-1">{parts}</div>;
                    }

                    return (
                      <div className="flex flex-col gap-1">
                        {Object.entries(val).map(([k, v], i) => (
                          <div key={i} className="flex gap-1">
                            <span className="font-semibold text-gray-500 text-xs">{k}:</span>
                            <span className="text-xs">{renderSimpleAnswer(v)}</span>
                          </div>
                        ))}
                      </div>
                    );
                  }
                  return <span className="text-sm">{String(val)}</span>;
                };

                allRows.sort((a, b) => {
                  if (a.sectionTitle !== b.sectionTitle) return a.sectionTitle.localeCompare(b.sectionTitle);
                  return a.id.localeCompare(b.id);
                });

                const isChassisType = (val: any): boolean =>
                  val && typeof val === 'object' && ('chassisNumber' in val || 'status' in val || 'zone' in val || 'zones' in val || 'zonesData' in val);

                // Check if ANY row has zone/chassis data
                const hasAnyZoneData = allRows.some(r => isChassisType(r.answer));

                // Check if Category and Defects columns are actually needed
                const showCategoryCol = allRows.some(r => {
                  const val = r.answer;
                  if (!val || typeof val !== 'object') return false;
                  if (val.zonesData) {
                    return Object.values(val.zonesData).some((z: any) => z.categories && z.categories.length > 0);
                  }
                  return val.category || val.categoryName;
                });

                const showDefectsCol = allRows.some(r => {
                  const val = r.answer;
                  if (!val || typeof val !== 'object') return false;
                  if (val.zonesData) {
                    return Object.values(val.zonesData).some((z: any) =>
                      z.categories?.some((c: any) => c.defects && c.defects.length > 0)
                    );
                  }
                  return (val.defects && val.defects.length > 0) || val.defect || val.defects;
                });

                const columnCount = 4 + (hasAnyZoneData ? (5 + (showCategoryCol ? 1 : 0) + (showDefectsCol ? 1 : 0)) : 1) + 1;

                const rows: JSX.Element[] = [];
                let currentSection = '';
                let rowIndex = 0;

                allRows.forEach((row) => {
                  if (currentSection !== row.sectionTitle) {
                    currentSection = row.sectionTitle;
                    rows.push(
                      <tr key={`section-${currentSection}`} className="bg-gradient-to-r from-blue-100 to-indigo-100 dark:from-blue-900/40 dark:to-indigo-900/40">
                        <td
                          colSpan={columnCount}
                          className="px-4 py-2 text-[10px] font-bold text-blue-900 dark:text-blue-200 border border-gray-200 dark:border-gray-600 uppercase tracking-widest"
                        >
                          Section: {currentSection}
                        </td>
                      </tr>
                    );
                  }

                  const isEven = rowIndex % 2 === 0;
                  rowIndex++;

                  if (isZoneStructured(row.answer)) {
                    const val = row.answer;
                    const status = val?.status || '-';
                    const zones = val?.zone || val?.zones;
                    const zonesData = val?.zonesData || {};

                    type ZoneEntry = {
                      zone: string;
                      category: string;
                      defects: string;
                      remark: string;
                      file: string;
                      zoneRowSpan?: number;
                      categoryRowSpan?: number;
                      showZone?: boolean;
                      showCategory?: boolean;
                    };

                    const zoneEntries: ZoneEntry[] = [];

                    if (zonesData && typeof zonesData === 'object') {
                      Object.entries(zonesData).forEach(([zoneName, zoneVal]: [string, any]) => {
                        const categories = zoneVal?.categories;
                        const zoneStartIndex = zoneEntries.length;

                        if (Array.isArray(categories) && categories.length > 0) {
                          categories.forEach((cat: any) => {
                            const catName = typeof cat === 'string' ? cat : (cat?.name || cat?.category || '-');
                            const defects = cat?.defects;
                            const catStartIndex = zoneEntries.length;

                            if (Array.isArray(defects) && defects.length > 0) {
                              defects.forEach((defect: any) => {
                                const defectName = typeof defect === 'string'
                                  ? defect
                                  : (defect?.name || defect?.defect || defect?.title || '-');
                                const details = typeof defect === 'object' ? (defect?.details || {}) : {};
                                const remark = typeof defect === 'object'
                                  ? (details?.remark || details?.remarks || defect?.remark || defect?.remarks || defect?.comment || '-')
                                  : '-';
                                const file = typeof defect === 'object'
                                  ? (details?.fileUrl || details?.file || defect?.fileUrl || defect?.file || defect?.imageUrl || defect?.photo || defect?.image || defect?.url || '')
                                  : '';
                                zoneEntries.push({
                                  zone: zoneName,
                                  category: catName,
                                  defects: defectName,
                                  remark: remark || '-',
                                  file: file || '',
                                  showZone: false,
                                  showCategory: false,
                                });
                              });
                            } else {
                              const details = typeof cat === 'object' ? (cat?.details || {}) : {};
                              const remark = typeof cat === 'object' ? (details?.remark || cat?.remark || '-') : '-';
                              const file = typeof cat === 'object' ? (details?.fileUrl || details?.file || cat?.fileUrl || cat?.file || '') : '';
                              zoneEntries.push({
                                zone: zoneName,
                                category: catName,
                                defects: '-',
                                remark: remark || '-',
                                file: file || '',
                                showZone: false,
                                showCategory: false,
                              });
                            }

                            const catCount = zoneEntries.length - catStartIndex;
                            if (catCount > 0) {
                              zoneEntries[catStartIndex].showCategory = true;
                              zoneEntries[catStartIndex].categoryRowSpan = catCount;
                            }
                          });
                        } else {
                          zoneEntries.push({
                            zone: zoneName,
                            category: '-',
                            defects: '-',
                            remark: '-',
                            file: '',
                            showZone: false,
                            showCategory: true,
                            categoryRowSpan: 1,
                          });
                        }

                        const zoneCount = zoneEntries.length - zoneStartIndex;
                        if (zoneCount > 0) {
                          zoneEntries[zoneStartIndex].showZone = true;
                          zoneEntries[zoneStartIndex].zoneRowSpan = zoneCount;
                        }
                      });
                    }

                    if (zoneEntries.length === 0) {
                      const zonesStr = Array.isArray(zones) ? zones.join(', ') : (zones || '-');
                      const defectsStr = Array.isArray(val?.defects)
                        ? val.defects.map((d: any) => typeof d === 'string' ? d : (d?.name || d?.defect || '-')).join(', ')
                        : (val?.defect || val?.defects || '-');

                      zoneEntries.push({
                        zone: zonesStr,
                        category: val?.category || val?.categoryName || '-',
                        defects: defectsStr,
                        remark: val?.remark || val?.remarks || val?.comment || val?.comments || '-',
                        file: val?.evidenceUrl || val?.fileUrl || val?.file || val?.imageUrl || '',
                        showZone: true,
                        showCategory: true,
                        zoneRowSpan: 1,
                        categoryRowSpan: 1,
                      });
                    }

                    const totalRowSpan = zoneEntries.length;
                    const entry0 = zoneEntries[0];

                    rows.push(
                      <tr key={`${row.id}-0`} className={`${isEven ? 'bg-white dark:bg-gray-900' : 'bg-gray-50/50 dark:bg-gray-800/30'} hover:bg-blue-50/50 dark:hover:bg-blue-900/20 transition-colors`}>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-semibold text-gray-600 dark:text-gray-400 align-top">
                          {row.sectionTitle}
                        </td>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-medium text-gray-800 dark:text-gray-200 align-top max-w-[200px]">
                          {row.questionText}
                        </td>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-[10px] text-blue-700 dark:text-blue-300 align-top">
                          {row.subParam1 || '-'}
                        </td>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-[10px] text-emerald-700 dark:text-emerald-300 align-top">
                          {row.subParam2 || '-'}
                        </td>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 align-middle text-center w-[90px]">
                          <span className="inline-block bg-orange-100 dark:bg-orange-900/30 text-orange-800 dark:text-orange-300 px-2 py-1 rounded-full text-[10px] font-bold border border-orange-200 dark:border-orange-700 whitespace-nowrap">
                            {status}
                          </span>
                        </td>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 align-middle w-[100px]">
                          <div className="flex flex-col gap-1">
                            {(Array.isArray(zones) ? zones : [zones]).filter(Boolean).map((z: string, i: number) => (
                              <span key={i} className="inline-block bg-blue-100 dark:bg-blue-900/30 text-blue-800 dark:text-blue-300 px-1.5 py-0.5 rounded text-[10px] font-semibold border border-blue-200 dark:border-blue-700 whitespace-nowrap">
                                {z}
                              </span>
                            ))}
                          </div>
                        </td>
                        <td rowSpan={entry0.zoneRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-semibold text-gray-800 dark:text-gray-200 align-middle bg-blue-50/30 dark:bg-blue-900/10">
                          {entry0.zone}
                        </td>
                        {entry0.showCategory && showCategoryCol && (
                          <td rowSpan={entry0.categoryRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-700 dark:text-gray-300 align-middle">
                            {entry0.category}
                          </td>
                        )}
                        {showDefectsCol && (
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-700 dark:text-gray-300">
                            {entry0.defects}
                          </td>
                        )}
                        <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-600 dark:text-gray-400 italic">
                          {entry0.remark !== '-' ? <span className="flex items-center gap-1">📝 {entry0.remark}</span> : '-'}
                        </td>
                        <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs">
                          {entry0.file ? (
                            isImageUrl(entry0.file)
                              ? <ImageLink text={entry0.file} />
                              : <a href={entry0.file} target={"_blank"} rel={"noopener noreferrer"} className="text-blue-600 dark:text-blue-400 underline text-[10px] font-bold">View</a>
                          ) : '-'}
                        </td>
                        <td rowSpan={totalRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-bold text-center align-middle">
                          {response?.responseRanks?.[row.id] ? `#${response.responseRanks[row.id]}` : '-'}
                        </td>
                      </tr>
                    );

                    zoneEntries.slice(1).forEach((entry, idx) => {
                      rows.push(
                        <tr key={`${row.id}-${idx + 1}`} className={`${isEven ? 'bg-white dark:bg-gray-900' : 'bg-gray-50/50 dark:bg-gray-800/30'} hover:bg-blue-50/50 dark:hover:bg-blue-900/20 transition-colors`}>
                          {entry.showZone && (
                            <td rowSpan={entry.zoneRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-semibold text-gray-800 dark:text-gray-200 align-middle bg-blue-50/30 dark:bg-blue-900/10">
                              {entry.zone}
                            </td>
                          )}
                          {entry.showCategory && showCategoryCol && (
                            <td rowSpan={entry.categoryRowSpan} className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-700 dark:text-gray-300 align-middle">
                              {entry.category}
                            </td>
                          )}
                          {showDefectsCol && (
                            <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-700 dark:text-gray-300">
                              {entry.defects}
                            </td>
                          )}
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-600 dark:text-gray-400 italic">
                            {entry.remark !== '-' ? <span className="flex items-center gap-1">📝 {entry.remark}</span> : '-'}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs">
                            {entry.file ? (
                              isImageUrl(entry.file)
                                ? <ImageLink text={entry.file} />
                                : <a href={entry.file} target={"_blank"} rel={"noopener noreferrer"} className="text-blue-600 dark:text-blue-400 underline text-[10px] font-bold">View</a>
                            ) : '-'}
                          </td>
                        </tr>
                      );
                    });
                  } else {
                    const val = row.answer;
                    const isChassisRow = isChassisType(val);

                    if (hasAnyZoneData && isChassisRow) {
                      const status = val.status || '-';
                      const zones = val.zone || val.zones;
                      const chassis = val.chassisNumber || '-';
                      const remark = val.remark || '-';
                      const file = val.evidenceUrl || '';

                      rows.push(
                        <tr key={row.id} className={`${isEven ? 'bg-white dark:bg-gray-900' : 'bg-gray-50/50 dark:bg-gray-800/30'} hover:bg-blue-50/50 dark:hover:bg-blue-900/20 transition-colors`}>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-semibold text-gray-600 dark:text-gray-400 align-top">
                            {row.sectionTitle}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-medium text-gray-800 dark:text-gray-200 align-top max-w-[200px]">
                            {row.questionText}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-[10px] text-blue-700 dark:text-blue-300 align-top">
                            {row.subParam1 || '-'}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-[10px] text-emerald-700 dark:text-emerald-300 align-top">
                            {row.subParam2 || '-'}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 align-middle text-center w-[90px]">
                            <span className="inline-block bg-orange-100 dark:bg-orange-900/30 text-orange-800 dark:text-orange-300 px-2 py-1 rounded-full text-[10px] font-bold border border-orange-200 dark:border-orange-700 whitespace-nowrap">
                              {status}
                            </span>
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 align-middle w-[100px]">
                            <div className="flex flex-wrap gap-1">
                              {(Array.isArray(zones) ? zones : [zones]).filter(Boolean).map((z: string, i: number) => (
                                <span key={i} className="inline-block bg-blue-100 dark:bg-blue-900/30 text-blue-800 dark:text-blue-300 px-1.5 py-0.5 rounded text-[10px] font-semibold border border-blue-200 dark:border-blue-700 whitespace-nowrap">
                                  {z}
                                </span>
                              ))}
                            </div>
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-semibold text-gray-800 dark:text-gray-200 align-middle bg-blue-50/30 dark:bg-blue-900/10">
                            {chassis}
                          </td>
                          {showCategoryCol && (
                            <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-700 dark:text-gray-300 align-middle text-center">
                              -
                            </td>
                          )}
                          {showDefectsCol && (
                            <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-700 dark:text-gray-300 text-center">
                              -
                            </td>
                          )}
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-gray-600 dark:text-gray-400 italic">
                            {remark !== '-' ? <span className="flex items-center gap-1">📝 {remark}</span> : '-'}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs text-center">
                            {file ? (
                              isImageUrl(file)
                                ? <ImageLink text={file} />
                                : <a href={file} target={"_blank"} rel={"noopener noreferrer"} className="text-blue-600 dark:text-blue-400 underline font-bold uppercase tracking-tighter text-[10px]">View</a>
                            ) : '-'}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-bold text-center align-top">
                            {response?.responseRanks?.[row.id] ? `#${response.responseRanks[row.id]}` : '-'}
                          </td>
                        </tr>
                      );
                    } else {
                      // Normal non-zone row
                      rows.push(
                        <tr key={row.id} className={`${isEven ? 'bg-white dark:bg-gray-900' : 'bg-gray-50/50 dark:bg-gray-800/30'} hover:bg-blue-50/50 dark:hover:bg-blue-900/20 transition-colors`}>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-semibold text-gray-600 dark:text-gray-400 align-top">
                            {row.sectionTitle}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-medium text-gray-800 dark:text-gray-200 align-top max-w-[200px]">
                            {row.questionText}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-[10px] text-blue-700 dark:text-blue-300 align-top">
                            {row.subParam1 || '-'}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-[10px] text-emerald-700 dark:text-emerald-300 align-top">
                            {row.subParam2 || '-'}
                          </td>
                          <td colSpan={hasAnyZoneData ? (5 + (showCategoryCol ? 1 : 0) + (showDefectsCol ? 1 : 0)) : 1} className="px-3 py-2 border border-gray-200 dark:border-gray-700 align-top">
                            {renderSimpleAnswer(row.answer)}
                          </td>
                          <td className="px-3 py-2 border border-gray-200 dark:border-gray-700 text-xs font-bold text-center align-top">
                            {response?.responseRanks?.[row.id] ? `#${response.responseRanks[row.id]}` : '-'}
                          </td>
                        </tr>
                      );
                    }
                  }
                });

                if (rows.length === 0) {
                  return (
                    <table className="w-full border-collapse">
                      <tbody>
                        <tr>
                          <td colSpan={8} className="px-4 py-8 text-center text-gray-500 dark:text-gray-400">
                            No responses found
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  );
                }

                return (
                  <table className="w-full border-collapse text-sm">
                    <thead className="sticky top-0 z-10">
                      <tr className="bg-gradient-to-r from-gray-100 to-gray-200 dark:from-gray-800 dark:to-gray-700">
                        <th className="px-3 py-3 text-left text-[10px] font-bold text-gray-700 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                          Section
                        </th>
                        <th className="px-3 py-3 text-left text-[10px] font-bold text-gray-700 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600 min-w-[150px]">
                          Question
                        </th>
                        <th className="px-3 py-3 text-left text-[10px] font-bold text-blue-700 dark:text-blue-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                          Main Parameter
                        </th>
                        <th className="px-3 py-3 text-left text-[10px] font-bold text-emerald-700 dark:text-emerald-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                          Sub Parameter
                        </th>

                        {hasAnyZoneData ? (
                          <>
                            <th className="px-3 py-3 text-left text-[10px] font-bold text-orange-700 dark:text-orange-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                              Status
                            </th>
                            <th className="px-3 py-3 text-left text-[10px] font-bold text-blue-700 dark:text-blue-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                              Zones
                            </th>
                            <th className="px-3 py-3 text-left text-[10px] font-bold text-indigo-700 dark:text-indigo-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                              Zone / Chassis
                            </th>
                            {showCategoryCol && (
                              <th className="px-3 py-3 text-left text-[10px] font-bold text-indigo-700 dark:text-indigo-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                                Category
                              </th>
                            )}
                            {showDefectsCol && (
                              <th className="px-3 py-3 text-left text-[10px] font-bold text-indigo-700 dark:text-indigo-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                                Defects
                              </th>
                            )}
                            <th className="px-3 py-3 text-left text-[10px] font-bold text-indigo-700 dark:text-indigo-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                              Remark
                            </th>
                            <th className="px-3 py-3 text-left text-[10px] font-bold text-indigo-700 dark:text-indigo-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                              File
                            </th>
                          </>
                        ) : (
                          <th className="px-3 py-3 text-left text-[10px] font-bold text-gray-700 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                            Answer / Response
                          </th>
                        )}
                        <th className="px-3 py-3 text-left text-[10px] font-bold text-gray-700 dark:text-gray-300 uppercase tracking-wider border border-gray-200 dark:border-gray-600">
                          Rank
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows}
                    </tbody>
                  </table>
                );
              })()}
            </div>
          </div>
        )}
      </div>

      {editingResponse && editingFormLoading && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-white dark:bg-gray-900 rounded-xl shadow-2xl px-6 py-4 flex items-center gap-3 border border-gray-200 dark:border-gray-700">
            <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-primary-600"></div>
            <div className="text-primary-600 dark:text-primary-400 font-semibold">
              Loading form details...
            </div>
          </div>
        </div>
      )}

      {editingResponse && editingForm && !editingFormLoading && (
        <ResponseEdit
          response={editingResponse as any}
          question={editingForm as any}
          onSave={handleSaveEditedResponse}
          onCancel={handleCloseEdit}
        />
      )}

      {showSectionsPDFModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white dark:bg-gray-900 rounded-xl shadow-2xl w-full max-w-6xl max-h-[90vh] overflow-auto border border-gray-200 dark:border-gray-700">
            <div className="sticky top-0 z-40 bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 p-4 flex items-center justify-between">
              <h2 className="text-lg font-bold text-gray-900 dark:text-white flex items-center gap-2">
                <BarChart3 className="w-5 h-5" />
                View Sections - PDF Preview
              </h2>
              <button
                onClick={() => setShowSectionsPDFModal(false)}
                className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors"
              >
                <X className="w-5 h-5 text-gray-600 dark:text-gray-400" />
              </button>
            </div>

            <div className="p-6 space-y-6">
              {(() => {
                const totalYes = filteredSectionStats.reduce((sum, stat) => sum + stat.yes, 0);
                const totalNo = filteredSectionStats.reduce((sum, stat) => sum + stat.no, 0);
                const totalNA = filteredSectionStats.reduce((sum, stat) => sum + stat.na, 0);
                const totalCorrect = filteredSectionStats.reduce((sum, stat) => sum + stat.correct, 0);
                const totalWrong = filteredSectionStats.reduce((sum, stat) => sum + stat.wrong, 0);
                const totalQuestions = filteredSectionStats.reduce((sum, stat) => sum + stat.total, 0);

                const yesPercent = totalQuestions > 0 ? ((totalYes / totalQuestions) * 100).toFixed(1) : "0.0";
                const noPercent = totalQuestions > 0 ? ((totalNo / totalQuestions) * 100).toFixed(1) : "0.0";
                const naPercent = totalQuestions > 0 ? ((totalNA / totalQuestions) * 100).toFixed(1) : "0.0";
                const correctPercent = totalQuestions > 0 ? ((totalCorrect / totalQuestions) * 100).toFixed(1) : "0.0";
                const wrongPercent = totalQuestions > 0 ? ((totalWrong / totalQuestions) * 100).toFixed(1) : "0.0";

                const radarData = {
                  labels: filteredSectionStats.map((stat) => stat.title),
                  datasets: [
                    {
                      label: `Correct ${correctPercent}% (${totalCorrect})`,
                      data: filteredSectionStats.map((stat) =>
                        stat.total ? ((stat.correct / stat.total) * 100).toFixed(1) : 0
                      ),
                      borderColor: "#059669",
                      backgroundColor: "rgba(5, 150, 105, 0.15)",
                      borderWidth: 3,
                      pointBackgroundColor: "#059669",
                      pointBorderColor: "#fff",
                      pointHoverRadius: 6,
                    },
                    {
                      label: `Wrong ${wrongPercent}% (${totalWrong})`,
                      data: filteredSectionStats.map((stat) =>
                        stat.total ? ((stat.wrong / stat.total) * 100).toFixed(1) : 0
                      ),
                      borderColor: "#dc2626",
                      backgroundColor: "rgba(220, 38, 38, 0.15)",
                      borderWidth: 3,
                      pointBackgroundColor: "#dc2626",
                      pointBorderColor: "#fff",
                      pointHoverRadius: 6,
                    },
                    {
                      label: `${complianceLabels.yes} / Answered ${yesPercent}% (${totalYes})`,
                      data: filteredSectionStats.map((stat) =>
                        stat.total ? ((stat.yes / stat.total) * 100).toFixed(1) : 0
                      ),
                      borderColor: "#10b981",
                      backgroundColor: "rgba(16, 185, 129, 0.05)",
                      borderWidth: 2,
                      hidden: true,
                    },
                    {
                      label: `${complianceLabels.no} / Not Answered ${noPercent}% (${totalNo})`,
                      data: filteredSectionStats.map((stat) =>
                        stat.total ? ((stat.no / stat.total) * 100).toFixed(1) : 0
                      ),
                      borderColor: "#ef4444",
                      backgroundColor: "rgba(239, 68, 68, 0.05)",
                      borderWidth: 2,
                      hidden: true,
                    },
                    {
                      label: `${complianceLabels.na} ${naPercent}% (${totalNA})`,
                      data: filteredSectionStats.map((stat) =>
                        stat.total ? ((stat.na / stat.total) * 100).toFixed(1) : 0
                      ),
                      borderColor: "#f59e0b",
                      backgroundColor: "rgba(245, 158, 11, 0.05)",
                      borderWidth: 2,
                      hidden: true,
                    },
                  ],
                };

                const radarOptions = {
                  responsive: true,
                  maintainAspectRatio: false,
                  plugins: {
                    legend: {
                      position: "top" as const,
                      labels: {
                        color: document.documentElement.classList.contains("dark")
                          ? "#d1d5db"
                          : "#374151",
                      },
                    },
                  },
                  scales: {
                    r: {
                      beginAtZero: true,
                      max: 100,
                      ticks: {
                        color: document.documentElement.classList.contains("dark")
                          ? "#d1d5db"
                          : "#6b7280",
                      },
                      grid: {
                        color: document.documentElement.classList.contains("dark")
                          ? "rgba(107, 114, 128, 0.2)"
                          : "rgba(107, 114, 128, 0.1)",
                      },
                    },
                  },
                };

                return (
                  <div className="bg-gradient-to-br from-blue-50 to-indigo-50 dark:from-blue-900/20 dark:to-indigo-900/20 p-6 rounded-lg border border-blue-200 dark:border-blue-700">
                    <h3 className="text-base font-semibold text-gray-900 dark:text-white mb-4">
                      Section Performance Overview
                    </h3>
                    <div id="section-performance-chart" className="w-full" style={{ height: "400px" }}>
                      <Radar data={radarData} options={radarOptions} />
                    </div>
                  </div>
                );
              })()}

              {filteredSectionStats.map((sectionStat) => {
                const section = form?.sections?.find((s: any) => s.id === sectionStat.id);
                if (!section) return null;

                const questionStats = getSectionYesNoQuestionStats(sectionStat.id);
                const hasAnyYesNoInSection = questionStats.some(q => q.hasYesNo);
                const sectionTotals = {
                  yes: questionStats.reduce((sum, q) => sum + q.yes, 0),
                  no: questionStats.reduce((sum, q) => sum + q.no, 0),
                  na: questionStats.reduce((sum, q) => sum + q.na, 0),
                  correct: questionStats.reduce((sum, q) => sum + q.correct, 0),
                  wrong: questionStats.reduce((sum, q) => sum + q.wrong, 0),
                  total: questionStats.reduce((sum, q) => sum + q.total, 0),
                };

                const chartData = {
                  labels: [complianceLabels.correct, complianceLabels.wrong, complianceLabels.yes, complianceLabels.no, complianceLabels.na],
                  datasets: [
                    {
                      data: [sectionTotals.correct, sectionTotals.wrong, sectionTotals.yes, sectionTotals.no, sectionTotals.na],
                      backgroundColor: ["#059669", "#dc2626", "#1e40af", "#3b82f6", "#93c5fd"],
                      borderColor: ["#059669", "#dc2626", "#1e40af", "#3b82f6", "#93c5fd"],
                      borderWidth: 1,
                    },
                  ],
                };

                const chartOptions = {
                  responsive: true,
                  maintainAspectRatio: false,
                  plugins: {
                    legend: {
                      position: "bottom" as const,
                      labels: {
                        color: document.documentElement.classList.contains("dark")
                          ? "#d1d5db"
                          : "#374151",
                      },
                    },
                  },
                };

                return (
                  <div key={sectionStat.id} className="bg-gray-50 dark:bg-gray-800 p-4 rounded-lg border border-gray-200 dark:border-gray-700">
                    <h3 className="text-base font-semibold text-gray-900 dark:text-white mb-4">
                      {section.title || "Section"} - Response Analysis
                    </h3>
                    <div className="w-full h-64" id={`section-chart-${section.id}`}>
                      {sectionChartTypes[section.id] === "bar" ? (
                        <Bar
                          data={{
                            labels: questionStats.map(
                              (stat) => stat.subParam1 || "No parameter"
                            ),
                            datasets: [
                              {
                                label: complianceLabels.correct,
                                data: questionStats.map((stat) => stat.correct),
                                backgroundColor: "#059669",
                                borderColor: "#059669",
                                borderWidth: 1,
                              },
                              {
                                label: complianceLabels.wrong,
                                data: questionStats.map((stat) => stat.wrong),
                                backgroundColor: "#dc2626",
                                borderColor: "#dc2626",
                                borderWidth: 1,
                              },
                              {
                                label: `${complianceLabels.yes} / Answered`,
                                data: questionStats.map((stat) => stat.yes),
                                backgroundColor: "#1e40af",
                                borderColor: "#1e40af",
                                borderWidth: 1,
                              },
                              {
                                label: `${complianceLabels.no} / Not Answered`,
                                data: questionStats.map((stat) => stat.no),
                                backgroundColor: "#3b82f6",
                                borderColor: "#3b82f6",
                                borderWidth: 1,
                              },
                              {
                                label: complianceLabels.na,
                                data: questionStats.map((stat) => stat.na),
                                backgroundColor: "#93c5fd",
                                borderColor: "#93c5fd",
                                borderWidth: 1,
                              },
                            ],
                          }}
                          options={{
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                              legend: {
                                position: "top" as const,
                                labels: {
                                  color: document.documentElement.classList.contains(
                                    "dark"
                                  )
                                    ? "#d1d5db"
                                    : "#374151",
                                },
                              },
                            },
                          }}
                        />
                      ) : (
                        <Pie data={chartData} options={chartOptions} />
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            <div className="sticky bottom-0 bg-white dark:bg-gray-800 border-t border-gray-200 dark:border-gray-700 p-4 flex items-center justify-between">
              <button
                onClick={() => setShowSectionsPDFModal(false)}
                className="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleDownloadSectionsPDF}
                disabled={downloadingSectionsPDF}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white font-medium rounded-lg transition-colors"
              >
                {downloadingSectionsPDF ? (
                  <>
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Generating PDF...
                  </>
                ) : (
                  <>
                    <FileText className="w-4 h-4" />
                    Download PDF
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Update Confirmation Popup */}
      {showUpdatePopup && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[9999] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl p-6 max-w-md w-full border border-gray-200 dark:border-gray-700 animate-in fade-in zoom-in duration-300">
            <h3 className="text-xl font-bold text-gray-900 dark:text-white mb-4">
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
                  className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white"
                  placeholder="e.g. 05/10/24"
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
                  className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white"
                  placeholder="Enter revision details"
                />
              </div>
            </div>

            <div className="mt-6 flex justify-end gap-3">
              <button
                onClick={() => {
                  setShowUpdatePopup(false);
                  setSavingEdit(false);
                }}
                className="px-4 py-2 text-sm font-medium text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors"
                disabled={savingEdit}
              >
                Cancel
              </button>
              <button
                onClick={handleConfirmUpdate}
                disabled={savingEdit}
                className="px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition-colors disabled:opacity-50 flex items-center gap-2"
              >
                {savingEdit ? (
                  <>
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Saving...
                  </>
                ) : (
                  "Confirm Update"
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {pdfProgress && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[9999] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl p-8 max-w-md w-full border border-gray-200 dark:border-gray-700 animate-in fade-in zoom-in duration-300">
            <div className="flex flex-col items-center text-center">
              <div className="w-16 h-16 bg-blue-100 dark:bg-blue-900/30 rounded-full flex items-center justify-center mb-4">
                <FileText className="w-8 h-8 text-blue-600 dark:text-blue-400" />
              </div>
              <h3 className="text-xl font-bold text-gray-900 dark:text-white mb-2">
                {pdfProgress.stage === 'generating' ? 'Generating PDFs' :
                  pdfProgress.stage === 'downloading' ? 'Downloading' :
                    pdfProgress.stage === 'complete' ? 'Complete!' :
                      pdfProgress.stage === 'error' ? 'Error' : 'Processing'}
              </h3>
              <p className="text-gray-600 dark:text-gray-400 mb-6">
                {pdfProgress.message || 'Please wait while we prepare your files...'}
              </p>

              <div className="w-full bg-gray-200 dark:bg-gray-700 rounded-full h-3 mb-2 overflow-hidden">
                <div
                  className={`h-full transition-all duration-300 rounded-full ${pdfProgress.stage === 'error' ? 'bg-red-500' : 'bg-blue-600'
                    }`}
                  style={{ width: `${pdfProgress.percentage}%` }}
                />
              </div>
              <div className="flex justify-between w-full text-xs font-semibold text-gray-500 dark:text-gray-400">
                <span>{Math.round(pdfProgress.percentage)}%</span>
                <span>{pdfProgress.stage === 'complete' ? 'Success' : 'In Progress'}</span>
              </div>

              {(pdfProgress.stage === 'generating' || pdfProgress.stage === 'downloading') && (
                <button
                  onClick={() => {
                    isCancelledRef.current = true;
                    setPdfProgress(null);
                  }}
                  className="mt-6 px-6 py-2 bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400 font-bold rounded-xl hover:bg-red-200 dark:hover:bg-red-800/40 transition-all"
                >
                  Cancel
                </button>
              )}

              {(pdfProgress.stage === 'complete' || pdfProgress.stage === 'error') && (
                <button
                  onClick={() => setPdfProgress(null)}
                  className="mt-6 px-6 py-2 bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 font-bold rounded-xl hover:bg-gray-200 dark:hover:bg-gray-600 transition-all"
                >
                  Close
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );

}
