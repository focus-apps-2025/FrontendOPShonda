import React, { useState, useEffect, useMemo } from "react";
import { useNavigate, useParams, useLocation } from "react-router-dom";
import {
  Save,
  ArrowLeft,
  Plus,
  Trash2,
  Eye,
  X,
  Edit2,
  BarChart3,
  Copy,
  MoreVertical,
  Scissors,
  Users,
  Calendar,
  FileText,
  ChevronUp,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Clipboard,
  Link as LinkIcon,
  MessageSquarePlus,
  Download,
  Upload,
} from "lucide-react";
import { apiClient } from "../api/client";
import { questionsApi } from "../api/storage";
import { useNotification } from "../context/NotificationContext";
import { useAuth } from "../context/AuthContext";
import { NestedFollowUpRenderer } from "./NestedFollowUpRenderer";
import {
  getLevel1Options,
  getLevel2Options,
  getLevel3Options,
  getLevel4Options,
  getLevel5Options,
  getLevel6Options,
} from "../config/npsHierarchy";
import ChildFormsManager from "./forms/ChildFormsManager";
import { SectionBranchingConfig } from "./forms/SectionBranchingConfig";
import { FormRoutingConfig } from "./forms/FormRoutingConfig";
import PreviewForm from "./PreviewForm";
import ParameterModal from "./ParameterModal";

import {
  downloadFormImportTemplate,
  downloadNestedFormImportTemplate,
  parseFormWorkbook,
} from "../utils/exportUtils";

const YES_NO_NA_OPTIONS = ["Yes", "No", "N/A"];
const YES_NO_NA_CORRECT = "Yes";
const MAX_VISIBLE_PAGE_BUTTONS = 3;

interface FormSection {
  id: string;
  title: string;
  description?: string;
  questions: Question[];
  parentSectionId?: string; // For subsections
  isSubsection?: boolean; // Mark if this is a subsection
  nextSectionId?: string; // Target section for direct link
}

interface Question {
  id: string;
  text: string;
  type: string;
  required: boolean;
  options?: string[];
  allowedFileTypes?: string[];
  description?: string;
  imageUrl?: string;
  suggestion?: string;
  trackResponseRank?: boolean;
  trackResponseRankLabel?: string;
  trackResponseRankType?: string;
  subParam1?: string;
  subParam2?: string;
  followUpQuestions?: FollowUpQuestion[];
  showWhen?: ShowWhen;
  parentId?: string;
  correctAnswer?: string;
  followUpConfig?: Record<
    string,
    {
      hasFollowUp: boolean;
      required: boolean;
      linkedFormId?: string;
    }
  >;
  branchingRules?: Array<{
    optionLabel: string;
    targetSectionId: string;
    isOtherOption?: boolean;
  }>;
  hierarchyLevels?: Array<{
    levelNumber: number;
    name: string;
    enabled: boolean;
  }>;
  selectedHierarchyValues?: {
    level1?: string;
    level2?: string;
    level3?: string;
    level4?: string;
    level5?: string;
    level6?: string;
  };
  // For chassisNumber type - stores tenant assignment for each chassis number
  chassisNumberTenantConfig?: Array<{
    chassisNumber: string;
    assignedTenants: string[];
  }>;
}

interface ShowWhen {
  questionId: string;
  value: string | number;
}

interface FollowUpQuestion {
  id: string;
  text: string;
  type: string;
  required: boolean;
  options?: string[];
  allowedFileTypes?: string[];
  description?: string;
  imageUrl?: string;
  suggestion?: string;
  trackResponseRank?: boolean;
  trackResponseRankLabel?: string;
  trackResponseRankType?: string;

  subParam1?: string;
  subParam2?: string;
  showWhen?: ShowWhen;
  parentId: string;
  followUpQuestions?: FollowUpQuestion[]; // Support nested follow-ups
  requireFollowUp?: boolean; // Make follow-up mandatory for certain question types
  correctAnswer?: string;
  suggestion?: string;
}

export default function FormCreator() {
  const navigate = useNavigate();
  const { id } = useParams<{ id: string }>();
  const location = useLocation();
  const { user, tenant } = useAuth();
  const isInspector = user?.role === "inspector";
  const canManage = (user?.role === "admin" || user?.role === "superadmin" || user?.role === "subadmin") && !isInspector;
  const [mode, setMode] = useState<"list" | "create">(
    id ||
      location.pathname.includes("/create") ||
      location.pathname.includes("/edit")
      ? "create"
      : "list"
  );
  const [forms, setForms] = useState<any[]>([]);
const [showPreview, setShowPreview] = useState(false);
   const [tenants, setTenants] = useState<any[]>([]);
   const [selectedTenantId, setSelectedTenantId] = useState<string>("");
   const [isGlobal, setIsGlobal] = useState<boolean>(
     new URLSearchParams(location.search).get("isGlobal") === "true"
   );
   const [sharedWithTenants, setSharedWithTenants] = useState<string[]>([]);
   const [currentPage, setCurrentPage] = useState<number>(0); // For multi-page navigation
   const [pageWindowStart, setPageWindowStart] = useState<number>(0);
   const [openOptionMenu, setOpenOptionMenu] = useState<string | null>(null); // Track which option's menu is open
   const [openQuestionMenu, setOpenQuestionMenu] = useState<string | null>(null); // Track which question's menu is open
   const [openChassisMenu, setOpenChassisMenu] = useState<number | null>(null); // Track which chassis's menu is open
   const [showSectionSelector, setShowSectionSelector] = useState(false); // Show modal for selecting a section
   const [pendingSectionLink, setPendingSectionLink] = useState<{
     sectionId: string;
     parentQuestionId: string;
     triggerValue: string;
     path: string[];
   } | null>(null);
   const [showFormSelector, setShowFormSelector] = useState(false); // Show modal for selecting a form
   const [pendingFormLink, setPendingFormLink] = useState<{
     sectionId: string;
     parentQuestionId: string;
     triggerValue: string;
     path: string[];
   } | null>(null);
   const [showBranchingConfig, setShowBranchingConfig] = useState(false);
   const [branchingConfigQuestion, setBranchingConfigQuestion] = useState<{
     questionId: string;
     sectionId: string;
     options: string[];
   } | null>(null);
   // State for chassis number tenant configuration
   const [showChassisTenantModal, setShowChassisTenantModal] = useState(false);
   const [chassisTenantConfig, setChassisTenantConfig] = useState<{
     chassisNumber: string;
     index: number;
   } | null>(null);
   const [newChassisNumber, setNewChassisNumber] = useState("");
   const [newPartDescription, setNewPartDescription] = useState("");
   const [chassisTenantAssignments, setChassisTenantAssignments] = useState<Record<string, string[]>>({});
   const [form, setForm] = useState({
     id: "",
     title: "",
     description: "",
     isVisible: true,
     locationEnabled: true,
     sections: [
       {
         id: crypto.randomUUID(),
         title: "Section 1",
         description: "",
         questions: [],
       },
     ] as FormSection[],
     followUpQuestions: [] as FollowUpQuestion[],
     chassisNumbers: [] as Array<{ chassisNumber: string; partDescription: string }>,
   });
   const [formSectionBranching, setFormSectionBranching] = useState<any[]>([]);
   const [showFormRoutingConfig, setShowFormRoutingConfig] = useState(false);
   const [formRoutingConfigQuestion, setFormRoutingConfigQuestion] = useState<{
     questionId: string;
     sectionId: string;
     options: string[];
   } | null>(null);
   const [showParameterModal, setShowParameterModal] = useState(false);
   const [parameters, setParameters] = useState<any[]>([]);
   const [tempParameters, setTempParameters] = useState<any[]>([]);
   const { showSuccess, showError, showConfirm } = useNotification();

  useEffect(() => {
    const params = new URLSearchParams(location.search);
    if (params.get("isGlobal") === "true") {
      setIsGlobal(true);
      setMode("create");
    } else if (
      location.pathname.includes("/create") ||
      location.pathname.includes("/edit") ||
      id
    ) {
      setMode("create");
    }
  }, [location.search, location.pathname, id]);

  // Fetch tenants for superadmin
  useEffect(() => {
    const fetchTenants = async () => {
      if (user?.role === "superadmin") {
        try {
          const response = await apiClient.getTenants();
          setTenants(response.tenants || []);
          // Auto-select first tenant if available
          if (response.tenants && response.tenants.length > 0) {
            setSelectedTenantId(response.tenants[0]._id);
          }
        } catch (error) {
          console.error("Failed to fetch tenants:", error);
        }
      }
    };
    fetchTenants();
  }, [user]);

  useEffect(() => {
    const loadForms = async () => {
      try {
        const response = await apiClient.getForms();
        setForms(response.forms || []);
      } catch (error) {
        console.error("Failed to fetch forms:", error);
        setForms([]);
      }
    };
    loadForms();

    if (id) {
      // Load existing form from API for editing
      const loadForm = async () => {
        try {
          const response = await apiClient.getForm(id);
          const backendForm = response.form;

          console.log("=== DEBUG: Form loaded from backend ===");
          console.log("Backend form structure:", backendForm);

          // Check first few questions for suggestions
          backendForm.sections?.forEach((section: any, sIndex: number) => {
            console.log(`\nSection ${sIndex + 1}: ${section.title}`);
            section.questions?.slice(0, 3).forEach((q: any, qIndex: number) => {
              console.log(`  Q${qIndex + 1}: "${q.text?.substring(0, 50)}..."`);
              console.log(`    Has suggestion field: ${"suggestion" in q}`);
              console.log(
                `    Suggestion value: "${q.suggestion || "NO SUGGESTION"}"`
              );
            });
          });

          // Set the tenant ID from the loaded form
          if (backendForm.tenantId) {
            setSelectedTenantId(backendForm.tenantId);
          }

          if (backendForm.isGlobal !== undefined) {
            setIsGlobal(backendForm.isGlobal);
          }

          if (backendForm.sharedWithTenants) {
            setSharedWithTenants(backendForm.sharedWithTenants);
          }

           // Transform backend form to frontend format
           // Reconstruct nested follow-up questions from flat array or use existing nested structure
           const sectionsWithNestedFollowUps = (backendForm.sections || []).map(
             (section: any) => {
               const mainQuestions: Question[] = [];
               const followUpMap = new Map<string, FollowUpQuestion[]>();

               // First pass: separate main questions and follow-ups (for backward compatibility with flat structure)
               section.questions.forEach((q: any) => {
                 if (q.showWhen && q.showWhen.questionId) {
                   // This is a follow-up question
                   const parentId = q.showWhen.questionId;
                   if (!followUpMap.has(parentId)) {
                     followUpMap.set(parentId, []);
                   }
                   followUpMap.get(parentId)!.push(q as FollowUpQuestion);
                 } else {
                   // This is a main question
                   mainQuestions.push(q as Question);
                 }
               });

               // Second pass: attach follow-ups to their parent questions and initialize followUpConfig
               const questionsWithFollowUps = mainQuestions.map((q) => {
                 // Use existing nested follow-ups if present, otherwise reconstruct from flat array
                 const followUps =
                   q.followUpQuestions && q.followUpQuestions.length > 0
                     ? q.followUpQuestions
                     : followUpMap.get(q.id) || [];

                 // Initialize followUpConfig if not present
                 const followUpConfig = q.followUpConfig || {};
                 if (q.options && q.options.length > 0) {
                   q.options.forEach((option: string) => {
                     if (!followUpConfig[option]) {
                       followUpConfig[option] = {
                         hasFollowUp: false,
                         required: false,
                       };
                     }
                   });
                   // Mark options that have follow-ups
                   followUps.forEach((followUp) => {
                     if (
                       followUp.showWhen?.value &&
                       followUpConfig[followUp.showWhen.value]
                     ) {
                       followUpConfig[followUp.showWhen.value].hasFollowUp =
                         true;
                     }
                   });
                 }

                 return {
                   ...q,
                   followUpQuestions: followUps,
                   followUpConfig: followUpConfig,
                 };
               });

               return {
                 ...section,
                 id: section.id || section._id,
                 nextSectionId: section.nextSectionId,
                 questions: questionsWithFollowUps,
               };
             }
           );

           setForm({
             id: backendForm.id,
             title: backendForm.title,
             description: backendForm.description,
             isVisible: backendForm.isVisible,
             locationEnabled: backendForm.locationEnabled !== false,
             sections: sectionsWithNestedFollowUps,
             followUpQuestions: backendForm.followUpQuestions || [],
             chassisNumbers: (backendForm.chassisNumbers || []).map((cn: any) => 
               typeof cn === 'string' ? { chassisNumber: cn, partDescription: '' } : cn
             ),
           });

          if (backendForm.sectionBranching) {
            setFormSectionBranching(backendForm.sectionBranching);
            console.log(
              "Branching rules loaded from form:",
              backendForm.sectionBranching
            );
          }
          // Load chassis number tenant configuration
          if (Array.isArray(backendForm.chassisTenantAssignments)) {
            const assignmentsRecord: Record<string, string[]> = {};
            backendForm.chassisTenantAssignments.forEach((assignment: any) => {
              const key = assignment.chassisNumber;
              if (key) {
                assignmentsRecord[key] = assignment.assignedTenants;
              }
            });
            setChassisTenantAssignments(assignmentsRecord);
          } else if (backendForm.chassisTenantAssignments) {
            // Already a record (backward compatibility)
            setChassisTenantAssignments(backendForm.chassisTenantAssignments);
          }
        }
         catch (error) {
          console.error("Failed to load form:", error);
           // Fallback to local storage if API fails
          const existingForm = questionsApi.getById(id);
          if (existingForm) {
            setForm(existingForm);
          }
        }
      };
      loadForm();
      setMode("create");
    } else {
      // Check if we should start in create mode from navigation state
      const state = location.state as { mode?: string; formData?: any };
      if (state?.mode === "create") {
        setMode("create");
        if (state.formData) {
          // Load pre-populated form data
          setForm({
            id: "",
            title: state.formData.title || "",
            description: state.formData.description || "",
            isVisible: state.formData.isVisible !== false,
            locationEnabled: state.formData.locationEnabled !== false,
            sections: (state.formData.sections || []).map((section: any) => ({
              id: section.id || crypto.randomUUID(),
              title: section.title || "",
              description: section.description || "",
              questions: (section.questions || []).map((question: any) => ({
                id: question.id || crypto.randomUUID(),
                text: question.text || "",
                type: question.type || "text",
                required: question.required || false,
                options: question.options || undefined,
                allowedFileTypes: question.allowedFileTypes || undefined,
                description: question.description || undefined,
                imageUrl: question.imageUrl || undefined,
                subParam1: question.subParam1 || undefined,
                subParam2: question.subParam2 || undefined,
                trackResponseRank: question.trackResponseRank || false,
                followUpQuestions: question.followUpQuestions || [],
                showWhen: question.showWhen || undefined,
                parentId: question.parentId || undefined,
                correctAnswer: question.correctAnswer || undefined,
                followUpConfig: question.followUpConfig || undefined,
                branchingRules: question.branchingRules || undefined,
              })),
            })),
            followUpQuestions: state.formData.followUpQuestions || [],
            chassisNumbers: (state.formData.chassisNumbers || []).map((cn: any) => 
              typeof cn === 'string' ? { chassisNumber: cn, partDescription: '' } : cn
            ),
          });
        } else {
          // Start with empty form
          setForm({
            id: "",
            title: "",
            description: "",
            isVisible: true,
            locationEnabled: true,
            sections: [
              {
                id: crypto.randomUUID(),
                title: "Section 1",
                description: "",
                questions: [],
              },
            ],
            followUpQuestions: [],
            chassisNumbers: [],
          });
        }
      } else {
        // If we are on /create or /edit path, ensure we are in create mode
        if (
          location.pathname.includes("/create") ||
          location.pathname.includes("/edit")
        ) {
          setMode("create");
        } else {
          setMode("list");
        }
      }
    }
  }, [id, location.state, location.pathname]);

  // Fetch tenants when chassis modal opens
  useEffect(() => {
    if (showChassisTenantModal && tenants.length === 0) {
      const fetchTenants = async () => {
        try {
          const response = await apiClient.getTenantsMinimal();
          setTenants(response.tenants || []);
        } catch (error) {
          console.error("Failed to fetch tenants for modal:", error);
        }
      };
      fetchTenants();
    }
  }, [showChassisTenantModal, tenants.length]);

  // Fetch parameters
  useEffect(() => {
    const fetchParameters = async () => {
      if (id) {
        // For existing forms, fetch from API
        try {
          const response = await apiClient.getParameters({ formId: id });
          setParameters(response.parameters || []);
        } catch (error) {
          console.error("Failed to fetch parameters:", error);
        }
      } else {
        // For new forms, use temporary parameters
        setParameters(tempParameters);
      }
    };
    fetchParameters();
  }, [id, tempParameters]);

  // Helper function to determine if a question is a followup question
  const isFollowupQuestion = (question: Question, section: FormSection) => {
    // Check if question has parentId (for nested followups)
    if (question.parentId) return true;

    // Check if question is in followUpQuestions of another question
    const isInFollowUps = section.questions.some((q) =>
      q.followUpQuestions?.some((fq) => fq.id === question.id)
    );

    return isInFollowUps;
  };

  const handleAddChassis = () => {
    if (!newChassisNumber.trim()) return;
    const chassisKey = newChassisNumber.trim();
    if (form.chassisNumbers.some((cn) => cn.chassisNumber === chassisKey)) {
      showError("Chassis number already exists", "Duplicate Error");
      return;
    }
    setForm((prev) => ({
      ...prev,
      chassisNumbers: [...prev.chassisNumbers, { chassisNumber: chassisKey, partDescription: newPartDescription.trim() }],
    }));
    setNewChassisNumber("");
    setNewPartDescription("");
  };

  const handleRemoveChassis = (index: number) => {
    const chassisToRemove = form.chassisNumbers[index];
    setForm((prev) => ({
      ...prev,
      chassisNumbers: prev.chassisNumbers.filter((_, i) => i !== index),
    }));
    // Clean up assignments
    const newAssignments = { ...chassisTenantAssignments };
    delete newAssignments[chassisToRemove.chassisNumber];
    setChassisTenantAssignments(newAssignments);
  };

  const handleExportTemplate = async (type: 'flat' | 'nested' = 'flat') => {
    try {
      if (type === 'nested') {
        downloadNestedFormImportTemplate();
      } else {
        await downloadFormImportTemplate({}, undefined, form.opsTemplateConfig);
      }
      showSuccess(`${type === 'nested' ? 'Nested' : 'Flat'} template exported successfully!`, "Export Template");
    } catch (error) {
      console.error("Failed to export template:", error);
      showError("Failed to export template", "Error");
    }
  };

  const handleImportTemplate = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const importedData = await parseFormWorkbook(file);

       const newForm = {
         id: "",
         title: (importedData as any).title || "Imported Form",
         description: (importedData as any).description || "",
         isVisible: true,
         locationEnabled: true,
         sections: (importedData as any).sections || [],
         followUpQuestions: [],
         chassisNumbers: [],
       };

      setForm(newForm);

      // Extract branching rules from imported questions to populate formSectionBranching state
      const branchingRulesToSet: any[] = [];
      importedData.sections.forEach((section: any) => {
        section.questions.forEach((question: any) => {
          if (question.branchingRules && question.branchingRules.length > 0) {
            branchingRulesToSet.push({
              sectionId: section.id,
              questionId: question.id,
              rules: question.branchingRules.map((rule: any) => ({
                optionLabel: rule.optionLabel,
                targetSectionId: rule.targetSectionId,
              })),
            });
          }
        });
      });

      if (branchingRulesToSet.length > 0) {
        console.log("Setting imported branching rules:", branchingRulesToSet);
        setFormSectionBranching(branchingRulesToSet);
      }

      if (
        importedData.parametersToCreate &&
        importedData.parametersToCreate.length > 0
      ) {
        setTempParameters(importedData.parametersToCreate);
      }
      showSuccess("Template imported successfully!", "Import Template");
    } catch (error) {
      console.error("Failed to import template:", error);
      showError(
        error instanceof Error ? error.message : "Failed to import template",
        "Import Error"
      );
    }
  };

  // Function to load sample data for testing
  const loadSampleData = async () => {
    try {
      showConfirm(
        "This will replace your current form with sample data. Continue?",
        "Load Sample Data",
        async () => {
          // Create sample parameters first
          const tenantId =
            user?.role === "superadmin" ? selectedTenantId : user?.tenantId;

          const sampleParameters = [
            // Main parameters
            { name: "Safety Score", type: "main" as const },
            { name: "Quality Rating", type: "main" as const },
            { name: "Efficiency Index", type: "main" as const },
            { name: "Compliance Level", type: "main" as const },
            // Followup parameters
            { name: "Root Cause", type: "followup" as const },
            { name: "Action Required", type: "followup" as const },
            { name: "Timeline", type: "followup" as const },
            { name: "Responsible Party", type: "followup" as const },
          ];

          // Create parameters in batch
          const createdParameters = [];
          for (const param of sampleParameters) {
            try {
              const response = await apiClient.createParameter({
                ...param,
                tenantId,
              });
              createdParameters.push(response.parameter);
            } catch (error) {
              console.warn(`Failed to create parameter ${param.name}:`, error);
            }
          }

          // Update parameters state
          setParameters((prev) => [...prev, ...createdParameters]);

          // Create sample form data
          const sampleForm = {
            title: "Manufacturing Quality Inspection Form",
            description:
              "Comprehensive quality inspection form with follow-up questions for manufacturing processes",
            isVisible: true,
            locationEnabled: true,
            sections: [
              {
                id: crypto.randomUUID(),
                title: "Process Overview",
                description: "Initial assessment of the manufacturing process",
                questions: [
                  {
                    id: crypto.randomUUID(),
                    text: "What is the current production line status?",
                    type: "radio",
                    required: true,
                    options: [
                      "Running Normally",
                      "Minor Issues",
                      "Major Issues",
                      "Stopped",
                    ],
                    subParam1: "Efficiency Index",
                    subParam2: "Safety Score",
                    followUpConfig: {
                      "Minor Issues": { hasFollowUp: true, required: true },
                      "Major Issues": { hasFollowUp: true, required: true },
                      Stopped: { hasFollowUp: true, required: true },
                    },
                    followUpQuestions: [
                      {
                        id: crypto.randomUUID(),
                        text: "Please describe the issue in detail",
                        type: "textarea",
                        required: true,
                        subParam1: "Root Cause",
                        subParam2: "Action Required",
                        showWhen: { questionId: "", value: "Minor Issues" },
                        followUpQuestions: [
                          {
                            id: crypto.randomUUID(),
                            text: "Who is responsible for resolving this issue?",
                            type: "text",
                            required: true,
                            subParam1: "Responsible Party",
                            showWhen: { questionId: "", value: "" },
                          },
                        ],
                      },
                      {
                        id: crypto.randomUUID(),
                        text: "What is the estimated downtime?",
                        type: "select",
                        required: true,
                        options: [
                          "< 1 hour",
                          "1-4 hours",
                          "4-8 hours",
                          "> 8 hours",
                        ],
                        subParam1: "Timeline",
                        showWhen: { questionId: "", value: "Major Issues" },
                      },
                    ],
                  },
                  {
                    id: crypto.randomUUID(),
                    text: "Rate the overall quality of recent production",
                    type: "radio",
                    required: true,
                    options: [
                      "Excellent",
                      "Good",
                      "Average",
                      "Poor",
                      "Critical",
                    ],
                    subParam1: "Quality Rating",
                    followUpConfig: {
                      Poor: { hasFollowUp: true, required: true },
                      Critical: { hasFollowUp: true, required: true },
                    },
                    followUpQuestions: [
                      {
                        id: crypto.randomUUID(),
                        text: "What quality issues were identified?",
                        type: "textarea",
                        required: true,
                        subParam1: "Root Cause",
                        showWhen: { questionId: "", value: "Poor" },
                      },
                    ],
                  },
                ],
              },
              {
                id: crypto.randomUUID(),
                title: "Compliance Check",
                description: "Verification of regulatory and safety compliance",
                questions: [
                  {
                    id: crypto.randomUUID(),
                    text: "Are all safety protocols being followed?",
                    type: "radio",
                    required: true,
                    options: ["Yes", "No", "Partially"],
                    subParam1: "Safety Score",
                    subParam2: "Compliance Level",
                    followUpConfig: {
                      No: { hasFollowUp: true, required: true },
                      Partially: { hasFollowUp: true, required: true },
                    },
                    followUpQuestions: [
                      {
                        id: crypto.randomUUID(),
                        text: "Which protocols are not being followed?",
                        type: "textarea",
                        required: true,
                        subParam1: "Root Cause",
                        showWhen: { questionId: "", value: "No" },
                      },
                    ],
                  },
                ],
              },
              {
                id: crypto.randomUUID(),
                title: "Final Assessment",
                description: "Overall evaluation and recommendations",
                questions: [
                  {
                    id: crypto.randomUUID(),
                    text: "Overall process efficiency rating",
                    type: "radio",
                    required: true,
                    options: [
                      "Highly Efficient",
                      "Efficient",
                      "Needs Improvement",
                      "Inefficient",
                    ],
                    subParam1: "Efficiency Index",
                    followUpConfig: {
                      "Needs Improvement": {
                        hasFollowUp: true,
                        required: false,
                      },
                      Inefficient: { hasFollowUp: true, required: true },
                    },
                    followUpQuestions: [
                      {
                        id: crypto.randomUUID(),
                        text: "What improvements are recommended?",
                        type: "textarea",
                        required: true,
                        subParam1: "Action Required",
                        showWhen: {
                          questionId: "",
                          value: "Needs Improvement",
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          };

          // Fix followUp questionId references
          sampleForm.sections.forEach((section) => {
            section.questions.forEach((question) => {
              if (question.followUpQuestions) {
                question.followUpQuestions.forEach((followUp) => {
                  if (followUp.showWhen) {
                    followUp.showWhen.questionId = question.id;
                  }
                  if (followUp.followUpQuestions) {
                    followUp.followUpQuestions.forEach((nestedFollowUp) => {
                      if (nestedFollowUp.showWhen) {
                        nestedFollowUp.showWhen.questionId = followUp.id;
                      }
                    });
                  }
                });
              }
            });
          });

          setForm({
            ...sampleForm,
            id: "",
            followUpQuestions: [],
            chassisNumbers: [],
          });
          showSuccess("Sample data loaded successfully!", "Sample Data");
        }
      );
    } catch (error) {
      console.error("Failed to load sample data:", error);
      showError("Failed to load sample data", "Error");
    }
  };

  const handleSave = async () => {
    if (!canManage) {
      showError("You do not have permission to save forms.", "Permission Denied");
      return;
    }
    if (!form.title.trim()) {
      showError("Please enter a form title", "Validation Error");
      return;
    }

    if (!form.description.trim()) {
      showError("Please enter a form description", "Validation Error");
      return;
    }

    // Validate tenantId for superadmin
    if (user?.role === "superadmin" && !selectedTenantId && !isGlobal) {
      showError("Please select a tenant for this form", "Validation Error");
      return;
    }

    // Determine tenantId based on user role
    const tenantId =
      user?.role === "superadmin" ? (isGlobal ? undefined : selectedTenantId) : user?.tenantId;

    // Validate tenantId exists
    if (!tenantId && !isGlobal) {
      showError(
        "Unable to determine tenant. Please try logging in again.",
        "Validation Error"
      );
      return;
    }

    // Check form limit for free trial tenants
    if (!id && user?.role !== "superadmin" && tenant?.subscription?.plan === "free") {
      const tenantFormsCount = forms.length;
      if (tenantFormsCount >= 5) {
        showError(
          "You have reached the limit of 5 forms for the Free Trial plan. Please contact admin to upgrade your plan.",
          "Limit Reached"
        );
        return;
      }
    }

    // Validate sections
    if (!form.sections || form.sections.length === 0) {
      showError("Form must have at least one section", "Validation Error");
      return;
    }

    // Validate each section has required fields
    for (const section of form.sections) {
      if (!section.id) {
        showError("All sections must have an ID", "Validation Error");
        return;
      }
      if (!section.title || !section.title.trim()) {
        showError("All sections must have a title", "Validation Error");
        return;
      }

      // Validate questions in each section
      for (const question of section.questions) {
        if (!question.id) {
          showError(
            `Question in section "${section.title}" is missing an ID`,
            "Validation Error"
          );
          return;
        }

        const questionHasText = Boolean(question.text && question.text.trim());
        const questionHasImage = Boolean(
          question.imageUrl && question.imageUrl.trim()
        );
        const questionLabel = question.text?.trim() || "Image question";

        if (!questionHasText && !questionHasImage) {
          showError(
            `Question in section "${section.title}" must include text or an image`,
            "Validation Error"
          );
          return;
        }

        if (!question.type) {
          showError(
            `Question "${questionLabel}" in section "${section.title}" is missing a type`,
            "Validation Error"
          );
          return;
        }

        if (
          question.followUpQuestions &&
          question.followUpQuestions.length > 0
        ) {
          for (const followUp of question.followUpQuestions) {
            if (!followUp.id) {
              showError(
                `Follow-up question for "${questionLabel}" is missing an ID`,
                "Validation Error"
              );
              return;
            }

            const followUpHasText = Boolean(
              followUp.text && followUp.text.trim()
            );
            const followUpHasImage = Boolean(
              followUp.imageUrl && followUp.imageUrl.trim()
            );

            if (!followUpHasText && !followUpHasImage) {
              showError(
                `Follow-up question for "${questionLabel}" must include text or an image`,
                "Validation Error"
              );
              return;
            }

            const followUpLabel = followUp.text?.trim() || "Image question";

            if (!followUp.type) {
              showError(
                `Follow-up question "${followUpLabel}" is missing a type`,
                "Validation Error"
              );
              return;
            }
          }
        }
      }
    }

    // Flatten follow-up questions into section questions array (recursively)
    const formToSave = {
      ...form,
      // Include tenantId for form creation
      tenantId: tenantId,
      isGlobal: isGlobal,
      sharedWithTenants: sharedWithTenants,
      chassisNumbers: form.chassisNumbers || [],
      // Include chassis number tenant configuration (Convert Record to Array of Objects for Mongoose/Mongo indexing)
      chassisTenantAssignments: Object.entries(chassisTenantAssignments || {}).map(([chassisNumber, tenants]) => {
        return {
          chassisNumber,
          assignedTenants: tenants,
        };
      }),
      sections: form.sections.map((section) => {
        const allQuestions: Question[] = [];

        // Recursive function to process a question and all its nested follow-ups
        const processQuestionRecursively = (
          question: Question,
          depth: number = 0
        ) => {
          const indent = "  ".repeat(depth);
          const questionLabelForLog = question.text?.trim() || "Image question";
          console.log(
            `${indent}Processing question: "${questionLabelForLog}" (depth: ${depth})`
          );

          const { followUpQuestions, ...mainQuestion } = question;
          allQuestions.push(mainQuestion as Question);

          if (followUpQuestions && followUpQuestions.length > 0) {
            console.log(
              `${indent}Found ${followUpQuestions.length} follow-up questions`
            );
            followUpQuestions.forEach((followUp) => {
              const followUpWithShowWhen = {
                ...followUp,
                showWhen: followUp.showWhen || {
                  questionId: question.id,
                  value: followUp.showWhen?.value || "",
                },
              };
              const followUpLabelForLog =
                followUpWithShowWhen.text?.trim() || "Image question";
              console.log(
                `${indent}Adding follow-up: "${followUpLabelForLog}"`
              );

              processQuestionRecursively(
                followUpWithShowWhen as Question,
                depth + 1
              );
            });
          }
        };

        // Process only main questions (those without showWhen)
        section.questions.forEach((question) => {
          if (!question.showWhen) {
            processQuestionRecursively(question);
          }
        });

        return {
          ...section,
          questions: allQuestions,
          nextSectionId: section.nextSectionId,
        };
      }),
    };

    console.log("=== Form Validation Passed ===");
    console.log("TenantId:", tenantId);
    console.log("User role:", user?.role);
    console.log("Form to save:", JSON.stringify(formToSave, null, 2));

    try {
      if (id) {
        // Update existing form
        console.log("Updating form with ID:", id);
        await apiClient.updateForm(id, formToSave);

        // Save branching rules if any exist
        if (formSectionBranching.length > 0) {
          console.log("Saving branching rules:", formSectionBranching);
          await apiClient.request(`/forms/${id}/section-branching`, {
            method: "POST",
            body: JSON.stringify({ rules: formSectionBranching }),
          });
          console.log("Branching rules saved successfully");
        }

        showSuccess("Form updated successfully", "Success");
        navigate("/superadmin/forms");
      } else {
        // Create new form
        console.log("Creating new form...");
        const response = await apiClient.createForm(formToSave);
        console.log("Form created successfully:", response);
        const newFormId = response.form._id;

        // Save branching rules if any exist
        if (formSectionBranching.length > 0) {
          console.log(
            "Saving branching rules for new form:",
            formSectionBranching
          );
          await apiClient.request(`/forms/${newFormId}/section-branching`, {
            method: "POST",
            body: JSON.stringify({ rules: formSectionBranching }),
          });
          console.log("Branching rules saved successfully");
        }

        // Create temporary parameters if any exist
        if (tempParameters.length > 0) {
          console.log("Creating parameters for new form:", tempParameters);
          try {
            const createPromises = tempParameters.map((param) =>
              apiClient.createParameter({
                name: param.name,
                type: param.type,
                formId: newFormId,
              })
            );
            await Promise.all(createPromises);
            console.log("Parameters created successfully");
            // Clear temporary parameters
            setTempParameters([]);
          } catch (error) {
            console.error("Failed to create parameters for new form:", error);
            showError(
              "Form created but failed to create parameters",
              "Warning"
            );
          }
        }

        showSuccess("Form created successfully", "Success");
        navigate("/superadmin/forms");
      }
    } catch (error: any) {
      console.error("=== Error saving form ===");
      console.error("Error object:", error);
      console.error("Error message:", error.message);
      console.error("Error status:", error.status);
      console.error("Error response:", error.response);

      // Show detailed error message
      const errorMsg =
        error.message ||
        "Error saving form. Please check the console for details.";
      showError(errorMsg, "Error");
    }
  };

  const handleDelete = (id: string, title: string) => {
    showConfirm(
      `Are you sure you want to delete "${title}"? This action cannot be undone.`,
      async () => {
        try {
          await apiClient.deleteForm(id);
          const formsResponse = await apiClient.getForms();
          setForms(formsResponse.forms || []);
          showSuccess("Form deleted successfully", "Success");
        } catch (error) {
          console.error("Failed to delete form:", error);
          showError("Failed to delete form", "Error");
        }
      },
      "Delete Form",
      "Delete",
      "Cancel"
    );
  };

  const handleDuplicate = async (form: any) => {
    try {
      await apiClient.duplicateForm(form._id);
      const formsResponse = await apiClient.getForms();
      setForms(formsResponse.forms || []);
      showSuccess("Form duplicated successfully", "Success");
    } catch (error) {
      console.error("Failed to duplicate form:", error);
      showError("Failed to duplicate form", "Error");
    }
  };

  const handleToggleVisibility = async (
    id: string,
    currentVisibility: boolean
  ) => {
    try {
      await apiClient.updateFormVisibility(id, !currentVisibility);
      const formsResponse = await apiClient.getForms();
      setForms(formsResponse.forms || []);
      showSuccess("Form visibility updated", "Success");
    } catch (error) {
      console.error("Failed to update form visibility:", error);
      showError("Failed to update form visibility", "Error");
    }
  };

  const handleCreateForm = () => {
    setMode("create");
    setForm({
      id: "",
      title: "",
      description: "",
      isVisible: true,
      locationEnabled: true,
      sections: [
        {
          id: crypto.randomUUID(),
          title: "Section 1",
          description: "",
          questions: [],
        },
      ] as FormSection[],
      followUpQuestions: [],
      chassisNumbers: [],
    });
  };

  const loadDemoData = () => {
    type SampleDataRow = {
      [key: string]: string | undefined;
    };
    const sampleData: SampleDataRow[] = [
      {
        "Form Title": "Bike Service & Maintenance Form",
        "Form Description":
          "Comprehensive bike service assessment with nested diagnostics",
        "Section Number": "1",
        "Section Title": "Basic Bike Information",
        "Section Description": "Basic details about the bike",
        "Section Merging": "",
        "After Section Action": "",
        Question: "What is your bike make and model?",
        "Question Description": "Manufacturer and specific model",
        "Question Type": "shortText",
        Required: "TRUE",
        Options: "",
        "Section Routing": "",
        Suggestion: "",
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
        "Section Number": "1.1",
        "Section Title": "Owner Details",
        "Section Description": "Details about the bike owner",
        "Subsection Of": "1",
        "After Section Action": "",
        Question: "Owner Name",
        "Question Type": "shortText",
        Required: "TRUE",
        Options: "",
        SubParam1: "Owner",
        SubParam2: "Identification",
      },
      {
        "Section Number": "2",
        "Section Title": "Service Requirements Assessment",
        "Section Description": "Evaluate what service the bike needs",
        "Section Merging": "",
        "After Section Action": "3",

        // ========== MAIN QUESTION 1: ENGINE ISSUES (WITH NESTED FOLLOW-UPS) ==========
        Question: "Are you experiencing any engine issues?",
        "Question Description": "Problems related to engine performance",
        "Question Type": "yesNoNA",
        Required: "TRUE",
        Options: "Yes,No,N/A",
        "Section Routing": "",
        Suggestion: "",
        SubParam1: "Engine Health",
        SubParam2: "Performance",

        // FU1: FOR YES (WITH NESTING)
        "FU1: Option": "Yes",
        "FU1: Question Type": "multipleChoice",
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
        "FU1.2: Question Type": "multipleChoice",
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
        "FU1: Question Type": "multipleChoice",
        "FU1: Required": "TRUE",
        "FU1: SubParam1": "Brake Problem Type",
        "FU1: SubParam2": "Safety Issue",
        "FU1: Question Text": "What brake problem are you experiencing?",
        "FU1: Options":
          "Soft Brake Lever,Grinding Noise,Brake Drag,Poor Stopping,Spongy Feel",

        // FU1.1: Nested under "Grinding Noise"
        "FU1.1: Option": "Grinding Noise",
        "FU1.1: Question Type": "multipleChoice",
        "FU1.1: Required": "TRUE",
        "FU1.1: SubParam1": "Noise Details",
        "FU1.1: SubParam2": "Wear Indicators",
        "FU1.1: Question Text": "When do you hear the grinding noise?",
        "FU1.1: Options":
          "Always When Braking,Only During Hard Braking,When Releasing Brakes,With Specific Speed",

        // FU1.1.1: Nested under "Always When Braking"
        "FU1.1.1: Option": "Always When Braking",
        "FU1.1.1: Question Type": "multipleChoice",
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
        "FU1: Options":
          "Too Soft,Too Hard,Uneven,Bottoming Out,Noise Over Bumps",

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

    // Helper function to map question types
    const mapQuestionType = (type: string | undefined): string => {
      if (!type) return "text";

      const normalized = type.toLowerCase().trim().replace(/\s+/g, '');

      const typeMap: Record<string, string> = {
        shorttext: "text",
        number: "number",
        multiplechoice: "radio",
        radio: "radio",
        dropdown: "select",
        select: "select",
        yesnona: "yesNoNA",
        checkbox: "checkbox",
        checkboxes: "checkbox",
        paragraph: "textarea",
        textarea: "textarea",
        rating: "rating",
        date: "date",
        time: "time",
        email: "email",
        url: "url",
        tel: "tel",
        location: "location",
        file: "file",
        fileupload: "file",
        "chassis-with-zone": "chassis-with-zone",
        chassiswithzone: "chassis-with-zone",
        "chassis-without-zone": "chassis-without-zone",
        chassiswithoutzone: "chassis-without-zone",
        "zone-in": "zone-in",
        zonein: "zone-in",
        "zone-out": "zone-out",
        zoneout: "zone-out"
      };
      return typeMap[normalized] || typeMap[type] || "text";
    };

    // Helper function to parse options string
    const parseOptions = (options: string | undefined): string[] => {
      if (!options || options.trim() === "") return [];
      return options.split(",").map((opt) => opt.trim());
    };

    // Helper function to safely get string values with fallback
    const getString = (
      value: string | undefined,
      fallback: string = ""
    ): string => {
      return value || fallback;
    };

    // Helper function to get boolean from string
    const getBoolean = (value: string | undefined): boolean => {
      return value?.toUpperCase() === "TRUE";
    };

    // Helper function to get number from string
    const getNumber = (
      value: string | undefined,
      fallback: number = 0
    ): number => {
      const num = parseInt(value || "");
      return isNaN(num) ? fallback : num;
    };

    // Parse the sampleData into form sections
    const sections: FormSection[] = [];
    let currentSection: FormSection | null = null;
    let currentQuestion: Question | null = null;

    sampleData.forEach((row, index) => {
      // Check if new section starts
      const sectionNumber = getString(row["Section Number"]);
      const sectionTitle = getString(row["Section Title"]);

      if (sectionNumber || (sectionTitle && !currentSection)) {
        // Save previous question if exists
        if (currentQuestion && currentSection) {
          currentSection.questions.push(currentQuestion);
          currentQuestion = null;
        }

        // Save previous section if exists
        if (currentSection) {
          sections.push(currentSection);
        }

        // Create new section
        currentSection = {
          id: crypto.randomUUID(),
          title:
            sectionTitle || `Section ${sectionNumber || sections.length + 1}`,
          description: getString(row["Section Description"]),
          questions: [],
        };
      }

      // Process main question
      const questionText = getString(row["Question"]);
      const questionType = getString(row["Question Type"]);

      if (questionText && questionType) {
        // Save previous question to section
        if (currentQuestion && currentSection) {
          currentSection.questions.push(currentQuestion);
        }

        // Create new question
        const questionId = crypto.randomUUID();
        currentQuestion = {
          id: questionId,
          text: questionText,
          type: mapQuestionType(questionType),
          required: getBoolean(row["Required"]),
          description: getString(row["Question Description"]),
          options: parseOptions(row["Options"]),
          subParam1: getString(row["SubParam1"] || row["Main Parameter"] || row["Sub Parameter 1"]),
          subParam2: getString(row["SubParam2"] || row["Followup Parameter"] || row["Sub Parameter 2"]),
          suggestion: "",
          followUpQuestions: [],
        };

        // Check if this question has follow-ups (looking for FU1: pattern)
        const followUps: FollowUpQuestion[] = [];

        // Find all follow-up question keys
        Object.keys(row).forEach((key) => {
          if (key.includes(": Question Text")) {
            const fuPrefix = key.split(": Question Text")[0];
            const triggerValue = getString(row[`${fuPrefix}: Option`]);

            if (triggerValue) {
              const followUpQuestion: FollowUpQuestion = {
                id: crypto.randomUUID(),
                text: getString(row[key]),
                type: mapQuestionType(
                  getString(row[`${fuPrefix}: Question Type`])
                ),
                required: getBoolean(row[`${fuPrefix}: Required`]),
                description: "",
                options: parseOptions(row[`${fuPrefix}: Options`]),
                subParam1: getString(row[`${fuPrefix}: SubParam1`]),
                subParam2: getString(row[`${fuPrefix}: SubParam2`]),
                parentId: questionId,
                showWhen: {
                  questionId: questionId,
                  value: triggerValue,
                },
                followUpQuestions: [],
              };

              // Check for nested follow-ups (FU1.1: pattern)
              Object.keys(row).forEach((nestedKey) => {
                if (
                  nestedKey.startsWith(`${fuPrefix}.`) &&
                  nestedKey.includes(": Question Text")
                ) {
                  const nestedPrefix = nestedKey.split(": Question Text")[0];
                  const nestedTriggerValue = getString(
                    row[`${nestedPrefix}: Option`]
                  );

                  if (nestedTriggerValue) {
                    const nestedFollowUp: FollowUpQuestion = {
                      id: crypto.randomUUID(),
                      text: getString(row[nestedKey]),
                      type: mapQuestionType(
                        getString(row[`${nestedPrefix}: Question Type`])
                      ),
                      required: getBoolean(row[`${nestedPrefix}: Required`]),
                      description: "",
                      options: parseOptions(row[`${nestedPrefix}: Options`]),
                      subParam1: getString(row[`${nestedPrefix}: SubParam1`]),
                      subParam2: getString(row[`${nestedPrefix}: SubParam2`]),
                      parentId: followUpQuestion.id,
                      showWhen: {
                        questionId: followUpQuestion.id,
                        value: nestedTriggerValue,
                      },
                      followUpQuestions: [],
                    };

                    // Check for deeper nesting (FU1.1.1: pattern)
                    Object.keys(row).forEach((deeperKey) => {
                      if (
                        deeperKey.startsWith(`${nestedPrefix}.`) &&
                        deeperKey.includes(": Question Text")
                      ) {
                        const deeperPrefix =
                          deeperKey.split(": Question Text")[0];
                        const deeperTriggerValue = getString(
                          row[`${deeperPrefix}: Option`]
                        );

                        if (deeperTriggerValue) {
                          const deeperFollowUp: FollowUpQuestion = {
                            id: crypto.randomUUID(),
                            text: getString(row[deeperKey]),
                            type: mapQuestionType(
                              getString(row[`${deeperPrefix}: Question Type`])
                            ),
                            required: getBoolean(
                              row[`${deeperPrefix}: Required`]
                            ),
                            description: "",
                            options: parseOptions(
                              row[`${deeperPrefix}: Options`]
                            ),
                            subParam1: getString(
                              row[`${deeperPrefix}: SubParam1`]
                            ),
                            subParam2: getString(
                              row[`${deeperPrefix}: SubParam2`]
                            ),
                            parentId: nestedFollowUp.id,
                            showWhen: {
                              questionId: nestedFollowUp.id,
                              value: deeperTriggerValue,
                            },
                            followUpQuestions: [],
                          };
                          nestedFollowUp.followUpQuestions!.push(
                            deeperFollowUp
                          );
                        }
                      }
                    });

                    followUpQuestion.followUpQuestions!.push(nestedFollowUp);
                  }
                }
              });

              followUps.push(followUpQuestion);
            }
          }
        });

        // Add follow-ups to question
        if (followUps.length > 0) {
          currentQuestion.followUpQuestions = followUps;
        }
      }
    });

    // Add the last question and section
    if (currentQuestion && currentSection) {
      currentSection.questions.push(currentQuestion);
    }
    if (currentSection) {
      sections.push(currentSection);
    }

    // Get form title and description from first row
    const formTitle = getString(
      sampleData[0]?.["Form Title"],
      "Bike Service & Maintenance Form"
    );
    const formDescription = getString(
      sampleData[0]?.["Form Description"],
      "Comprehensive bike service assessment with nested diagnostics"
    );

    // Create the form with proper typing
    const formToSet = {
      id: "",
      title: formTitle,
      description: formDescription,
      isVisible: true,
      locationEnabled: true,
      followUpQuestions: [],
      chassisNumbers: [],
      sections: (sections as any[]).map((section) => ({
        ...section,
        questions: section.questions.map((question) => {
          // Only add followUpConfig to main questions (not follow-ups)
          if (question.options && question.options.length > 0) {
            const followUpConfig: Record<
              string,
              { hasFollowUp: boolean; required: boolean; linkedFormId?: string }
            > = {};
            question.options.forEach((option) => {
              followUpConfig[option] = {
                hasFollowUp: false,
                required: false,
              };
            });

            // Check which options actually have follow-ups
            if (
              question.followUpQuestions &&
              question.followUpQuestions.length > 0
            ) {
              question.followUpQuestions.forEach((followUp) => {
                if (
                  followUp.showWhen?.value &&
                  followUpConfig[followUp.showWhen.value]
                ) {
                  followUpConfig[followUp.showWhen.value].hasFollowUp = true;
                  followUpConfig[followUp.showWhen.value].required =
                    followUp.required || false;
                }
              });
            }

            // Return the question with followUpConfig
            return {
              ...question,
              followUpConfig: followUpConfig,
            };
          }
          return question;
        }),
      })),
    };

    // Set the form state
    setForm(formToSet);

    showSuccess(
      "Bike Service Demo loaded! This includes nested follow-up questions for comprehensive diagnostics.",
      "Demo Data Loaded"
    );
  };

  const handleEditForm = (formId: string) => {
    navigate(`/forms/${formId}/edit`);
  };

  const handleViewResponses = (id: string) => {
    navigate(`/forms/${id}/responses`);
  };

  const handlePreviewForm = (id: string) => {
    navigate(`/forms/${id}/preview`);
  };

  const addSection = () => {
    setForm((prev) => ({
      ...prev,
      sections: [
        ...prev.sections,
        {
          id: crypto.randomUUID(),
          title: `Section ${prev.sections.length + 1}`,
          description: "",
          questions: [],
          isSubsection: false,
        },
      ],
    }));
  };

  const addSubsection = (parentSectionId: string) => {
    const parentIndex = form.sections.findIndex(
      (s) => s.id === parentSectionId
    );
    if (parentIndex === -1) return;

    // Count existing subsections for this parent
    const existingSubsections = form.sections.filter(
      (s) => s.parentSectionId === parentSectionId
    );

    const newSubsection: FormSection = {
      id: crypto.randomUUID(),
      title: `Subsection ${existingSubsections.length + 1}`,
      description: "",
      questions: [],
      isSubsection: true,
      parentSectionId: parentSectionId,
    };

    // Insert subsection after parent and its existing subsections
    const insertIndex = parentIndex + existingSubsections.length + 1;
    const updatedSections = [...form.sections];
    updatedSections.splice(insertIndex, 0, newSubsection);

    setForm((prev) => ({
      ...prev,
      sections: updatedSections,
    }));
  };

  // Group sections into pages (main sections = pages, subsections = same page as parent)
  const getPagesFromSections = (): FormSection[][] => {
    const pages: FormSection[][] = [];

    form.sections.forEach((section) => {
      if (!section.isSubsection) {
        const subsections = form.sections.filter(
          (s) => s.parentSectionId === section.id
        );
        pages.push([section, ...subsections]);
      }
    });

    return pages;
  };

  const pages = useMemo(() => getPagesFromSections(), [form.sections]);
  const totalPages = pages.length;

  useEffect(() => {
    if (totalPages === 0) {
      if (currentPage !== 0) {
        setCurrentPage(0);
      }
      setPageWindowStart(0);
      return;
    }

    if (currentPage > totalPages - 1) {
      setCurrentPage(totalPages - 1);
      return;
    }

    setPageWindowStart((prev) => {
      const maxStart = Math.max(0, totalPages - MAX_VISIBLE_PAGE_BUTTONS);
      let nextStart = prev;

      if (currentPage < nextStart) {
        nextStart = currentPage;
      } else if (currentPage >= nextStart + MAX_VISIBLE_PAGE_BUTTONS) {
        nextStart = currentPage - MAX_VISIBLE_PAGE_BUTTONS + 1;
      }

      if (nextStart > maxStart) {
        nextStart = maxStart;
      }

      if (nextStart < 0) {
        nextStart = 0;
      }

      return nextStart === prev ? prev : nextStart;
    });
  }, [currentPage, totalPages]);

  const handlePageChange = (pageIndex: number, total: number = totalPages) => {
    const nextTotal = Math.max(total, 0);

    if (nextTotal === 0) {
      setCurrentPage(0);
      setPageWindowStart(0);
      return;
    }

    const clampedPage = Math.max(0, Math.min(pageIndex, nextTotal - 1));
    setCurrentPage(clampedPage);

    setPageWindowStart((prev) => {
      const maxStart = Math.max(0, nextTotal - MAX_VISIBLE_PAGE_BUTTONS);
      let nextStart = prev;

      if (clampedPage < nextStart) {
        nextStart = clampedPage;
      } else if (clampedPage >= nextStart + MAX_VISIBLE_PAGE_BUTTONS) {
        nextStart = clampedPage - MAX_VISIBLE_PAGE_BUTTONS + 1;
      }

      if (nextStart > maxStart) {
        nextStart = maxStart;
      }

      if (nextStart < 0) {
        nextStart = 0;
      }

      return nextStart === prev ? prev : nextStart;
    });
  };

  const updateSection = (sectionId: string, updates: Partial<FormSection>) => {
    setForm((prev) => ({
      ...prev,
      sections: prev.sections.map((section) =>
        section.id === sectionId ? { ...section, ...updates } : section
      ),
    }));
  };

  const deleteSection = (sectionId: string) => {
    if (form.sections.length <= 1) {
      alert("Forms must have at least one section");
      return;
    }

    setForm((prev) => ({
      ...prev,
      sections: prev.sections.filter((section) => section.id !== sectionId),
    }));
  };

  const duplicateSection = (sectionId: string) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const duplicateQuestionRecursive = (
      q: Question | FollowUpQuestion
    ): Question | FollowUpQuestion => {
      return {
        ...q,
        id: crypto.randomUUID(),
        suggestion: q.suggestion || "",
        followUpQuestions: q.followUpQuestions?.map(
          (fq) =>
          ({
            ...duplicateQuestionRecursive(fq),
            parentId: crypto.randomUUID(),
          } as FollowUpQuestion)
        ),
      };
    };

    const duplicatedSection: FormSection = {
      ...section,
      id: crypto.randomUUID(),
      questions: section.questions.map(
        (q) => duplicateQuestionRecursive(q) as Question
      ),
    };

    const sectionIndex = form.sections.findIndex((s) => s.id === sectionId);
    const sections = [...form.sections];
    sections.splice(sectionIndex + 1, 0, duplicatedSection);

    setForm((prev) => ({
      ...prev,
      sections,
    }));
    showSuccess("Section duplicated successfully", "Success");
  };

  const addQuestion = (sectionId: string) => {
    const newQuestion: Question = {
      id: crypto.randomUUID(),
      text: "New Question",
      type: "text",
      required: false,
      description: "",
      imageUrl: "",
      subParam1: "",
      subParam2: "",
      suggestion: "",
    };

    updateSection(sectionId, {
      questions: [
        ...(form.sections.find((s) => s.id === sectionId)?.questions || []),
        newQuestion,
      ],
    });
  };

  // Insert question at specific position
  const insertQuestionAt = (sectionId: string, index: number) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const newQuestion: Question = {
      id: crypto.randomUUID(),
      text: "New Question",
      type: "text",
      required: false,
      description: "",
      imageUrl: "",
      subParam1: "",
      subParam2: "",
      suggestion: "",
    };

    const questions = [...section.questions];
    questions.splice(index, 0, newQuestion);

    updateSection(sectionId, { questions });
  };

  // Duplicate question
  const duplicateQuestion = (sectionId: string, questionId: string) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const question = section.questions.find((q) => q.id === questionId);
    if (!question) return;

    // Deep clone the question with new IDs
    const duplicateQuestionRecursive = (
      q: Question | FollowUpQuestion
    ): Question | FollowUpQuestion => {
      return {
        ...q,
        id: crypto.randomUUID(),
        suggestion: q.suggestion || "",
        followUpQuestions: q.followUpQuestions?.map(
          (fq) =>
          ({
            ...duplicateQuestionRecursive(fq),
            parentId: crypto.randomUUID(), // Will be updated later
          } as FollowUpQuestion)
        ),
      };
    };

    const duplicatedQuestion = duplicateQuestionRecursive(question) as Question;

    const questionIndex = section.questions.findIndex(
      (q) => q.id === questionId
    );
    const questions = [...section.questions];
    questions.splice(questionIndex + 1, 0, duplicatedQuestion);

    updateSection(sectionId, { questions });
    showSuccess("Question duplicated successfully", "Success");
  };

  // Move question up
  const moveQuestionUp = (sectionId: string, questionId: string) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const questionIndex = section.questions.findIndex(
      (q) => q.id === questionId
    );
    if (questionIndex <= 0) return; // Already at top

    const questions = [...section.questions];
    [questions[questionIndex - 1], questions[questionIndex]] = [
      questions[questionIndex],
      questions[questionIndex - 1],
    ];

    updateSection(sectionId, { questions });
  };

  // Move question down
  const moveQuestionDown = (sectionId: string, questionId: string) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const questionIndex = section.questions.findIndex(
      (q) => q.id === questionId
    );
    if (questionIndex === -1 || questionIndex >= section.questions.length - 1)
      return; // Already at bottom

    const questions = [...section.questions];
    [questions[questionIndex], questions[questionIndex + 1]] = [
      questions[questionIndex + 1],
      questions[questionIndex],
    ];

    updateSection(sectionId, { questions });
  };

  // Split section at a specific question
  const splitSectionAtQuestion = (sectionId: string, questionId: string) => {
    const sectionIndex = form.sections.findIndex((s) => s.id === sectionId);
    if (sectionIndex === -1) return;

    const section = form.sections[sectionIndex];
    const questionIndex = section.questions.findIndex(
      (q) => q.id === questionId
    );

    if (questionIndex === -1) return;

    // Split the questions: questions before the current one stay in the current section,
    // questions from the current one onwards move to the new section.
    const questionsForCurrentSection = section.questions.slice(0, questionIndex);
    const questionsForNewSection = section.questions.slice(questionIndex);

    if (questionsForNewSection.length === 0) return;

    // Create the new section
    const newSection: FormSection = {
      id: crypto.randomUUID(),
      title: `Section ${form.sections.length + 1}`,
      description: "",
      questions: questionsForNewSection,
      isSubsection: false,
    };

    // Update the sections list
    const updatedSections = [...form.sections];
    // Update the current section's questions
    updatedSections[sectionIndex] = {
      ...section,
      questions: questionsForCurrentSection,
    };
    // Insert the new section after the current one
    updatedSections.splice(sectionIndex + 1, 0, newSection);

    setForm((prev) => ({
      ...prev,
      sections: updatedSections,
    }));

    showSuccess("Section split successfully", "Success");
  };

  // Move section up
  const moveSectionUp = (sectionId: string) => {
    const sectionIndex = form.sections.findIndex((s) => s.id === sectionId);
    if (sectionIndex <= 0) return;

    const updatedSections = [...form.sections];
    [updatedSections[sectionIndex - 1], updatedSections[sectionIndex]] = [
      updatedSections[sectionIndex],
      updatedSections[sectionIndex - 1],
    ];

    setForm((prev) => ({
      ...prev,
      sections: updatedSections,
    }));
  };

  // Move section down
  const moveSectionDown = (sectionId: string) => {
    const sectionIndex = form.sections.findIndex((s) => s.id === sectionId);
    if (sectionIndex === -1 || sectionIndex >= form.sections.length - 1) return;

    const updatedSections = [...form.sections];
    [updatedSections[sectionIndex], updatedSections[sectionIndex + 1]] = [
      updatedSections[sectionIndex + 1],
      updatedSections[sectionIndex],
    ];

    setForm((prev) => ({
      ...prev,
      sections: updatedSections,
    }));
  };

  const updateQuestion = (
    sectionId: string,
    questionId: string,
    updates: Partial<Question>
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    // Helper to update question recursively in a list (main or follow-up)
    const updateInList = (list: any[]): any[] => {
      return list.map((q) => {
        if (q.id === questionId) {
          return { ...q, ...updates };
        }
        if (q.followUpQuestions && q.followUpQuestions.length > 0) {
          return {
            ...q,
            followUpQuestions: updateInList(q.followUpQuestions),
          };
        }
        return q;
      });
    };

    updateSection(sectionId, {
      questions: updateInList(section.questions),
    });
  };

  const MAX_QUESTION_IMAGE_BYTES = 50 * 1024;

  const compressQuestionImage = async (file: File): Promise<string> => {
    const readFileAsDataUrl = () =>
      new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = () => reject(new Error("Failed to read image"));
        reader.readAsDataURL(file);
      });

    const loadImageElement = (dataUrl: string) =>
      new Promise<HTMLImageElement>((resolve, reject) => {
        const imageElement = new Image();
        imageElement.onload = () => resolve(imageElement);
        imageElement.onerror = () => reject(new Error("Failed to load image"));
        imageElement.src = dataUrl;
      });

    const dataUrl = await readFileAsDataUrl();
    const imageElement = await loadImageElement(dataUrl);
    const canvas = document.createElement("canvas");

    if (typeof canvas.toBlob !== "function") {
      if (file.size <= MAX_QUESTION_IMAGE_BYTES) {
        return dataUrl;
      }
      throw new Error(
        "Unable to process image. Please upload a smaller file under 50KB."
      );
    }

    if (!canvas.getContext("2d")) {
      if (file.size <= MAX_QUESTION_IMAGE_BYTES) {
        return dataUrl;
      }
      throw new Error(
        "Unable to process image. Please upload a smaller file under 50KB."
      );
    }

    const drawImage = (width: number, height: number) => {
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d");
      if (!context) {
        throw new Error("Canvas not supported");
      }
      context.drawImage(imageElement, 0, 0, width, height);
    };

    const createBlob = async (quality: number) =>
      new Promise<Blob | null>((resolve) =>
        canvas.toBlob((blob) => resolve(blob), "image/jpeg", quality)
      );

    const maxDimension = 1024;
    const initialScale = Math.min(
      1,
      maxDimension / imageElement.width,
      maxDimension / imageElement.height
    );

    let width = Math.max(1, Math.round(imageElement.width * initialScale));
    let height = Math.max(1, Math.round(imageElement.height * initialScale));
    drawImage(width, height);

    let quality = 0.9;
    let blob = await createBlob(quality);
    if (!blob) {
      if (file.size <= MAX_QUESTION_IMAGE_BYTES) {
        return dataUrl;
      }
      throw new Error("Unable to process image");
    }

    const minQuality = 0.2;

    while (blob.size > MAX_QUESTION_IMAGE_BYTES && quality > minQuality) {
      quality = Math.max(minQuality, quality - 0.1);
      const nextBlob = await createBlob(quality);
      if (!nextBlob) {
        break;
      }
      blob = nextBlob;
    }

    while (
      blob.size > MAX_QUESTION_IMAGE_BYTES &&
      (width > 120 || height > 120)
    ) {
      width = Math.max(120, Math.floor(width * 0.85));
      height = Math.max(120, Math.floor(height * 0.85));
      drawImage(width, height);
      quality = 0.9;
      let nextBlob = await createBlob(quality);
      if (!nextBlob) {
        break;
      }
      blob = nextBlob;
      while (blob.size > MAX_QUESTION_IMAGE_BYTES && quality > minQuality) {
        quality = Math.max(minQuality, quality - 0.1);
        const smallerBlob = await createBlob(quality);
        if (!smallerBlob) {
          break;
        }
        blob = smallerBlob;
      }
    }

    if (blob.size > MAX_QUESTION_IMAGE_BYTES) {
      throw new Error(
        "Unable to compress image below 50KB. Try a smaller image."
      );
    }

    return await new Promise<string>((resolve, reject) => {
      const resultReader = new FileReader();
      resultReader.onload = () => resolve(resultReader.result as string);
      resultReader.onerror = () =>
        reject(new Error("Failed to read compressed image"));
      resultReader.readAsDataURL(blob);
    });
  };

  const handleQuestionImageUpload = async (
    sectionId: string,
    questionId: string,
    file: File
  ) => {
    try {
      const isImage = file.type.startsWith("image/");
      const isSpecialFile = file.name.toLowerCase().endsWith(".stp") || 
                           file.name.toLowerCase().endsWith(".step") || 
                           file.name.toLowerCase().endsWith(".pvz");

      if (isSpecialFile) {
        // For special files, use standard file upload
        const result = await apiClient.uploadFile(file, "form", questionId);
        const uploadedUrl = apiClient.resolveUploadedFileUrl(result);
        if (uploadedUrl) {
          updateQuestion(sectionId, questionId, { imageUrl: uploadedUrl });
          showSuccess(`${file.name.split('.').pop()?.toUpperCase()} file uploaded successfully`);
        }
        return;
      }

      if (!isImage) {
        showError("Please upload an image, STP, or PVZ file", "Invalid File Type");
        return;
      }

      const compressed = await compressQuestionImage(file);
      updateQuestion(sectionId, questionId, { imageUrl: compressed });
    } catch (error) {
      console.error("Image compression failed:", error);
      showError(
        error instanceof Error ? error.message : "Failed to process image",
        "Image Upload Error"
      );
    }
  };

  const clearQuestionImage = (sectionId: string, questionId: string) => {
    updateQuestion(sectionId, questionId, { imageUrl: undefined });
  };

  const deleteQuestion = (sectionId: string, questionId: string) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    updateSection(sectionId, {
      questions: section.questions.filter((q) => q.id !== questionId),
    });
  };

  const openBranchingConfig = (
    sectionId: string,
    questionId: string,
    options: string[]
  ) => {
    setBranchingConfigQuestion({ questionId, sectionId, options });
    setShowBranchingConfig(true);
  };

  const openFormRoutingConfig = (
    sectionId: string,
    questionId: string,
    options: string[]
  ) => {
    setFormRoutingConfigQuestion({ questionId, sectionId, options });
    setShowFormRoutingConfig(true);
  };

  const handleUpdateSection = (sectionId: string, updates: Partial<FormSection>) => {
    updateSection(sectionId, updates);
  };

  const handleSaveBranchingRules = (rules: any[]) => {
    if (!branchingConfigQuestion) return;

    const section = form.sections.find(
      (s) => s.id === branchingConfigQuestion.sectionId
    );
    if (!section) return;

    const question = section.questions.find(
      (q) => q.id === branchingConfigQuestion.questionId
    );
    if (!question) return;

    const branchingRules = rules.map((rule) => ({
      optionLabel: rule.optionLabel,
      targetSectionId: rule.targetSectionId,
      isOtherOption: rule.isOtherOption || false,
    }));

    const updatedQuestion = {
      ...question,
      branchingRules: branchingRules,
    };

    updateSection(branchingConfigQuestion.sectionId, {
      questions: section.questions.map((q) =>
        q.id === branchingConfigQuestion.questionId ? updatedQuestion : q
      ),
    });

    // Update formSectionBranching state with complete rule structure for API
    const apiRules = rules.map((rule) => ({
      questionId: branchingConfigQuestion.questionId,
      sectionId: branchingConfigQuestion.sectionId,
      optionLabel: rule.optionLabel,
      targetSectionId: rule.targetSectionId,
      isOtherOption: rule.isOtherOption || false,
    }));

    // Remove old rules for this question and add new ones
    const updatedBranching = formSectionBranching.filter(
      (r) =>
        !(
          r.questionId === branchingConfigQuestion.questionId &&
          r.sectionId === branchingConfigQuestion.sectionId
        )
    );
    setFormSectionBranching([...updatedBranching, ...apiRules]);
    console.log("Branching rules updated:", [...updatedBranching, ...apiRules]);

    setShowBranchingConfig(false);
    setBranchingConfigQuestion(null);
    showSuccess("Section routing configured successfully");
  };

  const hasSectionBranching = (sectionId: string) => {
    // Check if any rule in formSectionBranching exists for this section
    return formSectionBranching.some((rule) => rule.sectionId === sectionId);
  };

  const handleSaveFormRoutingConfig = (
    config: Record<string, { linkedFormId: string }>
  ) => {
    if (!formRoutingConfigQuestion) return;

    const section = form.sections.find(
      (s) => s.id === formRoutingConfigQuestion.sectionId
    );
    if (!section) return;

    const question = section.questions.find(
      (q) => q.id === formRoutingConfigQuestion.questionId
    );
    if (!question) return;

    const updatedQuestion = {
      ...question,
      followUpConfig: {
        ...(question.followUpConfig || {}),
        ...config,
      },
    };

    updateSection(formRoutingConfigQuestion.sectionId, {
      questions: section.questions.map((q) =>
        q.id === formRoutingConfigQuestion.questionId ? updatedQuestion : q
      ),
    });

    setShowFormRoutingConfig(false);
    setFormRoutingConfigQuestion(null);
    showSuccess("Follow-up form routing configured successfully");
  };

  const linkFollowUpSection = (
    sectionId: string,
    parentQuestionId: string,
    option: string
  ) => {
    // Show section selector modal
    setPendingSectionLink({
      sectionId,
      parentQuestionId,
      triggerValue: option,
      path: [],
    });
    setShowSectionSelector(true);
  };

  const linkFollowUpForm = (
    sectionId: string,
    parentQuestionId: string,
    option: string,
    linkedFormId?: string
  ) => {
    if (linkedFormId) {
      // Direct form linking from modal
      const section = form.sections.find((s) => s.id === sectionId);
      if (!section) return;

      const parentQuestion = section.questions.find(
        (q) => q.id === parentQuestionId
      );
      if (!parentQuestion) return;

      const existingConfig = parentQuestion.followUpConfig?.[option] || {
        hasFollowUp: true,
        required: false,
      };

      const updatedParentQuestion = {
        ...parentQuestion,
        followUpConfig: {
          ...parentQuestion.followUpConfig,
          [option]: {
            ...existingConfig,
            linkedFormId: linkedFormId,
          },
        },
      };

      updateSection(sectionId, {
        questions: section.questions.map((q) =>
          q.id === parentQuestionId ? updatedParentQuestion : q
        ),
      });
      showSuccess(`Form linked successfully for "${option}"`);
    } else {
      // Show form selector modal
      setPendingFormLink({
        sectionId,
        parentQuestionId,
        triggerValue: option,
        path: [],
      });
      setShowFormSelector(true);
    }
  };

  const addFollowUpQuestion = (
    sectionId: string,
    parentQuestionId: string,
    triggerValue: string
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const parentQuestion = section.questions.find(
      (q) => q.id === parentQuestionId
    );
    if (!parentQuestion) return;

    const newFollowUpQuestion: FollowUpQuestion = {
      id: crypto.randomUUID(),
      text: "Follow-up Question",
      type: "text",
      required: false,
      description: "",
      imageUrl: "",
      subParam1: "",
      subParam2: "",
      parentId: parentQuestionId,
      showWhen: {
        questionId: parentQuestionId,
        value: triggerValue,
      },
    };

    const followUpConfig = {
      ...(parentQuestion.followUpConfig || {}),
    };

    if (parentQuestion.options && parentQuestion.options.length > 0) {
      parentQuestion.options.forEach((option) => {
        if (!followUpConfig[option]) {
          followUpConfig[option] = {
            hasFollowUp: false,
            required: false,
          };
        }
      });
    }

    if (!followUpConfig[triggerValue]) {
      followUpConfig[triggerValue] = {
        hasFollowUp: true,
        required: newFollowUpQuestion.required || false,
      };
    } else {
      followUpConfig[triggerValue] = {
        ...followUpConfig[triggerValue],
        hasFollowUp: true,
        required:
          newFollowUpQuestion.required ?? followUpConfig[triggerValue].required,
      };
    }

    const updatedParentQuestion = {
      ...parentQuestion,
      followUpQuestions: [
        ...(parentQuestion.followUpQuestions || []),
        newFollowUpQuestion,
      ],
      followUpConfig,
    };

    updateSection(sectionId, {
      questions: section.questions.map((q) =>
        q.id === parentQuestionId ? updatedParentQuestion : q
      ),
    });
  };

  const updateFollowUpQuestion = (
    sectionId: string,
    parentQuestionId: string,
    followUpQuestionId: string,
    updates: Partial<FollowUpQuestion>
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const parentQuestion = section.questions.find(
      (q) => q.id === parentQuestionId
    );
    if (!parentQuestion) return;

    const updatedParentQuestion = {
      ...parentQuestion,
      followUpQuestions: (parentQuestion.followUpQuestions || []).map((fq) =>
        fq.id === followUpQuestionId ? { ...fq, ...updates } : fq
      ),
    };

    updateSection(sectionId, {
      questions: section.questions.map((q) =>
        q.id === parentQuestionId ? updatedParentQuestion : q
      ),
    });
  };

  const deleteFollowUpQuestion = (
    sectionId: string,
    parentQuestionId: string,
    followUpQuestionId: string
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const parentQuestion = section.questions.find(
      (q) => q.id === parentQuestionId
    );
    if (!parentQuestion) return;

    const updatedParentQuestion = {
      ...parentQuestion,
      followUpQuestions: (parentQuestion.followUpQuestions || []).filter(
        (fq) => fq.id !== followUpQuestionId
      ),
    };

    updateSection(sectionId, {
      questions: section.questions.map((q) =>
        q.id === parentQuestionId ? updatedParentQuestion : q
      ),
    });
  };

  // Recursive function to add nested follow-up question
  const addNestedFollowUpQuestion = (
    sectionId: string,
    parentQuestionId: string,
    triggerValue: string,
    path: string[] = [] // Path to the parent question in the nested structure
  ) => {
    const newFollowUpQuestion: FollowUpQuestion = {
      id: crypto.randomUUID(),
      text: "Follow-up Question",
      type: "text",
      required: false,
      description: "",
      imageUrl: "",
      parentId: parentQuestionId,
      showWhen: {
        questionId: parentQuestionId,
        value: triggerValue,
      },
      followUpQuestions: [],
    };

    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) {
      console.error("Section not found:", sectionId);
      return;
    }

    // Helper function to recursively update nested follow-ups
    // The path includes the full path TO the target question (including the target itself)
    const updateNestedFollowUps = (
      questions: (Question | FollowUpQuestion)[],
      pathIndex: number = 0
    ): (Question | FollowUpQuestion)[] => {
      return questions.map((q) => {
        // Check if this is the question we're looking for in the path
        if (pathIndex < path.length && q.id === path[pathIndex]) {
          // Check if this is the last element in the path (the target question)
          if (pathIndex === path.length - 1) {
            // This is the target question, add the follow-up
            console.log(
              "Adding follow-up to question:",
              q.id,
              "at path index:",
              pathIndex
            );
            return {
              ...q,
              followUpQuestions: [
                ...(q.followUpQuestions || []),
                newFollowUpQuestion,
              ],
            };
          } else {
            // Continue traversing the path through follow-ups
            console.log(
              "Traversing through question:",
              q.id,
              "at path index:",
              pathIndex
            );
            return {
              ...q,
              followUpQuestions: updateNestedFollowUps(
                q.followUpQuestions || [],
                pathIndex + 1
              ) as FollowUpQuestion[],
            };
          }
        }

        return q;
      });
    };

    console.log("addNestedFollowUpQuestion called with:", {
      sectionId,
      parentQuestionId,
      triggerValue,
      path,
    });

    const updatedQuestions = updateNestedFollowUps(section.questions, 0);
    console.log("Updated questions:", updatedQuestions);

    updateSection(sectionId, {
      questions: updatedQuestions as Question[],
    });
  };

  const handleAddFollowUpSection = (
    sectionId: string,
    parentQuestionId: string,
    triggerValue: string,
    path: string[]
  ) => {
    setPendingSectionLink({
      sectionId,
      parentQuestionId,
      triggerValue,
      path,
    });
    setShowSectionSelector(true);
  };

  const handleSelectSectionForLink = (selectedSectionId: string) => {
    if (!pendingSectionLink) return;

    const section = form.sections.find(
      (s) => s.id === pendingSectionLink.sectionId
    );
    if (!section) {
      showError("Section not found");
      return;
    }

    const updateFollowUpConfig = (
      questions: (Question | FollowUpQuestion)[],
      pathIndex: number = 0
    ): (Question | FollowUpQuestion)[] => {
      return questions.map((q) => {
        if (
          pathIndex < pendingSectionLink.path.length &&
          q.id === pendingSectionLink.path[pathIndex]
        ) {
          if (pathIndex === pendingSectionLink.path.length - 1) {
            return {
              ...q,
              followUpConfig: {
                ...(q.followUpConfig || {}),
                [pendingSectionLink.triggerValue]: {
                  hasFollowUp: true,
                  required: false,
                  linkedSectionId: selectedSectionId,
                },
              },
            };
          } else {
            return {
              ...q,
              followUpQuestions: updateFollowUpConfig(
                q.followUpQuestions || [],
                pathIndex + 1
              ),
            };
          }
        }
        return q;
      });
    };

    const updatedQuestions = updateFollowUpConfig(section.questions, 0);
    updateSection(pendingSectionLink.sectionId, {
      questions: updatedQuestions as Question[],
    });

    setShowSectionSelector(false);
    setPendingSectionLink(null);
    showSuccess(`Section linked successfully`);
  };

  const handleAddFollowUpForm = (
    sectionId: string,
    parentQuestionId: string,
    triggerValue: string,
    path: string[]
  ) => {
    setPendingFormLink({
      sectionId,
      parentQuestionId,
      triggerValue,
      path,
    });
    setShowFormSelector(true);
  };

  const handleSelectFormForLink = (selectedFormId: string) => {
    if (!pendingFormLink) return;

    const section = form.sections.find(
      (s) => s.id === pendingFormLink.sectionId
    );
    if (!section) {
      showError("Section not found");
      return;
    }

    const updateFollowUpConfig = (
      questions: (Question | FollowUpQuestion)[],
      pathIndex: number = 0
    ): (Question | FollowUpQuestion)[] => {
      return questions.map((q) => {
        if (
          pathIndex < pendingFormLink.path.length &&
          q.id === pendingFormLink.path[pathIndex]
        ) {
          if (pathIndex === pendingFormLink.path.length - 1) {
            return {
              ...q,
              followUpConfig: {
                ...(q.followUpConfig || {}),
                [pendingFormLink.triggerValue]: {
                  hasFollowUp: true,
                  required: false,
                  linkedFormId: selectedFormId,
                },
              },
            };
          } else {
            return {
              ...q,
              followUpQuestions: updateFollowUpConfig(
                q.followUpQuestions || [],
                pathIndex + 1
              ),
            };
          }
        }
        return q;
      });
    };

    const updatedQuestions = updateFollowUpConfig(section.questions, 0);
    updateSection(pendingFormLink.sectionId, {
      questions: updatedQuestions as Question[],
    });

    setShowFormSelector(false);
    setPendingFormLink(null);
    showSuccess(`Form linked successfully`);
  };

  // Recursive function to update nested follow-up question
  const updateNestedFollowUpQuestion = (
    sectionId: string,
    followUpQuestionId: string,
    updates: Partial<FollowUpQuestion>,
    path: string[] = [] // Path to the follow-up question
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const updateInQuestions = (
      questions: (Question | FollowUpQuestion)[],
      pathIndex: number = 0
    ): (Question | FollowUpQuestion)[] => {
      return questions.map((q) => {
        if (pathIndex === path.length && q.id === followUpQuestionId) {
          return { ...q, ...updates };
        } else if (pathIndex < path.length && q.id === path[pathIndex]) {
          return {
            ...q,
            followUpQuestions: updateInQuestions(
              q.followUpQuestions || [],
              pathIndex + 1
            ) as FollowUpQuestion[],
          };
        } else if (q.followUpQuestions && q.followUpQuestions.length > 0) {
          const updated = updateInQuestions(q.followUpQuestions, pathIndex);
          if (updated !== q.followUpQuestions) {
            return { ...q, followUpQuestions: updated as FollowUpQuestion[] };
          }
        }
        return q;
      });
    };

    updateSection(sectionId, {
      questions: updateInQuestions(section.questions) as Question[],
    });
  };

  const handleNestedFollowUpImageUpload = async (
    sectionId: string,
    followUpQuestionId: string,
    file: File,
    path: string[]
  ) => {
    try {
      const compressed = await compressQuestionImage(file);
      updateNestedFollowUpQuestion(
        sectionId,
        followUpQuestionId,
        { imageUrl: compressed },
        path
      );
    } catch (error) {
      console.error("Nested image compression failed:", error);
      showError(
        error instanceof Error ? error.message : "Failed to process image",
        "Image Upload Error"
      );
    }
  };

  const clearNestedFollowUpImage = (
    sectionId: string,
    followUpQuestionId: string,
    path: string[]
  ) => {
    updateNestedFollowUpQuestion(
      sectionId,
      followUpQuestionId,
      { imageUrl: undefined },
      path
    );
  };

  // Recursive function to delete nested follow-up question
  const deleteNestedFollowUpQuestion = (
    sectionId: string,
    followUpQuestionId: string,
    path: string[] = [] // Path to the parent of the follow-up question
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const deleteInQuestions = (
      questions: (Question | FollowUpQuestion)[],
      pathIndex: number = 0
    ): (Question | FollowUpQuestion)[] => {
      return questions.map((q) => {
        if (pathIndex === path.length) {
          // At the parent level, filter out the target follow-up
          return {
            ...q,
            followUpQuestions: (q.followUpQuestions || []).filter(
              (fq) => fq.id !== followUpQuestionId
            ),
          };
        } else if (pathIndex < path.length && q.id === path[pathIndex]) {
          return {
            ...q,
            followUpQuestions: deleteInQuestions(
              q.followUpQuestions || [],
              pathIndex + 1
            ) as FollowUpQuestion[],
          };
        } else if (q.followUpQuestions && q.followUpQuestions.length > 0) {
          const updated = deleteInQuestions(q.followUpQuestions, pathIndex);
          if (updated !== q.followUpQuestions) {
            return { ...q, followUpQuestions: updated as FollowUpQuestion[] };
          }
        }
        return q;
      });
    };

    updateSection(sectionId, {
      questions: deleteInQuestions(section.questions) as Question[],
    });
  };

  const addOption = (sectionId: string, questionId: string) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const question = section.questions.find((q) => q.id === questionId);
    if (!question) return;
    if (question.type === "yesNoNA" || question.type === "chassisNumber") return;

    updateQuestion(sectionId, questionId, {
      options: [...(question.options || []), ""],
    });
  };

  const updateOption = (
    sectionId: string,
    questionId: string,
    optionIndex: number,
    value: string
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const question = section.questions.find((q) => q.id === questionId);
    if (!question || !question.options) return;

    const updatedOptions = [...question.options];
    updatedOptions[optionIndex] = value;

    updateQuestion(sectionId, questionId, {
      options: updatedOptions,
    });
  };

  const removeOption = (
    sectionId: string,
    questionId: string,
    optionIndex: number
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const question = section.questions.find((q) => q.id === questionId);
    if (!question || !question.options) return;
    if (question.type === "yesNoNA" || question.type === "chassisNumber") return;

    updateQuestion(sectionId, questionId, {
      options: question.options.filter((_, i) => i !== optionIndex),
    });
  };

  // Duplicate option
  const duplicateOption = (
    sectionId: string,
    questionId: string,
    optionIndex: number
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const question = section.questions.find((q) => q.id === questionId);
    if (!question || !question.options) return;

    const optionToDuplicate = question.options[optionIndex];
    const newOptions = [...question.options];
    newOptions.splice(optionIndex + 1, 0, `${optionToDuplicate} (Copy)`);

    updateQuestion(sectionId, questionId, {
      options: newOptions,
    });
  };

  const addFollowUpOption = (
    sectionId: string,
    parentQuestionId: string,
    followUpQuestionId: string
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const parentQuestion = section.questions.find(
      (q) => q.id === parentQuestionId
    );
    if (!parentQuestion) return;

    const followUpQ = parentQuestion.followUpQuestions?.find(
      (fq) => fq.id === followUpQuestionId
    );
    if (!followUpQ) return;

    const updatedFollowUpQ = {
      ...followUpQ,
      options: [...(followUpQ.options || []), ""],
    };

    updateFollowUpQuestion(
      sectionId,
      parentQuestionId,
      followUpQuestionId,
      updatedFollowUpQ
    );
  };

  const updateFollowUpOption = (
    sectionId: string,
    parentQuestionId: string,
    followUpQuestionId: string,
    optionIndex: number,
    value: string
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const parentQuestion = section.questions.find(
      (q) => q.id === parentQuestionId
    );
    if (!parentQuestion) return;

    const followUpQ = parentQuestion.followUpQuestions?.find(
      (fq) => fq.id === followUpQuestionId
    );
    if (!followUpQ || !followUpQ.options) return;

    const updatedOptions = [...followUpQ.options];
    updatedOptions[optionIndex] = value;

    updateFollowUpQuestion(sectionId, parentQuestionId, followUpQuestionId, {
      options: updatedOptions.filter((opt) => opt.trim() !== ""),
    });
  };

  const removeFollowUpOption = (
    sectionId: string,
    parentQuestionId: string,
    followUpQuestionId: string,
    optionIndex: number
  ) => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return;

    const parentQuestion = section.questions.find(
      (q) => q.id === parentQuestionId
    );
    if (!parentQuestion) return;

    const followUpQ = parentQuestion.followUpQuestions?.find(
      (fq) => fq.id === followUpQuestionId
    );
    if (!followUpQ || !followUpQ.options) return;

    updateFollowUpQuestion(sectionId, parentQuestionId, followUpQuestionId, {
      options: followUpQ.options.filter((_, i) => i !== optionIndex),
    });
  };

  // Recursive functions for nested follow-up options
  const addNestedFollowUpOption = (
    sectionId: string,
    followUpQuestionId: string,
    path: string[] = []
  ) => {
    updateNestedFollowUpQuestion(
      sectionId,
      followUpQuestionId,
      {
        options: [
          ...((path.length > 0
            ? findNestedQuestion(sectionId, followUpQuestionId, path)?.options
            : null) || []),
          "",
        ],
      },
      path
    );
  };

  const updateNestedFollowUpOption = (
    sectionId: string,
    followUpQuestionId: string,
    optionIndex: number,
    value: string,
    path: string[] = []
  ) => {
    const question = findNestedQuestion(sectionId, followUpQuestionId, path);
    if (!question || !question.options) return;

    const updatedOptions = [...question.options];
    updatedOptions[optionIndex] = value;

    updateNestedFollowUpQuestion(
      sectionId,
      followUpQuestionId,
      {
        options: updatedOptions.filter((opt) => opt.trim() !== ""),
      },
      path
    );
  };

  const removeNestedFollowUpOption = (
    sectionId: string,
    followUpQuestionId: string,
    optionIndex: number,
    path: string[] = []
  ) => {
    const question = findNestedQuestion(sectionId, followUpQuestionId, path);
    if (!question || !question.options) return;

    updateNestedFollowUpQuestion(
      sectionId,
      followUpQuestionId,
      {
        options: question.options.filter((_, i) => i !== optionIndex),
      },
      path
    );
  };

  // Helper function to find a nested question by path
  const findNestedQuestion = (
    sectionId: string,
    questionId: string,
    path: string[] = []
  ): Question | FollowUpQuestion | null => {
    const section = form.sections.find((s) => s.id === sectionId);
    if (!section) return null;

    const searchInQuestions = (
      questions: (Question | FollowUpQuestion)[],
      pathIndex: number = 0
    ): Question | FollowUpQuestion | null => {
      for (const q of questions) {
        if (pathIndex === path.length && q.id === questionId) {
          return q;
        } else if (pathIndex < path.length && q.id === path[pathIndex]) {
          return searchInQuestions(q.followUpQuestions || [], pathIndex + 1);
        } else if (q.followUpQuestions && q.followUpQuestions.length > 0) {
          const found = searchInQuestions(q.followUpQuestions, pathIndex);
          if (found) return found;
        }
      }
      return null;
    };

    return searchInQuestions(section.questions);
  };

  // Helper function to check if a question type requires follow-ups
  const requiresFollowUp = (type: string): boolean => {
    return ["radio", "checkbox", "select", "search-select", "yesNoNA", "chassisNumber"].includes(
      type
    );
  };

  const handleQuestionTypeChange = (
    sectionId: string,
    questionId: string,
    currentQuestion: Question,
    newType: string
  ) => {
    const updates: Partial<Question> = { type: newType };

    if (newType === "yesNoNA" || newType === "chassisNumber") {
      updates.options = [...YES_NO_NA_OPTIONS];
      updates.correctAnswer = YES_NO_NA_CORRECT;
      } else if (newType === "chassis-with-zone" || newType === "chassis-without-zone") {
      updates.options = undefined;
      updates.correctAnswer = undefined;
    } else if (requiresFollowUp(newType)) {
      const baseOptions =
        currentQuestion.type === "yesNoNA"
          ? undefined
          : currentQuestion.options;
      updates.options =
        baseOptions && baseOptions.length > 0
          ? baseOptions
          : ["Option 1", "Option 2"];
      if (currentQuestion.type === "yesNoNA") {
        updates.correctAnswer = undefined;
      }
    } else {
      updates.options = undefined;
      updates.correctAnswer = undefined;
      updates.followUpConfig = undefined;
    }

    if (newType === "file") {
      updates.allowedFileTypes =
        currentQuestion.allowedFileTypes &&
          currentQuestion.allowedFileTypes.length > 0
          ? [...currentQuestion.allowedFileTypes]
          : ["image", "pdf", "excel"];
    } else if (currentQuestion.allowedFileTypes) {
      updates.allowedFileTypes = undefined;
    }

    updateQuestion(sectionId, questionId, updates);
  };

  const questionTypes = [
    {
      value: "text",
      label: "Short Text",
      description: "Single line text input",
    },
    {
      value: "textarea",
      label: "Long Text",
      description: "Multi-line text area",
    },
    {
      value: "radio",
      label: "Multiple Choice",
      description: "Select one option from many",
    },
    {
      value: "yesNoNA",
      label: "Yes / No / N/A",
      description: "Preset with Yes, No, and N/A options",
    },
    {
      value: "checkbox",
      label: "Checkboxes",
      description: "Select multiple options",
    },
    {
      value: "select",
      label: "Dropdown",
      description: "Choose from dropdown list",
    },
    {
      value: "search-select",
      label: "Search/Filter Dropdown",
      description: "Dropdown with search and filter functionality",
    },
    { value: "email", label: "Email", description: "Email address input" },
    { value: "number", label: "Number", description: "Numeric input only" },
    { value: "date", label: "Date", description: "Date picker" },
    {
      value: "file",
      label: "File Upload",
      description: "Upload files/documents",
    },
    {
      value: "slider-feedback",
      label: "Slider Feedback (1-10)",
      description: "Drag slider with emoji feedback",
    },
    {
      value: "emoji-star-feedback",
      label: "Star Rating Feedback",
      description: "5-star emoji rating",
    },
    {
      value: "emoji-reaction-feedback",
      label: "Emoji Reaction Feedback",
      description: "Emoji reactions (sad to laugh)",
    },
    {
      value: "chassisNumber",
      label: "Chassis Number (Multi-Tenant)",
      description: "Yes/No/N/A with tenant-based response sharing",
    },
    {
      value: "rating-number",
      label: "Ratings by Number",
      description: "0-10 circular rating scale",
    },
    {
      value: "satisfaction-rating",
      label: "Satisfaction Rating (1-5)",
      description: "1-5 circular rating scale",
    },
    {
      value: "productNPSTGWBuckets",
      label: "Product NPS TGW Buckets",
      description: "6-level hierarchical complaint categorization: L1 (Groups) → L2 (Sub-issues) → L3 (Questions) → L4 (Answers) → L5 (Details) → L6 (Final Options)",
    },
        {
      value: "chassis-with-zone",
      label: "Chassis with Zone",
      description: "Chassis number input with zone-based status follow-up",
    },
    {
      value: "chassis-without-zone",
      label: "Chassis without Zone",
      description: "Chassis number input with direct status follow-up",
    },
    {
      value: "zone-in",
      label: "Zone In",
      description: "Chassis number input with Zone In status (Accepted/Rework/Rejected)",
    },
    {
      value: "zone-out",
      label: "Zone Out",
      description: "Chassis number input with Zone Out status (Accepted/Rework/Rejected)",
    },

  ];

  const fileTypeOptions = [
    { value: "any", label: "Any file type" },
    { value: "image", label: "Images (JPG, PNG, GIF)" },
    { value: "pdf", label: "PDF" },
    { value: "excel", label: "Excel (XLS, XLSX)" },
    { value: "doc", label: "Word (DOC, DOCX)" },
    { value: "stp", label: "STP File (.stp)" },
    { value: "pvz", label: "PVZ File (.pvz)" },

  ];

  if (mode === "list") {
    return (
      <div className="p-4">
        <div className="mb-6">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between">
            <div>
              <h1 className="text-2xl font-bold text-primary-800">
                Service Request Management
              </h1>
              <p className="text-primary-600">
                Create, edit, and manage service request forms
              </p>
            </div>
            {canManage && (
            <button
              onClick={handleCreateForm}
              className="btn-primary mt-4 sm:mt-0"
            >
              <Plus className="w-4 h-4 mr-2" />
              Create New Service Form
            </button>
            )}
          </div>
        </div>

        {forms.length === 0 ? (
          <div className="text-center py-12 bg-white dark:bg-gray-900 rounded-lg border border-neutral-200 dark:border-gray-700">
            <FileText className="w-12 h-12 text-primary-300 mx-auto mb-4" />
            <h3 className="text-lg font-medium text-primary-600 mb-2">
              No service forms created yet
            </h3>
            {/* <p className="text-primary-500 mb-6">
              Create your first service form to get started
            </p> */}
            {canManage && (
            <button onClick={handleCreateForm} className="btn-primary">
              <Plus className="w-4 h-4 mr-2" />
              Create Your First Form
            </button>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {forms.map((form: any) => (
              <div
                key={form.id}
                className="bg-white dark:bg-gray-900 rounded-lg border border-neutral-200 dark:border-gray-700 p-6 hover:shadow-md transition-shadow"
              >
                <div className="flex items-start justify-between mb-4">
                  <div className="flex-1">
                    <h3 className="font-medium text-primary-800 mb-2 line-clamp-2">
                      {form.title}
                    </h3>
                    <p className="text-sm text-primary-600 line-clamp-2">
                      {form.description}
                    </p>
                  </div>
                  <div className="relative ml-2">
                    <button className="p-1 hover:bg-neutral-100 rounded">
                      <MoreVertical className="w-4 h-4 text-primary-400" />
                    </button>
                  </div>
                </div>

                {/* Form Stats */}
                <div className="flex items-center justify-between text-xs text-primary-500 mb-4">
                  <div className="flex items-center">
                    <Users className="w-3 h-3 mr-1" />0 responses
                  </div>
                  <div className="flex items-center">
                    <Calendar className="w-3 h-3 mr-1" />
                    {new Date().toLocaleDateString()}
                  </div>
                </div>

                {/* Visibility Status */}
                <div className="flex items-center justify-between mb-4">
                  <span
                    className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${form.isVisible
                      ? "bg-green-100 text-green-800"
                      : "bg-yellow-100 text-yellow-800"
                      }`}
                  >
                    {form.isVisible ? "Public" : "Private"}
                  </span>
                  <button
                    onClick={() =>
                      handleToggleVisibility(form.id, form.isVisible)
                    }
                    className="text-xs text-primary-600 hover:text-primary-800"
                  >
                    Toggle
                  </button>
                </div>

                {/* Actions */}
                <div className="flex items-center justify-between">
                  <div className="flex space-x-2">
                    <button
                      onClick={() => handlePreviewForm(form.id)}
                      className="px-3 py-2 text-sm font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700"
                    >
                      View
                    </button>
                    <button
                      onClick={() => handleEditForm(form.id)}
                      className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
                      title="Edit form"
                    >
                      <Edit2 className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => handleViewResponses(form.id)}
                      className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
                      title="View responses"
                    >
                      <BarChart3 className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="flex space-x-2">
                    <button
                      onClick={() => handleDuplicate(form)}
                      className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
                      title="Duplicate form"
                    >
                      <Copy className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => handleDelete(form.id, form.title)}
                      className="p-2 text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-colors"
                      title="Delete form"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }



  return (
    <div className="w-full overflow-auto bg-gradient-to-br from-blue-50 via-white to-blue-50 dark:from-gray-950 dark:via-gray-900 dark:to-gray-950">
      {/* Header */}
      <div className="bg-white dark:bg-gray-900 shadow-sm border-b border-gray-200 dark:border-gray-700 sticky top-0 z-50 w-full">
        <div className="w-full px-4 sm:px-6 py-4">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <h1 className="text-2xl font-black text-gray-900 dark:text-white tracking-tight">
                    {id ? "Edit Form" : "Create New Form"}
                  </h1>
                </div>
                <p className="text-sm text-gray-500 font-medium">
                  Design and deploy powerful service forms
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center bg-gray-100 dark:bg-gray-800 p-1 rounded-xl mr-2">
                <button
                  onClick={() => handleExportTemplate('flat')}
                  title="Download Flat Structure Template (FU1-FU99)"
                  className="px-4 py-2 text-xs font-bold text-gray-600 dark:text-gray-400 hover:text-primary-600 dark:hover:text-primary-400 hover:bg-white dark:hover:bg-gray-700 rounded-lg transition-all flex items-center gap-2"
                >
                  <Download className="w-4 h-4" />
                  Export Flat
                </button>
                <div className="w-px h-4 bg-gray-300 dark:bg-gray-600 mx-1" />
                <button
                  onClick={() => handleExportTemplate('nested')}
                  title="Download Nested Hierarchical Template (FU1.1, FU1.1.1)"
                  className="px-4 py-2 text-xs font-bold text-gray-600 dark:text-gray-400 hover:text-primary-600 dark:hover:text-primary-400 hover:bg-white dark:hover:bg-gray-700 rounded-lg transition-all flex items-center gap-2"
                >
                  <Download className="w-4 h-4" />
                  Export Nested
                </button>
                <div className="w-px h-4 bg-gray-300 dark:bg-gray-600 mx-1" />
                {canManage && (
                  <button
                    onClick={(e) => {
                      const input = e.currentTarget
                        .nextElementSibling as HTMLInputElement;
                      input?.click();
                    }}
                    className="px-4 py-2 text-xs font-bold text-gray-600 dark:text-gray-400 hover:text-primary-600 dark:hover:text-primary-400 hover:bg-white dark:hover:bg-gray-700 rounded-lg transition-all flex items-center gap-2"
                  >
                    <Upload className="w-4 h-4" />
                    Import
                  </button>
                )}
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleImportTemplate}
                  className="hidden"
                />
              </div>

              {canManage && (
                <button
                  onClick={() => setShowParameterModal(true)}
                  className="px-4 py-2.5 bg-white dark:bg-gray-800 border-2 border-gray-100 dark:border-gray-700 text-gray-700 dark:text-gray-300 font-bold text-sm rounded-xl hover:border-primary-500 hover:text-primary-600 transition-all flex items-center gap-2 shadow-sm"
                >
                  <Plus className="w-4 h-4" />
                  Parameters
                </button>
              )}

              <div className="w-px h-8 bg-gray-200 dark:bg-gray-700 mx-1 hidden sm:block" />

              <button
                onClick={() => navigate("/forms/analytics")}
                className="px-5 py-2.5 text-gray-500 hover:text-gray-900 dark:hover:text-white font-bold text-sm transition-colors"
              >
                Cancel
              </button>

              {canManage && (
                <button
                  onClick={handleSave}
                  className="px-8 py-3 bg-primary-600 text-white font-black text-sm rounded-xl shadow-[0_10px_20px_rgba(30,58,138,0.2)] hover:shadow-primary-600/30 hover:bg-primary-700 transition-all active:scale-95 flex items-center gap-2"
                >
                  <Save className="w-4 h-4" />
                  {id ? "Update Changes" : "Save & Publish"}
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="w-full py-6 px-4 sm:py-8 sm:px-6">
        <div className="flex flex-col lg:flex-row gap-6">
          {/* Form Editor */}
          <div className="flex-1 space-y-6 min-w-0">
            {/* Form Details */}
            <div className="bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-xl shadow-sm border-l-4 border-l-blue-500 overflow-hidden">
              <div className="bg-gradient-to-r from-blue-50 to-blue-50 dark:from-blue-950/60 dark:to-blue-900/60 px-6 py-4 border-b border-blue-100 dark:border-blue-900/60">
                <h2 className="text-lg font-bold text-blue-900 dark:text-blue-100">
                  📋 Form Details
                </h2>
                <p className="text-sm text-blue-700 dark:text-blue-300 mt-1">
                  Set your form's basic information
                </p>
              </div>

              <div className="p-6 space-y-5">
                <div>
                  <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                    Form Title *
                  </label>
                  <input
                    type="text"
                    value={form.title}
                    onChange={(e) =>
                      setForm((prev) => ({ ...prev, title: e.target.value }))
                    }
                    className="w-full px-4 py-3 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-base font-medium"
                    placeholder="e.g., Customer Feedback Survey"
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                    Description
                  </label>
                  <textarea
                    value={form.description}
                    onChange={(e) =>
                      setForm((prev) => ({
                        ...prev,
                        description: e.target.value,
                      }))
                    }
                    className="w-full px-4 py-3 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all resize-none text-sm"
                    rows={3}
                    placeholder="Describe the purpose of your form..."
                  />
                </div>

                {/* Chassis Number Management */}
                <div className="pt-4 border-t border-gray-200 dark:border-gray-700">
                  <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                    Chassis Numbers
                  </label>
                  <div className="p-4 bg-purple-50 dark:bg-purple-900/20 border border-purple-100 dark:border-purple-800 rounded-xl space-y-4">
                    <div className="space-y-2">
                      <input
                        type="text"
                        value={newChassisNumber}
                        onChange={(e) => setNewChassisNumber(e.target.value)}
                        className="w-full px-4 py-2 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-purple-500 bg-white dark:bg-gray-800 text-gray-900 dark:text-white text-sm"
                        placeholder="Chassis Number (e.g., GE122879)"
                        onKeyPress={(e) => e.key === "Enter" && handleAddChassis()}
                      />
                      <input
                        type="text"
                        value={newPartDescription}
                        onChange={(e) => setNewPartDescription(e.target.value)}
                        className="w-full px-4 py-2 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-purple-500 bg-white dark:bg-gray-800 text-gray-900 dark:text-white text-sm"
                        placeholder="Part Description (e.g., Chassis N603)"
                        onKeyPress={(e) => e.key === "Enter" && handleAddChassis()}
                      />
                      <button
                        type="button"
                        onClick={handleAddChassis}
                        className="w-full px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors text-sm font-bold shadow-sm"
                      >
                        Add
                      </button>
                    </div>

                    <div className="space-y-2">
                      {form.chassisNumbers.map((cn, idx) => (
                        <div
                          key={`${cn.chassisNumber}-${idx}`}
                          className="flex items-center justify-between p-3 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-lg shadow-sm"
                        >
                          <div className="flex items-center gap-2">
                            <div className="flex flex-col">
                              <span className="text-sm font-medium text-gray-700 dark:text-gray-300">
                                {cn.chassisNumber}
                              </span>
                              {cn.partDescription && (
                                <span className="text-xs text-gray-500 dark:text-gray-400">
                                  {cn.partDescription}
                                </span>
                              )}
                            </div>
                            {chassisTenantAssignments[cn.chassisNumber]?.length > 0 && (
                              <span className="px-1.5 py-0.5 bg-purple-100 dark:bg-purple-900/40 text-purple-700 dark:text-purple-300 text-[10px] rounded-full font-bold">
                                {chassisTenantAssignments[cn.chassisNumber].length} Tenants
                              </span>
                            )}
                          </div>
                          <div className="flex items-center gap-1">
                            <div className="relative">
                              <button
                                type="button"
                                onClick={() =>
                                  setOpenChassisMenu(
                                    openChassisMenu === idx ? null : idx
                                  )
                                }
                                className="p-1.5 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors"
                              >
                                <MoreVertical className="w-4 h-4 text-gray-500" />
                              </button>
                              {openChassisMenu === idx && (
                                <div className="absolute right-0 top-full mt-1 w-48 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl shadow-xl z-[40] overflow-hidden border-t-4 border-t-purple-500">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      setChassisTenantConfig({
                                        chassisNumber: cn.chassisNumber,
                                        index: idx,
                                      });
                                      setShowChassisTenantModal(true);
                                      setOpenChassisMenu(null);
                                    }}
                                    className="w-full text-left px-4 py-3 text-sm text-purple-600 dark:text-purple-400 hover:bg-purple-50 dark:hover:bg-gray-700/50 flex items-center gap-2 transition-colors border-b border-gray-100 dark:border-gray-700"
                                  >
                                    <Users className="w-4 h-4" />
                                    Link Tenants
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      handleRemoveChassis(idx);
                                      setOpenChassisMenu(null);
                                    }}
                                    className="w-full text-left px-4 py-3 text-sm text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-gray-700/50 flex items-center gap-2 transition-colors"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                    Remove
                                  </button>
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                    {form.chassisNumbers.length === 0 && (
                      <p className="text-center text-xs text-gray-500 dark:text-gray-400 italic py-2">
                        No chassis numbers added yet.
                      </p>
                    )}
                  </div>
                </div>

                {/* Tenant Selector for SuperAdmin */}
                {user?.role === "superadmin" && (
                  <div className="space-y-4 pt-4 border-t border-gray-200 dark:border-gray-700">
                    <div className="flex items-center gap-2">
                      <input
                        type="checkbox"
                        id="isGlobal"
                        checked={isGlobal}
                        onChange={(e) => setIsGlobal(e.target.checked)}
                        className="rounded border-neutral-300 text-primary-600 focus:ring-primary-500"
                      />
                      <label
                        htmlFor="isGlobal"
                        className="text-sm font-bold text-primary-900 dark:text-primary-100"
                      >
                        Make this a Global Form
                      </label>
                    </div>

                    {!isGlobal ? (
                      <div>
                        <label className="block text-sm font-medium text-primary-700 mb-2">
                          Primary Tenant *
                        </label>
                        <select
                          value={selectedTenantId}
                          onChange={(e) => setSelectedTenantId(e.target.value)}
                          className="input-field"
                          required={!isGlobal}
                        >
                          <option value="">Select a tenant...</option>
                          {tenants.map((tenant) => (
                            <option key={tenant._id} value={tenant._id}>
                              {tenant.companyName} ({tenant.name})
                            </option>
                          ))}
                        </select>
                      </div>
                    ) : (
                      <div>
                        <label className="block text-sm font-medium text-primary-700 mb-2">
                          Assign to Tenants
                        </label>
                        <div className="grid grid-cols-1 gap-2 max-h-48 overflow-y-auto p-3 border-2 border-gray-200 dark:border-gray-700 rounded-lg">
                          {tenants.map((tenant) => (
                            <label
                              key={tenant._id}
                              className="flex items-center gap-3 p-2 hover:bg-gray-50 dark:hover:bg-gray-800 rounded-md cursor-pointer transition-colors"
                            >
                              <input
                                type="checkbox"
                                checked={sharedWithTenants.includes(tenant._id)}
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    setSharedWithTenants((prev) => [
                                      ...prev,
                                      tenant._id,
                                    ]);
                                  } else {
                                    setSharedWithTenants((prev) =>
                                      prev.filter((id) => id !== tenant._id)
                                    );
                                  }
                                }}
                                className="rounded border-neutral-300 text-primary-600 focus:ring-primary-500"
                              />
                              <span className="text-sm text-gray-700 dark:text-gray-300">
                                {tenant.companyName}
                              </span>
                            </label>
                          ))}
                        </div>
                        <p className="mt-1 text-xs text-primary-500">
                          {sharedWithTenants.length} tenants selected
                        </p>
                      </div>
                    )}
                  </div>
                )}

                <div className="flex items-center">
                  <input
                    type="checkbox"
                    id="isVisible"
                    checked={form.isVisible}
                    onChange={(e) =>
                      setForm((prev) => ({
                        ...prev,
                        isVisible: e.target.checked,
                      }))
                    }
                    className="rounded border-neutral-300 text-primary-600 focus:ring-primary-500"
                  />
                  <label
                    htmlFor="isVisible"
                    className="ml-2 text-sm text-primary-700 dark:text-primary-200"
                  >
                    Make form publicly visible
                  </label>
                </div>

                <div className="flex items-center">
                  <input
                    type="checkbox"
                    id="locationEnabled"
                    checked={form.locationEnabled !== false}
                    onChange={(e) =>
                      setForm((prev) => ({
                        ...prev,
                        locationEnabled: e.target.checked,
                      }))
                    }
                    className="rounded border-neutral-300 text-primary-600 focus:ring-primary-500"
                  />
                  <label
                    htmlFor="locationEnabled"
                    className="ml-2 text-sm text-primary-700 dark:text-primary-200"
                  >
                    Enable location tracking for responses
                  </label>
                </div>

                {/* Load Demo Data Button */}
                <div className="pt-4 border-t border-neutral-200 dark:border-gray-700">
                  <button
                    type="button"
                    onClick={loadDemoData}
                    className="w-full px-4 py-2 bg-gradient-to-r from-blue-500 to-blue-600 text-white rounded-lg hover:from-blue-600 hover:to-blue-700 transition-all duration-200 flex items-center justify-center gap-2 shadow-md hover:shadow-lg"
                  >
                    <FileText className="w-4 h-4" />
                    Load Demo Data (For Testing)
                  </button>
                  <p className="mt-2 text-xs text-center text-primary-500">
                    Loads a complete form with follow-up questions for testing
                  </p>
                </div>
              </div>
            </div>

            {/* Page Navigation */}
            {(() => {
              if (totalPages > 1) {
                const visiblePages = pages.slice(
                  pageWindowStart,
                  pageWindowStart + MAX_VISIBLE_PAGE_BUTTONS
                );
                const canScrollLeft = pageWindowStart > 0;
                const canScrollRight =
                  pageWindowStart + MAX_VISIBLE_PAGE_BUTTONS < totalPages;

                return (
                  <div className="bg-white dark:bg-gray-900 rounded-xl shadow-md border border-blue-200 dark:border-blue-900/60 p-4 mb-6">
                    <div className="flex items-center justify-between">
                      <h4 className="text-sm font-semibold text-gray-700 dark:text-gray-300">
                        Form Pages:
                      </h4>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() =>
                            setPageWindowStart((prev) => Math.max(0, prev - 1))
                          }
                          disabled={!canScrollLeft}
                          className="w-10 h-10 rounded-lg bg-gray-100 text-gray-600 dark:bg-gray-800 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-gray-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center transition-all duration-200"
                        >
                          <ChevronLeft className="w-4 h-4" />
                        </button>
                        <div className="flex items-center gap-2">
                          {visiblePages.map((_, index) => {
                            const pageIndex = pageWindowStart + index;
                            const isActive = currentPage === pageIndex;

                            return (
                              <button
                                key={pageIndex}
                                onClick={() => handlePageChange(pageIndex)}
                                className={`w-10 h-10 rounded-lg font-bold text-sm transition-all duration-200 ${isActive
                                  ? "bg-gradient-to-r from-blue-600 to-blue-700 text-white shadow-lg scale-110"
                                  : "bg-gray-100 text-gray-600 hover:bg-gray-200 dark:bg-gray-800 dark:text-gray-300 dark:hover:bg-gray-700"
                                  }`}
                              >
                                {pageIndex + 1}
                              </button>
                            );
                          })}
                        </div>
                        <button
                          onClick={() =>
                            setPageWindowStart((prev) => {
                              const maxStart = Math.max(
                                0,
                                totalPages - MAX_VISIBLE_PAGE_BUTTONS
                              );
                              return Math.min(prev + 1, maxStart);
                            })
                          }
                          disabled={!canScrollRight}
                          className="w-10 h-10 rounded-lg bg-gray-100 text-gray-600 dark:bg-gray-800 dark:text-gray-300 hover:bg-gray-200 dark:hover:bg-gray-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center transition-all duration-200"
                        >
                          <ChevronRight className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-2">
                      Viewing Page {currentPage + 1} of {totalPages} · Each main
                      section is a separate page
                    </p>
                  </div>
                );
              }
              return null;
            })()}

            {/* Sections - Display current page */}
            {(() => {
              const pages = getPagesFromSections();
              const currentPageSections = pages[currentPage] || [];

              return currentPageSections.map((section, indexInPage) => {
                // Calculate the global section index
                const globalSectionIndex = form.sections.findIndex(
                  (s) => s.id === section.id
                );
                const sectionLabel = section.isSubsection
                  ? String.fromCharCode(
                    65 +
                    form.sections.findIndex(
                      (s) => s.id === section.parentSectionId
                    )
                  ) +
                  "." +
                  indexInPage
                  : String.fromCharCode(65 + currentPage); // A, B, C for main sections

                return (
                  <div
                    key={section.id}
                    className={`bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-xl shadow-sm hover:shadow-md transition-shadow duration-200 overflow-hidden ${section.isSubsection
                      ? "border-l-4 border-l-green-500 ml-6"
                      : "border-l-4 border-l-blue-500"
                      }`}
                  >
                    {/* Section Header */}
                    <div
                      className={`px-6 py-4 border-b ${section.isSubsection
                        ? "bg-gradient-to-r from-green-50 to-teal-50 border-green-100 dark:from-green-950/60 dark:to-teal-900/60 dark:border-green-900/60"
                        : "bg-gradient-to-r from-blue-50 to-blue-50 border-blue-100 dark:from-blue-950/60 dark:to-blue-900/60 dark:border-blue-900/60"
                        }`}
                    >
                      <div className="flex items-start justify-between">
                        <div className="flex-1">
                          <div className="flex items-center gap-3 mb-2">
                            <div
                              className={`flex items-center justify-center w-8 h-8 rounded-full text-white font-bold text-sm shadow-md ${section.isSubsection
                                ? "bg-green-600"
                                : "bg-blue-600"
                                }`}
                            >
                              {sectionLabel}
                            </div>
                            <h3
                              className={`text-lg font-bold ${section.isSubsection
                                ? "text-green-900 dark:text-green-100"
                                : "text-blue-900 dark:text-blue-100"
                                }`}
                            >
                              {section.isSubsection ? "Subsection" : "Section"}{" "}
                              {sectionLabel}
                              {section.title && (
                                <span
                                  className={`font-normal ml-2 ${section.isSubsection
                                    ? "text-green-600 dark:text-green-300"
                                    : "text-blue-600 dark:text-blue-300"
                                    }`}
                                >
                                  · {section.title}
                                </span>
                              )}
                            </h3>
                          </div>
                          {section.description && (
                            <p
                              className={`text-sm ml-11 ${section.isSubsection
                                ? "text-green-700 dark:text-green-300"
                                : "text-blue-700 dark:text-blue-300"
                                }`}
                            >
                              {section.description}
                            </p>
                          )}
                          {section.isSubsection && section.parentSectionId && (
                            <p className="text-xs ml-11 text-green-600 dark:text-green-400 mt-2">
                              Merged with:{" "}
                              <strong>
                                {form.sections.find(
                                  (s) => s.id === section.parentSectionId
                                )?.title || "Parent Section"}
                              </strong>
                            </p>
                          )}
                        </div>
                        <div className="flex items-center gap-2">
                          {!section.isSubsection && (
                            <button
                              onClick={() => addSubsection(section.id)}
                              className="p-2 text-green-600 hover:text-green-700 hover:bg-green-50 rounded-lg transition-colors"
                              title="Add subsection"
                            >
                              <Plus className="w-5 h-5" />
                            </button>
                          )}
                          <button
                            onClick={() => duplicateSection(section.id)}
                            className="p-2 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-colors"
                            title="Duplicate section"
                          >
                            <Copy className="w-5 h-5" />
                          </button>
                          {form.sections.filter((s) => !s.isSubsection).length >
                            1 && (
                              <button
                                onClick={() => deleteSection(section.id)}
                                className="p-2 text-red-500 hover:text-red-700 hover:bg-red-50 rounded-lg transition-colors"
                                title="Delete section"
                              >
                                <Trash2 className="w-5 h-5" />
                              </button>
                            )}
                        </div>
                      </div>
                    </div>

                    {/* Section Details (Optional) */}
                    <div className="px-6 py-4 bg-gray-50 dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700">
                      <div className="space-y-3">
                        <div>
                          <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-1.5 uppercase tracking-wide">
                            Section Title (Optional)
                          </label>
                          <input
                            type="text"
                            value={section.title}
                            onChange={(e) =>
                              updateSection(section.id, {
                                title: e.target.value,
                              })
                            }
                            className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all text-sm"
                            placeholder="e.g., Personal Information, Contact Details"
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-1.5 uppercase tracking-wide">
                            Section Description (Optional)
                          </label>
                          <textarea
                            value={section.description}
                            onChange={(e) =>
                              updateSection(section.id, {
                                description: e.target.value,
                              })
                            }
                            className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all resize-none text-sm"
                            rows={2}
                            placeholder="Brief description for respondents"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Questions Container */}
                    <div className="p-6 space-y-5">
                      {/* Insert button at the beginning */}
                      {section.questions.length > 0 && (
                        <div className="flex justify-center -mb-3 relative z-10">
                          <button
                            onClick={() => insertQuestionAt(section.id, 0)}
                            className="group bg-white dark:bg-gray-900 border-2 border-dashed border-blue-300 rounded-full p-2 text-blue-400 hover:bg-blue-50 hover:border-blue-500 hover:text-blue-600 transition-all duration-200"
                            title="Insert question at the beginning"
                          >
                            <Plus className="w-5 h-5" />
                          </button>
                        </div>
                      )}

                      {section.questions.map((question, questionIndex) => {
                        const selectedFileType =
                          question.allowedFileTypes?.[0] ?? "any";
                        const selectedFileTypeOption = fileTypeOptions.find(
                          (option) => option.value === selectedFileType
                        );

                        return (
                          <React.Fragment key={question.id}>
                            {/* Insert button between questions */}
                            {questionIndex > 0 && (
                              <div className="flex justify-center -my-3 relative z-10">
                                <button
                                  onClick={() =>
                                    insertQuestionAt(section.id, questionIndex)
                                  }
                                  className="group bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 border-2 border-blue-400 rounded-full p-2 text-white transition-all duration-200 shadow-lg hover:shadow-xl transform hover:scale-110"
                                  title="Insert question here"
                                >
                                  <Plus className="w-5 h-5" />
                                </button>
                              </div>
                            )}

                            {/* Question Card */}
                            <div className="bg-white dark:bg-gray-900 border-2 border-gray-200 dark:border-gray-700 rounded-xl shadow-sm hover:shadow-lg hover:border-blue-300 transition-all duration-200">
                              {/* Question Header */}
                              <div className="flex items-center justify-between px-5 py-3 bg-gradient-to-r from-gray-50 to-gray-100 border-b border-gray-200 dark:border-gray-700 rounded-t-xl">
                                <div className="flex items-center gap-3">
                                  <div className="flex items-center justify-center w-7 h-7 rounded-full bg-blue-600 text-white font-bold text-xs shadow">
                                    {questionIndex + 1}
                                  </div>
                                  <span className="text-sm font-semibold text-gray-700 dark:text-gray-300">
                                    Question {questionIndex + 1}
                                  </span>
                                </div>

                                {/* Action Buttons */}
                               <div className="flex items-center gap-1">
  {/* Desktop: show all buttons inline */}
  <div className="hidden md:flex items-center gap-1">
    <label className="flex items-center space-x-1.5 cursor-pointer px-3 py-1 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg hover:border-blue-400 transition-all shadow-sm" title="Track Question">
      <input
        type="checkbox"
        checked={question.trackResponseQuestion || false}
        disabled={question.trackResponseRank || false}
        onChange={(e) => updateQuestion(section.id, question.id, { trackResponseQuestion: e.target.checked })}
        className="w-4 h-4 text-blue-600 focus:ring-2 focus:ring-blue-500 border-gray-300 dark:border-gray-600 rounded transition-all disabled:opacity-50"
      />
      <span className="text-xs font-bold text-blue-700 dark:text-blue-300 whitespace-nowrap">Track Question</span>
    </label>
    <label className="flex items-center space-x-1.5 cursor-pointer px-3 py-1 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg hover:border-blue-400 transition-all shadow-sm ml-1" title="Track Rank">
      <input
        type="checkbox"
        checked={question.trackResponseRank || false}
        disabled={question.trackResponseQuestion || false}
        onChange={(e) => updateQuestion(section.id, question.id, { trackResponseRank: e.target.checked })}
        className="w-4 h-4 text-blue-600 focus:ring-2 focus:ring-blue-500 border-gray-300 dark:border-gray-600 rounded transition-all disabled:opacity-50"
      />
      <span className="text-xs font-bold text-blue-700 dark:text-blue-300 whitespace-nowrap">Track Rank</span>
    </label>
    <button onClick={() => moveQuestionUp(section.id, question.id)} disabled={questionIndex === 0} className="p-2 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed" title="Move up">
      <ChevronUp className="w-5 h-5" />
    </button>
    <button onClick={() => moveQuestionDown(section.id, question.id)} disabled={questionIndex === section.questions.length - 1} className="p-2 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed" title="Move down">
      <ChevronDown className="w-5 h-5" />
    </button>
    <button onClick={() => duplicateQuestion(section.id, question.id)} className="p-2 text-green-600 hover:text-green-800 hover:bg-green-50 rounded-lg transition-colors" title="Duplicate question">
      <Copy className="w-5 h-5" />
    </button>
    <button onClick={() => splitSectionAtQuestion(section.id, question.id)} disabled={questionIndex === 0} className="p-2 text-purple-600 hover:text-purple-800 hover:bg-purple-50 rounded-lg transition-colors disabled:opacity-30 disabled:cursor-not-allowed" title="Split into new section">
      <Scissors className="w-5 h-5" />
    </button>
    <button onClick={() => deleteQuestion(section.id, question.id)} className="p-2 text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-colors" title="Delete question">
      <Trash2 className="w-5 h-5" />
    </button>
  </div>

  {/* Mobile: 3-dot dropdown menu */}
  <div className="relative md:hidden">
    <button
      onClick={() => setOpenQuestionMenu(openQuestionMenu === question.id ? null : question.id)}
      className="p-2 text-gray-600 hover:text-gray-900 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors"
      title="More options"
    >
      <MoreVertical className="w-5 h-5" />
    </button>

    {openQuestionMenu === question.id && (
      <>
        <div className="fixed inset-0 z-10" onClick={() => setOpenQuestionMenu(null)} />
        <div className="absolute right-0 top-full mt-1 w-56 bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl shadow-xl z-20 overflow-hidden border-t-4 border-t-blue-500">
          
          {/* Track Question */}
          <label className="flex items-center gap-3 px-4 py-3 hover:bg-blue-50 dark:hover:bg-gray-700 cursor-pointer border-b border-gray-100 dark:border-gray-700">
            <input
              type="checkbox"
              checked={question.trackResponseQuestion || false}
              disabled={question.trackResponseRank || false}
              onChange={(e) => {
                updateQuestion(section.id, question.id, { trackResponseQuestion: e.target.checked });
                setOpenQuestionMenu(null);
              }}
              className="w-4 h-4 text-blue-600 rounded disabled:opacity-50"
            />
            <span className="text-sm text-blue-700 dark:text-blue-300 font-medium">Track Question</span>
          </label>

          {/* Track Rank */}
          <label className="flex items-center gap-3 px-4 py-3 hover:bg-blue-50 dark:hover:bg-gray-700 cursor-pointer border-b border-gray-100 dark:border-gray-700">
            <input
              type="checkbox"
              checked={question.trackResponseRank || false}
              disabled={question.trackResponseQuestion || false}
              onChange={(e) => {
                updateQuestion(section.id, question.id, { trackResponseRank: e.target.checked });
                setOpenQuestionMenu(null);
              }}
              className="w-4 h-4 text-blue-600 rounded disabled:opacity-50"
            />
            <span className="text-sm text-blue-700 dark:text-blue-300 font-medium">Track Rank</span>
          </label>

          {/* Move Up */}
          <button
            onClick={() => { moveQuestionUp(section.id, question.id); setOpenQuestionMenu(null); }}
            disabled={questionIndex === 0}
            className="w-full text-left px-4 py-3 text-sm text-gray-700 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-700 flex items-center gap-3 disabled:opacity-40 disabled:cursor-not-allowed border-b border-gray-100 dark:border-gray-700"
          >
            <ChevronUp className="w-4 h-4 text-blue-500" />
            Move Up
          </button>

          {/* Move Down */}
          <button
            onClick={() => { moveQuestionDown(section.id, question.id); setOpenQuestionMenu(null); }}
            disabled={questionIndex === section.questions.length - 1}
            className="w-full text-left px-4 py-3 text-sm text-gray-700 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-700 flex items-center gap-3 disabled:opacity-40 disabled:cursor-not-allowed border-b border-gray-100 dark:border-gray-700"
          >
            <ChevronDown className="w-4 h-4 text-blue-500" />
            Move Down
          </button>

          {/* Duplicate */}
          <button
            onClick={() => { duplicateQuestion(section.id, question.id); setOpenQuestionMenu(null); }}
            className="w-full text-left px-4 py-3 text-sm text-green-700 dark:text-green-300 hover:bg-green-50 dark:hover:bg-gray-700 flex items-center gap-3 border-b border-gray-100 dark:border-gray-700"
          >
            <Copy className="w-4 h-4" />
            Duplicate
          </button>

          {/* Split Section */}
          <button
            onClick={() => { splitSectionAtQuestion(section.id, question.id); setOpenQuestionMenu(null); }}
            disabled={questionIndex === 0}
            className="w-full text-left px-4 py-3 text-sm text-purple-700 dark:text-purple-300 hover:bg-purple-50 dark:hover:bg-gray-700 flex items-center gap-3 disabled:opacity-40 disabled:cursor-not-allowed border-b border-gray-100 dark:border-gray-700"
          >
            <Scissors className="w-4 h-4" />
            Split Section Here
          </button>

          {/* Delete */}
          <button
            onClick={() => { deleteQuestion(section.id, question.id); setOpenQuestionMenu(null); }}
            className="w-full text-left px-4 py-3 text-sm text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-gray-700 flex items-center gap-3"
          >
            <Trash2 className="w-4 h-4" />
            Delete Question
          </button>
        </div>
      </>
    )}
  </div>
</div>
                              </div>

                              {/* Question Body */}
                              <div className="p-5">
                                {/* Parameter Display */}
                                {question.subParam1 && (
                                  <div className="mb-4">
                                    <div className="inline-flex items-center gap-2 px-3 py-1.5 bg-gradient-to-r from-purple-100 to-purple-200 dark:from-purple-900/30 dark:to-purple-800/30 border border-purple-300 dark:border-purple-600 rounded-lg">
                                      <div className="w-2 h-2 bg-purple-500 rounded-full"></div>
                                      <span className="text-sm font-semibold text-purple-800 dark:text-purple-200">
                                        {question.subParam1}
                                      </span>
                                    </div>
                                  </div>
                                )}

                               <div className="space-y-4">
  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
    <div>
      <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
        Question Text (optional when using image)
      </label>
      <input
        type="text"
        value={question.text}
        onChange={(e) => updateQuestion(section.id, question.id, { text: e.target.value })}
        className="w-full px-3 py-2.5 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm"
        placeholder="Enter your question"
      />
    </div>
    <div>
      <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
        Question Type
      </label>
      <select
        value={question.type}
        onChange={(e) => handleQuestionTypeChange(section.id, question.id, question, e.target.value)}
        className="w-full px-3 py-2.5 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm"
      >
        {questionTypes.map((type) => (
          <option key={type.value} value={type.value}>{type.label}</option>
        ))}
      </select>
    </div>
  </div>
  {/* Image upload below both fields */}
  <div className="mt-1 space-y-2">
    {question.imageUrl ? (
      <div className="flex items-center gap-4 bg-gray-50 dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-lg p-3">
        <img
          src={question.imageUrl}
          alt={`Question ${questionIndex + 1} image`}
          className="h-20 w-20 object-contain rounded-md border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900"
        />
        <button
          type="button"
          onClick={() => clearQuestionImage(section.id, question.id)}
          className="px-3 py-2 text-sm font-semibold text-red-600 hover:text-white hover:bg-red-500 border border-red-200 rounded-lg transition-colors"
        >
          Remove Image
        </button>
      </div>
    ) : null}
    <label className="inline-flex items-center justify-center w-full px-3 py-2 border-2 border-dashed border-gray-300 dark:border-gray-600 rounded-lg cursor-pointer text-sm font-medium text-blue-600 hover:border-blue-400 hover:text-blue-700 transition-colors">
      <input
        type="file"
        accept="image/*,.stp,.step,.pvz"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) {
            void handleQuestionImageUpload(section.id, question.id, file);
            e.target.value = "";
          }
        }}
      />
      Upload Image/File
    </label>
    <p className="text-xs text-gray-500 dark:text-gray-500">Images, STP, or PVZ files.</p>
  </div>
</div>
                                 {question.trackResponseRank && (
                                  <div className="mt-4 p-4 bg-blue-50/50 dark:bg-blue-900/10 border border-blue-200 dark:border-blue-800 rounded-xl space-y-4">
                                    <div className="flex items-center gap-2 mb-1">
                                      <div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse"></div>
                                      <h4 className="text-sm font-bold text-blue-900 dark:text-blue-100 uppercase tracking-wider">Track Rank Configuration</h4>
                                    </div>
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                      <div>
                                        <label className="block text-xs font-semibold text-blue-700 dark:text-blue-300 mb-2 uppercase tracking-wide">
                                          Track Rank Question Label
                                        </label>
                                        <input
                                          type="text"
                                          value={question.trackResponseRankLabel || ""}
                                          onChange={(e) =>
                                            updateQuestion(section.id, question.id, {
                                              trackResponseRankLabel: e.target.value,
                                            })
                                          }
                                          className="w-full px-3 py-2.5 border-2 border-blue-200 dark:border-blue-800 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm bg-white dark:bg-gray-900"
                                          placeholder="Enter label for rank tracking"
                                        />
                                      </div>
                                      <div>
                                        <label className="block text-xs font-semibold text-blue-700 dark:text-blue-300 mb-2 uppercase tracking-wide">
                                          Track Rank Question Type
                                        </label>
                                        <select
                                          value={question.trackResponseRankType || "text"}
                                          onChange={(e) =>
                                            updateQuestion(section.id, question.id, {
                                              trackResponseRankType: e.target.value,
                                            })
                                          }
                                          className="w-full px-3 py-2.5 border-2 border-blue-200 dark:border-blue-800 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm bg-white dark:bg-gray-900"
                                        >
                                          {questionTypes.map((type) => (
                                            <option key={type.value} value={type.value}>
                                              {type.label}
                                            </option>
                                          ))}
                                        </select>
                                      </div>
                                    </div>
                                  </div>
                                )}

                                {question.trackResponseQuestion && (
                                  <div className="mt-4 p-4 bg-indigo-50/50 dark:bg-indigo-900/10 border border-indigo-200 dark:border-indigo-800 rounded-xl space-y-4">
                                    <div className="flex items-center gap-2 mb-1">
                                      <div className="w-2 h-2 bg-indigo-500 rounded-full animate-pulse"></div>
                                      <h4 className="text-sm font-bold text-indigo-900 dark:text-indigo-100 uppercase tracking-wider">Track Question Configuration</h4>
                                    </div>
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                      <div>
                                        <label className="block text-xs font-semibold text-indigo-700 dark:text-indigo-300 mb-2 uppercase tracking-wide">
                                          Track Question Label
                                        </label>
                                        <input
                                          type="text"
                                          value={question.trackResponseQuestionLabel || ""}
                                          onChange={(e) =>
                                            updateQuestion(section.id, question.id, {
                                              trackResponseQuestionLabel: e.target.value,
                                            })
                                          }
                                          className="w-full px-3 py-2.5 border-2 border-indigo-200 dark:border-indigo-800 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all text-sm bg-white dark:bg-gray-900"
                                          placeholder="Enter label for tracking question"
                                        />
                                      </div>
                                      <div>
                                        <label className="block text-xs font-semibold text-indigo-700 dark:text-indigo-300 mb-2 uppercase tracking-wide">
                                          Track Question Type
                                        </label>
                                        <select
                                          value={question.trackResponseQuestionType || "text"}
                                          onChange={(e) =>
                                            updateQuestion(section.id, question.id, {
                                              trackResponseQuestionType: e.target.value,
                                            })
                                          }
                                          className="w-full px-3 py-2.5 border-2 border-indigo-200 dark:border-indigo-800 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all text-sm bg-white dark:bg-gray-900"
                                        >
                                          {questionTypes.map((type) => (
                                            <option key={type.value} value={type.value}>
                                              {type.label}
                                            </option>
                                          ))}
                                        </select>
                                      </div>
                                    </div>
                                  </div>
                                )}

                                {(question.type === "productNPSTGWBuckets") && (
                                  <div className="mt-4 p-4 bg-blue-50 dark:bg-blue-900 border border-blue-200 dark:border-blue-700 rounded-lg">
                                    <div className="flex items-center justify-between mb-3">
                                      <h4 className="font-semibold text-blue-900 dark:text-blue-100">Hierarchy Levels (Cascading)</h4>
                                      <span className="text-xs font-medium text-blue-700 dark:text-blue-300">Up to 6 levels</span>
                                    </div>
                                    <div className="space-y-3">
                                      {(() => {
                                        const level1Options = getLevel1Options();
                                        let selectedValues = question.selectedHierarchyValues || {};

                                        if (!selectedValues.level1 && level1Options.length > 0) {
                                          const defaultLevel1 = level1Options[0];
                                          const level2Options = getLevel2Options(defaultLevel1);
                                          selectedValues = {
                                            level1: defaultLevel1,
                                            level2: level2Options.length > 0 ? level2Options[0] : undefined,
                                          };
                                        }

                                        const defaultLabels = [
                                          "Complaint Groups",
                                          "Sub-complaints",
                                          "Probing Questions",
                                          "Initial Answers",
                                          "Secondary Details",
                                          "Final Options"
                                        ];

                                        const handleLevelChange = (levelNum: number, value: string) => {
                                          const newValues = { ...selectedValues };
                                          newValues[`level${levelNum}` as keyof typeof selectedValues] = value;

                                          for (let i = levelNum + 1; i <= 6; i++) {
                                            newValues[`level${i}` as keyof typeof selectedValues] = undefined;
                                          }

                                          updateQuestion(section.id, question.id, { selectedHierarchyValues: newValues });
                                        };

                                        return (
                                          <>
                                            {/* Level 1 */}
                                            <div>
                                              <label className="block text-xs font-bold text-blue-900 dark:text-blue-200 mb-1">L1: {defaultLabels[0]}</label>
                                              <select
                                                value={selectedValues.level1 || ""}
                                                onChange={(e) => handleLevelChange(1, e.target.value)}
                                                className="w-full px-2 py-1 text-xs border border-blue-300 dark:border-blue-600 rounded bg-white dark:bg-blue-800 text-blue-900 dark:text-blue-100 focus:ring-2 focus:ring-blue-500"
                                              >
                                                <option value="">Select Level 1</option>
                                                {getLevel1Options().map((opt: string) => (
                                                  <option key={opt} value={opt}>{opt}</option>
                                                ))}
                                              </select>
                                            </div>

                                            {/* Level 2 */}
                                            <div>
                                              <label className="block text-xs font-bold text-blue-900 dark:text-blue-200 mb-1">L2: {defaultLabels[1]}</label>
                                              <select
                                                value={selectedValues.level2 || ""}
                                                onChange={(e) => handleLevelChange(2, e.target.value)}
                                                className="w-full px-2 py-1 text-xs border border-blue-300 dark:border-blue-600 rounded bg-white dark:bg-blue-800 text-blue-900 dark:text-blue-100 focus:ring-2 focus:ring-blue-500"
                                              >
                                                <option value="">Select Level 2</option>
                                                {getLevel2Options(selectedValues.level1 || "").map((opt: string) => (
                                                  <option key={opt} value={opt}>{opt}</option>
                                                ))}
                                              </select>
                                            </div>

                                            {/* Level 3 */}
                                            <div>
                                              <label className="block text-xs font-bold text-blue-900 dark:text-blue-200 mb-1">L3: {defaultLabels[2]}</label>
                                              <select
                                                value={selectedValues.level3 || ""}
                                                onChange={(e) => handleLevelChange(3, e.target.value)}
                                                className="w-full px-2 py-1 text-xs border border-blue-300 dark:border-blue-600 rounded bg-white dark:bg-blue-800 text-blue-900 dark:text-blue-100 focus:ring-2 focus:ring-blue-500"
                                              >
                                                <option value="">Select Level 3</option>
                                                {getLevel3Options(selectedValues.level1 || "", selectedValues.level2 || "").map((opt: string) => (
                                                  <option key={opt} value={opt}>{opt}</option>
                                                ))}
                                              </select>
                                            </div>

                                            {/* Level 4 */}
                                            <div>
                                              <label className="block text-xs font-bold text-blue-900 dark:text-blue-200 mb-1">L4: {defaultLabels[3]}</label>
                                              <select
                                                value={selectedValues.level4 || ""}
                                                onChange={(e) => handleLevelChange(4, e.target.value)}
                                                className="w-full px-2 py-1 text-xs border border-blue-300 dark:border-blue-600 rounded bg-white dark:bg-blue-800 text-blue-900 dark:text-blue-100 focus:ring-2 focus:ring-blue-500"
                                              >
                                                <option value="">Select Level 4</option>
                                                {getLevel4Options(selectedValues.level1 || "", selectedValues.level2 || "", selectedValues.level3 || "").map((opt: string) => (
                                                  <option key={opt} value={opt}>{opt}</option>
                                                ))}
                                              </select>
                                            </div>

                                            {/* Level 5 */}
                                            <div>
                                              <label className="block text-xs font-bold text-blue-900 dark:text-blue-200 mb-1">L5: {defaultLabels[4]}</label>
                                              <select
                                                value={selectedValues.level5 || ""}
                                                onChange={(e) => handleLevelChange(5, e.target.value)}
                                                className="w-full px-2 py-1 text-xs border border-blue-300 dark:border-blue-600 rounded bg-white dark:bg-blue-800 text-blue-900 dark:text-blue-100 focus:ring-2 focus:ring-blue-500"
                                              >
                                                <option value="">Select Level 5</option>
                                                {getLevel5Options(selectedValues.level1 || "", selectedValues.level2 || "", selectedValues.level3 || "", selectedValues.level4 || "").map((opt: string) => (
                                                  <option key={opt} value={opt}>{opt}</option>
                                                ))}
                                              </select>
                                            </div>

                                            {/* Level 6 */}
                                            <div>
                                              <label className="block text-xs font-bold text-blue-900 dark:text-blue-200 mb-1">L6: {defaultLabels[5]}</label>
                                              <select
                                                value={selectedValues.level6 || ""}
                                                onChange={(e) => handleLevelChange(6, e.target.value)}
                                                className="w-full px-2 py-1 text-xs border border-blue-300 dark:border-blue-600 rounded bg-white dark:bg-blue-800 text-blue-900 dark:text-blue-100 focus:ring-2 focus:ring-blue-500"
                                              >
                                                <option value="">Select Level 6</option>
                                                {getLevel6Options(selectedValues.level1 || "", selectedValues.level2 || "", selectedValues.level3 || "", selectedValues.level4 || "", selectedValues.level5 || "").map((opt: string) => (
                                                  <option key={opt} value={opt}>{opt}</option>
                                                ))}
                                              </select>
                                            </div>
                                          </>
                                        );
                                      })()}
                                    </div>
                                    <p className="text-xs text-blue-700 dark:text-blue-300 mt-3 italic">Select Level 1 first, then each subsequent level will show only available options based on your selection.</p>
                                  </div>
                                )}

                                {question.type === "file" ? (
                                  <div className="mt-4">
                                    <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                                      Allowed file type
                                    </label>
                                    <select
                                      value={selectedFileType}
                                      onChange={(e) =>
                                        updateQuestion(
                                          section.id,
                                          question.id,
                                          {
                                            allowedFileTypes:
                                              e.target.value === "any"
                                                ? undefined
                                                : [e.target.value],
                                          }
                                        )
                                      }
                                      className="w-full px-3 py-2.5 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm"
                                    >
                                      {fileTypeOptions.map((option) => (
                                        <option
                                          key={`${question.id}-file-type-${option.value}`}
                                          value={option.value}
                                        >
                                          {option.label}
                                        </option>
                                      ))}
                                    </select>
                                    <p className="text-xs text-gray-500 dark:text-gray-500 mt-1">
                                      {selectedFileType === "any"
                                        ? "Respondents can upload any file type."
                                        : `Respondents must upload files matching ${selectedFileTypeOption?.label}.`}
                                    </p>
                                  </div>
                                ) : null}

                                <div className="mt-4">
                                  <label className="flex items-center cursor-pointer group">
                                    <input
                                      type="checkbox"
                                      checked={question.required}
                                      onChange={(e) =>
                                        updateQuestion(
                                          section.id,
                                          question.id,
                                          {
                                            required: e.target.checked,
                                          }
                                        )
                                      }
                                      className="w-4 h-4 rounded border-gray-300 dark:border-gray-600 text-blue-600 focus:ring-2 focus:ring-blue-500 focus:ring-offset-1 transition-all"
                                    />
                                    <span className="ml-1.5 text-sm text-gray-700 dark:text-gray-300 group-hover:text-blue-700 font-small transition-colors">
                                      Required question
                                    </span>
                                  </label>
                                </div>
                                <div className="mb-4">
                                  <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                                    Suggestion
                                  </label>
                                  <textarea
                                    value={question.suggestion || ""}
                                    onChange={(e) =>
                                      updateQuestion(section.id, question.id, {
                                        suggestion: e.target.value,
                                      })
                                    }
                                    placeholder="Suggestions or recommendations for this question"
                                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-md shadow-sm focus:ring-primary-500 focus:border-primary-500 dark:bg-gray-800 dark:text-gray-100"
                                    rows={3}
                                  />
                                </div>

                               <div className="mt-4 grid grid-cols-1 sm:grid-cols-2 gap-3">
                                  <div>
                                    <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                                      Sub Parameter 1
                                    </label>
                                    <select
                                      value={question.subParam1 || ""}
                                      onChange={(e) =>
                                        updateQuestion(
                                          section.id,
                                          question.id,
                                          {
                                            subParam1: e.target.value,
                                          }
                                        )
                                      }
                                      className="w-full px-3 py-2.5 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm dark:bg-gray-800 dark:text-gray-100"
                                    >
                                      <option value="">
                                        -- Select Parameter --
                                      </option>
                                      {parameters
                                        .filter(
                                          (param) => param.type === "main"
                                        )
                                        .map((param) => (
                                          <option
                                            key={param.id}
                                            value={param.name}
                                          >
                                            {param.name} ({param.type})
                                          </option>
                                        ))}
                                    </select>
                                  </div>
                                  <div>
                                    <label className="block text-xs font-semibold text-gray-600 dark:text-gray-400 mb-2 uppercase tracking-wide">
                                      Sub Parameter 2
                                    </label>
                                    <select
                                      value={question.subParam2 || ""}
                                      onChange={(e) =>
                                        updateQuestion(
                                          section.id,
                                          question.id,
                                          {
                                            subParam2: e.target.value,
                                          }
                                        )
                                      }
                                      className="w-full px-3 py-2.5 border-2 border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all text-sm dark:bg-gray-800 dark:text-gray-100"
                                    >
                                      <option value="">
                                        -- Select Parameter --
                                      </option>
                                      {parameters
                                        .filter(
                                          (param) => param.type === "followup"
                                        )
                                        .map((param) => (
                                          <option
                                            key={param.id}
                                            value={param.name}
                                          >
                                            {param.name} ({param.type})
                                          </option>
                                        ))}
                                    </select>
                                  </div>
                                </div>

                                {(question.type === "radio" ||
                                  question.type === "yesNoNA" ||
                                  question.type === "checkbox" ||
                                  question.type === "select" ||
                                  question.type === "search-select") && (
                                    <div className="mt-5 p-4 bg-blue-50 rounded-lg border border-blue-200 shadow-sm transition-all hover:shadow-md">
                                      {/* Active Config Status Indicators (Badges) */}
                                      <div className="flex flex-wrap gap-2 mb-4 empty:hidden">
                                        {question.branchingRules &&
                                          question.branchingRules.length > 0 && (
                                            <div className="group relative flex items-center gap-2 px-3 py-1.5 bg-purple-100 text-purple-700 rounded-lg text-[10px] font-bold border border-purple-200 shadow-sm">
                                              <span className="text-xs">🔀</span>
                                              <span>Routing Active ({question.branchingRules.length})</span>
                                            </div>
                                          )}
                                        {question.followUpConfig &&
                                          Object.values(
                                            question.followUpConfig
                                          ).some((c) => c.linkedFormId) && (
                                            <div className="group relative flex items-center gap-2 px-3 py-1.5 bg-green-100 text-green-700 rounded-lg text-[10px] font-bold border border-green-200 shadow-sm">
                                              <span className="text-xs">🔗</span>
                                              <span>Forms Linked</span>
                                            </div>
                                          )}
                                        {question.correctAnswer && (
                                          <div className="flex items-center gap-2 px-3 py-1.5 bg-blue-100 text-blue-700 rounded-lg text-[10px] font-bold border border-blue-200 shadow-sm">
                                            <span className="text-xs">✅</span>
                                            <span>Key: {question.correctAnswer}</span>
                                          </div>
                                        )}
                                      </div>

                                      <label className="block text-xs font-semibold text-blue-800 mb-3 uppercase tracking-wide opacity-70">
                                        Options
                                      </label>
                                      <div className="space-y-2.5">
                                        {(question.options || []).map(
                                          (option, index) => {
                                            const menuKey = `${section.id}-${question.id}-${index}`;
                                            const isMenuOpen =
                                              openOptionMenu === menuKey;
                                            const isYesNoNa =
                                              question.type === "yesNoNA" || question.type === "chassisNumber";

                                            return (
                                              <div
                                                key={index}
                                                className="flex items-center gap-2 group"
                                              >
                                                <div className="flex items-center justify-center w-6 h-6 rounded-full bg-blue-200 text-blue-700 font-bold text-xs">
                                                  {index + 1}
                                                </div>
                                                <input
                                                  type="text"
                                                  value={option}
                                                  onChange={(e) =>
                                                    updateOption(
                                                      section.id,
                                                      question.id,
                                                      index,
                                                      e.target.value
                                                    )
                                                  }
                                                  className="flex-1 px-3 py-2 border-2 border-gray-300 dark:border-gray-600 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all disabled:bg-gray-100 dark:bg-gray-700"
                                                  placeholder={`Option ${index + 1
                                                    }`}
                                                />
                                                {!isYesNoNa && (
                                                  <>
                                                    <button
                                                      onClick={() =>
                                                        duplicateOption(
                                                          section.id,
                                                          question.id,
                                                          index
                                                        )
                                                      }
                                                      className="p-2 text-blue-600 hover:text-blue-800 hover:bg-blue-100 rounded-lg transition-all"
                                                      title="Duplicate option"
                                                    >
                                                      <Clipboard className="w-5 h-5" />
                                                    </button>
                                                    <button
                                                      onClick={() =>
                                                        removeOption(
                                                          section.id,
                                                          question.id,
                                                          index
                                                        )
                                                      }
                                                      className="p-2 text-red-600 hover:text-red-800 hover:bg-red-100 rounded-lg transition-all"
                                                      title="Delete option"
                                                    >
                                                      <X className="w-5 h-5" />
                                                    </button>
                                                  </>
                                                )}

                                                {(
                                                  <div className="relative">
                                                    <button
                                                      onClick={() =>
                                                        setOpenOptionMenu(
                                                          isMenuOpen
                                                            ? null
                                                            : menuKey
                                                        )
                                                      }
                                                      className="p-2 text-gray-600 dark:text-gray-400 hover:text-gray-800 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-700 dark:bg-gray-700 rounded-lg transition-all"
                                                      title="Follow-up options"
                                                    >
                                                      <MoreVertical className="w-5 h-5" />
                                                    </button>

                                                    {isMenuOpen && (
                                                      <>
                                                        <div
                                                          className="fixed inset-0 z-10"
                                                          onClick={() =>
                                                            setOpenOptionMenu(
                                                              null
                                                            )
                                                          }
                                                        />

                                                        <div className="absolute right-0 mt-1 w-56 bg-white dark:bg-gray-800 rounded-lg shadow-xl border border-gray-200 dark:border-gray-700 z-20 overflow-hidden">
                                                          <div className="py-1">
                                                            <button
                                                              onClick={() => {
                                                                addFollowUpQuestion(
                                                                  section.id,
                                                                  question.id,
                                                                  option
                                                                );
                                                                setOpenOptionMenu(
                                                                  null
                                                                );
                                                              }}
                                                              className="w-full text-left px-4 py-2 text-sm text-gray-700 dark:text-gray-300 hover:bg-blue-50 dark:hover:bg-gray-700 flex items-center gap-2"
                                                            >
                                                              <span className="text-base">📝</span>
                                                              Follow-up Question
                                                            </button>

                                                            <button
                                                              onClick={() => {
                                                                openBranchingConfig(
                                                                  section.id,
                                                                  question.id,
                                                                  question.options || []
                                                                );
                                                                setOpenOptionMenu(
                                                                  null
                                                                );
                                                              }}
                                                              className="w-full text-left px-4 py-2 text-sm text-purple-700 dark:text-purple-300 hover:bg-purple-50 dark:hover:bg-gray-700 flex items-center gap-2"
                                                            >
                                                              <span className="text-base">🔀</span>
                                                              Section Routing
                                                            </button>

                                                            <button
                                                              onClick={() => {
                                                                openFormRoutingConfig(
                                                                  section.id,
                                                                  question.id,
                                                                  question.options || []
                                                                );
                                                                setOpenOptionMenu(
                                                                  null
                                                                );
                                                              }}
                                                              className="w-full text-left px-4 py-2 text-sm text-green-700 dark:text-green-300 hover:bg-green-50 dark:hover:bg-gray-700 flex items-center gap-2"
                                                            >
                                                              <span className="text-base">🔗</span>
                                                              Form Routing
                                                            </button>

                                                            <button
                                                              onClick={() => {
                                                                updateQuestion(
                                                                  section.id,
                                                                  question.id,
                                                                  {
                                                                    correctAnswer:
                                                                      question.correctAnswer ===
                                                                        option
                                                                        ? undefined
                                                                        : option,
                                                                  }
                                                                );
                                                                setOpenOptionMenu(
                                                                  null
                                                                );
                                                              }}
                                                              className={`w-full text-left px-4 py-2 text-sm flex items-center gap-2 ${question.correctAnswer ===
                                                                option
                                                                ? "text-green-600 bg-green-50 dark:bg-green-900/30"
                                                                : "text-gray-700 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700"
                                                                }`}
                                                            >
                                                              <span className="text-base">
                                                                {question.correctAnswer ===
                                                                  option
                                                                  ? "✅"
                                                                  : "✔️"}
                                                              </span>
                                                              {question.correctAnswer ===
                                                                option
                                                                ? "Correct Answer"
                                                                : "Mark as Correct"}
                                                            </button>

                                                            {/* Tenant Visibility for Chassis Number */}
                                                            {question.type === "chassisNumber" && (
                                                              <button
                                                                onClick={() => {
                                                                  const copiedSections = [...form.sections];
                                                                  const currentSection = copiedSections.find(s => s.id === section.id);
                                                                  if (currentSection) {
                                                                    const currentQuestion = currentSection.questions.find(q => q.id === question.id);
                                                                    if (currentQuestion && currentQuestion.options) {
                                                                      setChassisTenantConfig({
                                                                        questionId: question.id,
                                                                        sectionId: section.id,
                                                                        optionIndex: index,
                                                                        chassisNumber: currentQuestion.options[index],
                                                                      });
                                                                      setShowChassisTenantModal(true);
                                                                      setOpenOptionMenu(null);
                                                                    }
                                                                  }
                                                                }}
                                                                className="w-full text-left px-4 py-2 text-sm text-purple-700 dark:text-purple-300 hover:bg-purple-50 dark:hover:bg-gray-700 flex items-center gap-2"
                                                              >
                                                                <span className="text-base">👥</span>
                                                                Set Tenant Visibility
                                                              </button>
                                                            )}
                                                          </div>
                                                        </div>
                                                      </>
                                                    )}
                                                  </div>
                                                )}
                                              </div>
                                            );
                                          }
                                        )}
                                        {question.type !== "yesNoNA" && (
                                          <button
                                            onClick={() =>
                                              addOption(section.id, question.id)
                                            }
                                            className="flex items-center gap-2 px-4 py-2 text-blue-700 hover:text-blue-900 hover:bg-blue-100 rounded-lg text-sm font-medium transition-all"
                                          >
                                            <Plus className="w-5 h-5" />
                                            Add Option
                                          </button>
                                        )}
                                      </div>
                                    </div>
                                  )}

                                {/* Follow-up Questions Section - Now with Unlimited Nesting */}
                                {question.followUpQuestions &&
                                  question.followUpQuestions.length > 0 && (
                                    <div className="mt-6">
                                      <NestedFollowUpRenderer
                                        followUpQuestions={
                                          question.followUpQuestions || []
                                        }
                                        sectionId={section.id}
                                        parentQuestion={{
                                          id: question.id,
                                          options: question.options,
                                        }}
                                        path={[question.id]}
                                        parameters={parameters}
                                        onUpdate={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          updates,
                                          nestedPath
                                        ) =>
                                          updateNestedFollowUpQuestion(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            updates,
                                            nestedPath
                                          )
                                        }
                                        onImageUpload={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          file,
                                          nestedPath
                                        ) =>
                                          handleNestedFollowUpImageUpload(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            file,
                                            nestedPath
                                          )
                                        }
                                        onImageRemove={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          nestedPath
                                        ) =>
                                          clearNestedFollowUpImage(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            nestedPath
                                          )
                                        }
                                        onDelete={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          nestedPath
                                        ) =>
                                          deleteNestedFollowUpQuestion(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            nestedPath
                                          )
                                        }
                                        onAddNested={(
                                          nestedSectionId,
                                          parentFollowUpId,
                                          triggerValue,
                                          nestedPath
                                        ) =>
                                          addNestedFollowUpQuestion(
                                            nestedSectionId,
                                            parentFollowUpId,
                                            triggerValue,
                                            nestedPath
                                          )
                                        }
                                        onAddOption={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          nestedPath
                                        ) =>
                                          addNestedFollowUpOption(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            nestedPath
                                          )
                                        }
                                        onUpdateOption={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          optionIndex,
                                          value,
                                          nestedPath
                                        ) =>
                                          updateNestedFollowUpOption(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            optionIndex,
                                            value,
                                            nestedPath
                                          )
                                        }
                                        onRemoveOption={(
                                          nestedSectionId,
                                          followUpQuestionId,
                                          optionIndex,
                                          nestedPath
                                        ) =>
                                          removeNestedFollowUpOption(
                                            nestedSectionId,
                                            followUpQuestionId,
                                            optionIndex,
                                            nestedPath
                                          )
                                        }
                                        onAddFollowUpSection={
                                          handleAddFollowUpSection
                                        }
                                        onAddFollowUpForm={
                                          handleAddFollowUpForm
                                        }
                                        questionTypes={questionTypes}
                                      />
                                    </div>
                                  )}
                              </div>
                            </div>
                          </React.Fragment>
                        );
                      })}

                      {/* After Section Action / Section Routing */}
                      {!section.isSubsection && (
                        <div className="mt-8 pt-6 border-t border-gray-100 dark:border-gray-800">
                          <div className="flex items-center gap-2 mb-4">
                            <div className="w-8 h-8 rounded-full bg-blue-100 dark:bg-blue-900/40 flex items-center justify-center">
                              <LinkIcon className="w-4 h-4 text-blue-600 dark:text-blue-400" />
                            </div>
                            <h4 className="text-sm font-bold text-gray-900 dark:text-white uppercase tracking-wider">
                              After Section Action
                            </h4>
                          </div>

                          {hasSectionBranching(section.id) ? (
                            <div className="p-4 bg-orange-50 border border-orange-200 rounded-xl">
                              <p className="text-sm text-orange-800 flex items-center gap-2">
                                <span className="w-2 h-2 bg-orange-500 rounded-full animate-pulse"></span>
                                <strong>Direct Link Disabled:</strong> This section has options linked to other sections via branching. Direct linking is only available when no options have custom routing.
                              </p>
                            </div>
                          ) : (
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                              <button
                                onClick={() => updateSection(section.id, { nextSectionId: undefined })}
                                className={`flex flex-col items-center gap-2 p-4 rounded-xl border-2 transition-all ${!section.nextSectionId
                                  ? "bg-blue-50 border-blue-500 text-blue-700 shadow-sm"
                                  : "bg-white dark:bg-gray-900 border-gray-100 dark:border-gray-800 text-gray-500 hover:border-blue-200"
                                  }`}
                              >
                                <div className={`w-10 h-10 rounded-full flex items-center justify-center ${!section.nextSectionId ? "bg-blue-600 text-white" : "bg-gray-100 dark:bg-gray-800"
                                  }`}>
                                  <ChevronDown className="w-5 h-5" />
                                </div>
                                <span className="text-xs font-bold uppercase tracking-wide">Continue to Next</span>
                              </button>

                              <button
                                onClick={() => updateSection(section.id, { nextSectionId: "end" })}
                                className={`flex flex-col items-center gap-2 p-4 rounded-xl border-2 transition-all ${section.nextSectionId === "end"
                                  ? "bg-red-50 border-red-500 text-red-700 shadow-sm"
                                  : "bg-white dark:bg-gray-900 border-gray-100 dark:border-gray-800 text-gray-500 hover:border-red-200"
                                  }`}
                              >
                                <div className={`w-10 h-10 rounded-full flex items-center justify-center ${section.nextSectionId === "end" ? "bg-red-600 text-white" : "bg-gray-100 dark:bg-gray-800"
                                  }`}>
                                  <X className="w-5 h-5" />
                                </div>
                                <span className="text-xs font-bold uppercase tracking-wide">End Form</span>
                              </button>

                              <div className={`relative flex flex-col gap-2 p-4 rounded-xl border-2 transition-all ${!!section.nextSectionId && section.nextSectionId !== "end"
                                ? "bg-purple-50 border-purple-500 text-purple-700 shadow-sm"
                                : "bg-white dark:bg-gray-900 border-gray-100 dark:border-gray-800 text-gray-500 hover:border-purple-200"
                                }`}>
                                <div className="flex flex-col items-center gap-2 cursor-pointer"
                                  onClick={() => {
                                    const firstAvailable = form.sections.find(s => s.id !== section.id && !s.isSubsection);
                                    if (firstAvailable) {
                                      updateSection(section.id, { nextSectionId: firstAvailable.id });
                                    }
                                  }}>
                                  <div className={`w-10 h-10 rounded-full flex items-center justify-center ${!!section.nextSectionId && section.nextSectionId !== "end" ? "bg-purple-600 text-white" : "bg-gray-100 dark:bg-gray-800"
                                    }`}>
                                    <LinkIcon className="w-5 h-5" />
                                  </div>
                                  <span className="text-xs font-bold uppercase tracking-wide">Jump to Section</span>
                                </div>

                                {!!section.nextSectionId && section.nextSectionId !== "end" && (
                                  <select
                                    value={section.nextSectionId}
                                    onChange={(e) => updateSection(section.id, { nextSectionId: e.target.value })}
                                    className="mt-2 w-full px-2 py-1 text-xs border border-purple-200 rounded bg-white dark:bg-gray-800 focus:ring-2 focus:ring-purple-500 outline-none font-semibold"
                                  >
                                    {form.sections
                                      .filter((s) => s.id !== section.id && !s.isSubsection)
                                      .map((s, idx) => (
                                        <option key={s.id} value={s.id}>
                                          Section {String.fromCharCode(65 + form.sections.indexOf(s))}: {s.title || "Untitled"}
                                        </option>
                                      ))}
                                  </select>
                                )}
                              </div>
                            </div>
                          )}
                          <p className="mt-3 text-[10px] text-gray-500 dark:text-gray-400 font-medium">
                            * This setting determines where respondents go after completing all questions in this section.
                          </p>
                        </div>
                      )}

                      {/* Add question button at the end */}
                      {section.questions.length > 0 && (
                        <div className="flex justify-center -my-3 relative z-10">
                          <button
                            onClick={() => addQuestion(section.id)}
                            className="group bg-gradient-to-r from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 border-2 border-blue-400 rounded-full p-2 text-white transition-all duration-200 shadow-lg hover:shadow-xl transform hover:scale-110"
                            title="Add question at the end"
                          >
                            <Plus className="w-5 h-5" />
                          </button>
                        </div>
                      )}

                      {/* Or large button if no questions */}
                      {section.questions.length === 0 && (
                        <div className="text-center py-12">
                          <div
                            className="inline-flex flex-col items-center gap-3 p-8 border-2 border-dashed border-blue-300 rounded-xl hover:border-blue-500 hover:bg-blue-50 transition-all cursor-pointer"
                            onClick={() => addQuestion(section.id)}
                          >
                            <div className="flex items-center justify-center w-16 h-16 rounded-full bg-gradient-to-r from-blue-500 to-blue-600 text-white shadow-lg">
                              <Plus className="w-8 h-8" />
                            </div>
                            <p className="text-lg font-semibold text-blue-700">
                              Add Your First Question
                            </p>
                            <p className="text-sm text-gray-600 dark:text-gray-400">
                              Click to start building your form
                            </p>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                );
              });
            })()}

            {/* Add Section Button */}
            <div className="mt-8">
              <button
                onClick={() => {
                  addSection();
                  handlePageChange(totalPages, totalPages + 1);
                }}
                className="w-full group relative overflow-hidden p-6 border-2 border-dashed border-blue-300 rounded-xl text-blue-700 hover:border-blue-500 hover:bg-gradient-to-r hover:from-blue-50 hover:to-blue-50 transition-all duration-200 shadow-sm hover:shadow-md"
              >
                <div className="flex items-center justify-center gap-3">
                  <div className="flex items-center justify-center w-10 h-10 rounded-full bg-blue-600 text-white shadow-md group-hover:scale-110 transition-transform">
                    <Plus className="w-6 h-6" />
                  </div>
                  <span className="text-lg font-bold">
                    Add New Page (Section)
                  </span>
                </div>
                <p className="text-sm text-gray-600 dark:text-gray-400 mt-2">
                  Create a new main section on a separate page
                </p>
              </button>

              {/* Child Forms Manager - Only show when editing existing form */}
              {id ? (
                <div className="mt-6">
                  <ChildFormsManager
                    parentFormId={id}
                    onUpdate={() => {
                      // Optionally refresh form data if needed
                      console.log("Child forms updated");
                    }}
                  />
                </div>
              ) : (
                <div className="mt-6 bg-blue-50 border-2 border-blue-200 rounded-xl p-6">
                  <div className="flex items-start space-x-3">
                    <LinkIcon className="w-6 h-6 text-blue-600 flex-shrink-0 mt-1" />
                    <div>
                      <h3 className="text-lg font-semibold text-blue-900 mb-2">
                        🔗 Create Follow-Up Forms (Child Forms)
                      </h3>
                      <p className="text-sm text-blue-800 mb-3">
                        Want to link follow-up forms that appear after users
                        complete this form?
                      </p>
                      <div className="bg-white dark:bg-gray-900 border border-blue-200 rounded-lg p-4 text-sm text-blue-900">
                        <p className="font-medium mb-2">📝 How it works:</p>
                        <ol className="list-decimal list-inside space-y-1 text-blue-800">
                          <li>Save this form first</li>
                          <li>Come back to edit it</li>
                          <li>Link existing forms as "child forms"</li>
                          <li>
                            Child forms will appear to users after completing
                            this parent form
                          </li>
                        </ol>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Sidebar */}
          <div className="w-full lg:w-80 flex-shrink-0 space-y-6">
            {/* Sticky Page Navigation */}
            {(() => {
              const pages = getPagesFromSections();
              if (pages.length > 1) {
                return (
                  <div className="sticky top-1/2 -translate-y-1/2 h-0">
                    <div className="bg-white dark:bg-gray-900 rounded-lg border-2 border-blue-200 shadow-lg p-4 -translate-y-1/2">
                      <h3 className="text-sm font-bold text-blue-900 mb-3 text-center">
                        📄 Pages
                      </h3>
                      <div className="space-y-2">
                        {pages.map((page, pageIndex) => {
                          const isCurrent = currentPage === pageIndex;
                          const sectionLabel = String.fromCharCode(
                            65 + pageIndex
                          ); // A, B, C...
                          const mainSection = page.find((s) => !s.isSubsection);
                          const questionCount = page.reduce(
                            (total, s) => total + s.questions.length,
                            0
                          );

                          return (
                            <button
                              key={pageIndex}
                              type="button"
                              onClick={() => handlePageChange(pageIndex)}
                              className={`
                                w-full p-3 rounded-lg text-left transition-all duration-200
                                ${isCurrent
                                  ? "bg-gradient-to-r from-blue-600 to-blue-700 text-white shadow-md scale-105"
                                  : "bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200"
                                }
                              `}
                              title={
                                mainSection?.title?.trim().length
                                  ? mainSection.title
                                  : `Page ${pageIndex + 1}`
                              }
                            >
                              <div className="flex items-center gap-3">
                                <div
                                  className={`
                                    flex items-center justify-center w-8 h-8 rounded-full font-bold text-sm flex-shrink-0
                                    ${isCurrent
                                      ? "bg-white text-blue-600"
                                      : "bg-blue-600 text-white"
                                    }
                                  `}
                                >
                                  {sectionLabel}
                                </div>
                                <div className="flex-1 min-w-0">
                                  <p
                                    className={`text-sm font-bold truncate ${isCurrent ? "text-white" : "text-blue-900"
                                      }`}
                                  >
                                    {mainSection?.title?.trim().length
                                      ? mainSection.title
                                      : `Page ${pageIndex + 1}`}
                                  </p>
                                  <p
                                    className={`text-xs ${isCurrent
                                      ? "text-blue-100"
                                      : "text-blue-600"
                                      }`}
                                  >
                                    {questionCount}{" "}
                                    {questionCount === 1
                                      ? "question"
                                      : "questions"}
                                  </p>
                                  {page.filter((s) => s.isSubsection).length >
                                    0 && (
                                      <div
                                        className={`text-xs mt-2 pt-2 border-t ${isCurrent
                                          ? "border-blue-400"
                                          : "border-blue-200"
                                          }`}
                                      >
                                        <p
                                          className={`font-semibold mb-1 ${isCurrent
                                            ? "text-blue-100"
                                            : "text-green-700"
                                            }`}
                                        >
                                          Merged:
                                        </p>
                                        {page
                                          .filter((s) => s.isSubsection)
                                          .map((subsection) => (
                                            <p
                                              key={subsection.id}
                                              className={`ml-2 ${isCurrent
                                                ? "text-blue-100"
                                                : "text-green-600"
                                                }`}
                                            >
                                              • {subsection.title || "Subsection"}
                                            </p>
                                          ))}
                                      </div>
                                    )}
                                  {(() => {
                                    const routings = [];
                                    page.forEach((section) => {
                                      section.questions.forEach((question) => {
                                        if (
                                          question.branchingRules &&
                                          question.branchingRules.length > 0
                                        ) {
                                          question.branchingRules.forEach(
                                            (rule) => {
                                              const targetSection =
                                                form.sections.find(
                                                  (s) =>
                                                    s.id ===
                                                    rule.targetSectionId
                                                );
                                              if (targetSection) {
                                                routings.push({
                                                  from:
                                                    section.title ||
                                                    `Section ${pageIndex + 1}`,
                                                  option: rule.optionLabel,
                                                  to:
                                                    targetSection.title ||
                                                    "Unknown",
                                                });
                                              }
                                            }
                                          );
                                        }
                                      });
                                    });

                                    return routings.length > 0 ? (
                                      <div
                                        className={`text-xs mt-2 pt-2 border-t ${isCurrent
                                          ? "border-blue-400"
                                          : "border-purple-200"
                                          }`}
                                      >
                                        <p
                                          className={`font-semibold mb-1 ${isCurrent
                                            ? "text-blue-100"
                                            : "text-purple-700"
                                            }`}
                                        >
                                          🔀 Routes:
                                        </p>
                                        {routings.map((routing, idx) => (
                                          <p
                                            key={idx}
                                            className={`ml-2 text-xs leading-tight ${isCurrent
                                              ? "text-blue-100"
                                              : "text-purple-600"
                                              }`}
                                          >
                                            "{routing.option}" → {routing.to}
                                          </p>
                                        ))}
                                      </div>
                                    ) : null;
                                  })()}
                                </div>
                              </div>
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                );
              }
              return null;
            })()}

            <div className="bg-white dark:bg-gray-900 rounded-lg border border-blue-100 shadow p-4">
              <h3 className="text-sm font-bold text-blue-900 mb-3">
                Form Statistics
              </h3>
              <div className="space-y-2 text-xs text-blue-700">
                <div className="flex items-center justify-between">
                  <span>Sections</span>
                  <span>{form.sections.length}</span>
                </div>
                <div className="flex items-center justify-between">
                  <span>Questions</span>
                  <span>
                    {form.sections.reduce(
                      (total, section) => total + section.questions.length,
                      0
                    )}
                  </span>
                </div>
                <div className="flex items-center justify-between">
                  <span>Follow-ups</span>
                  <span>
                    {form.sections.reduce((total, section) => {
                      const countQuestions = (
                        questions: (Question | FollowUpQuestion)[]
                      ): number =>
                        questions.reduce((count, q) => {
                          const nested = q.followUpQuestions
                            ? countQuestions(q.followUpQuestions)
                            : 0;
                          return (
                            count + (q.followUpQuestions?.length || 0) + nested
                          );
                        }, 0);

                      return total + countQuestions(section.questions);
                    }, 0)}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Parameter Modal */}
      <ParameterModal
        isOpen={showParameterModal}
        onClose={() => setShowParameterModal(false)}
        formId={id}
        existingParameters={id ? undefined : tempParameters}
        onParameterCreated={async (createdParams) => {
          if (id) {
            // For existing forms, refresh from API
            try {
              const response = await apiClient.getParameters({ formId: id });
              setParameters(response.parameters || []);
            } catch (error) {
              console.error("Failed to refresh parameters:", error);
            }
          } else {
            // For new forms, store temporarily
            setTempParameters(createdParams || []);
          }
        }}
      />

      {/* Section Branching Config Modal */}
      {showBranchingConfig && branchingConfigQuestion && (
        <SectionBranchingConfig
          questionId={branchingConfigQuestion.questionId}
          sectionId={branchingConfigQuestion.sectionId}
          options={branchingConfigQuestion.options}
          sections={form.sections}
          existingRules={
            form.sections
              .find((s) => s.id === branchingConfigQuestion.sectionId)
              ?.questions.find(
                (q) => q.id === branchingConfigQuestion.questionId
              )?.branchingRules || []
          }
          onSave={handleSaveBranchingRules}
          onUpdateSection={handleUpdateSection}
          onClose={() => {
            setShowBranchingConfig(false);
            setBranchingConfigQuestion(null);
          }}
        />
      )}

      {/* Form Routing Config Modal */}
      {showFormRoutingConfig && formRoutingConfigQuestion && (
        <FormRoutingConfig
          questionId={formRoutingConfigQuestion.questionId}
          sectionId={formRoutingConfigQuestion.sectionId}
          options={formRoutingConfigQuestion.options}
          availableForms={forms.map((f) => ({ id: f.id, title: f.title }))}
          existingConfig={
            form.sections
              .find((s) => s.id === formRoutingConfigQuestion.sectionId)
              ?.questions.find(
                (q) => q.id === formRoutingConfigQuestion.questionId
              )?.followUpConfig || {}
          }
          onSave={handleSaveFormRoutingConfig}
          onClose={() => {
            setShowFormRoutingConfig(false);
            setFormRoutingConfigQuestion(null);
          }}
        />
      )}

      {/* Chassis Number Tenant Configuration Modal */}
      {showChassisTenantModal && chassisTenantConfig && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white dark:bg-gray-800 rounded-xl p-6 w-full max-w-md shadow-2xl">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-lg font-bold text-gray-900 dark:text-white">
                Set Tenant Visibility
              </h3>
              <button
                onClick={() => setShowChassisTenantModal(false)}
                className="p-1 hover:bg-gray-100 dark:hover:bg-gray-700 rounded"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <p className="text-sm text-gray-600 dark:text-gray-400 mb-4">
              Select which tenants can see and respond to this chassis number option.
            </p>

            {/* Get current option value */}
            {(() => {
              const currentChassis = chassisTenantConfig.chassisNumber;
              const configKey = currentChassis;

              return (
                <>
                  <div className="mb-4 p-3 bg-blue-50 dark:bg-blue-900/20 rounded-lg">
                    <span className="text-sm font-medium text-blue-700 dark:text-blue-300">
                      Chassis Number: {currentChassis}
                    </span>
                  </div>

                  <div className="max-h-64 overflow-y-auto border-2 border-gray-200 dark:border-gray-700 rounded-lg p-3 mb-4">
                    {tenants.length > 0 ? (
                      tenants.map((tenant) => {
                        const isSelected = (chassisTenantAssignments[configKey] || []).includes(tenant._id);
                        return (
                          <label
                            key={tenant._id}
                            className="flex items-center gap-3 p-2 hover:bg-gray-50 dark:hover:bg-gray-700 rounded-md cursor-pointer transition-colors"
                          >
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={(e) => {
                                setChassisTenantAssignments(prev => {
                                  const current = prev[configKey] || [];
                                  if (e.target.checked) {
                                    return { ...prev, [configKey]: [...current, tenant._id] };
                                  } else {
                                    return { ...prev, [configKey]: current.filter(id => id !== tenant._id) };
                                  }
                                });
                              }}
                              className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                            />
                            <div className="flex flex-col">
                              <span className="text-sm font-medium text-gray-700 dark:text-gray-300">
                                {tenant.companyName || tenant.name}
                              </span>
                              <span className="text-xs text-gray-500">
                                {tenant.name}
                              </span>
                            </div>
                          </label>
                        );
                      })
                    ) : (
                      <div className="py-8 text-center">
                        <Users className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                        <p className="text-sm text-gray-500 dark:text-gray-400">
                          No tenants found or still loading...
                        </p>
                        <button
                          onClick={async () => {
                            try {
                              const response = await apiClient.getTenantsMinimal();
                              setTenants(response.tenants || []);
                            } catch (err) {
                              console.error("Failed to fetch tenants:", err);
                              showError("Failed to fetch tenants", "Error");
                            }
                          }}
                          className="mt-2 text-xs text-blue-600 hover:text-blue-500 font-medium underline"
                        >
                          Retry Fetching
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="flex justify-end gap-2">
                    <button
                      onClick={() => setShowChassisTenantModal(false)}
                      className="px-4 py-2 text-gray-700 hover:bg-gray-100 dark:text-gray-300 dark:hover:bg-gray-700 rounded-lg transition-colors"
                    >
                      Cancel
                    </button>
                    <button
                      onClick={() => {
                        setShowChassisTenantModal(false);
                        showSuccess("Tenant visibility updated", "Success");
                      }}
                      className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
                    >
                      Save
                    </button>
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}
    </div>
  );
}
