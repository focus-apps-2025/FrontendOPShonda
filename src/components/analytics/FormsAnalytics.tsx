import React, {
  useState,
  useMemo,
  useRef,
  useEffect,
  ChangeEvent,
} from "react";
import { useNavigate } from "react-router-dom";
import {
  FileText,
  Eye,
  Users,
  Calendar,
  Layers,
  ChevronRight,
  Trash2,
  Edit2,
  PlusCircle,
  Search,
  Copy,
  BarChart3,
  List,
  MoreVertical,
  Link2,
  Share2,
  Check,
  Upload,
  Download,
  MapPin,
  X,
  Save,
  ChevronDown,
  Folder,
  Layout,
  Split,
} from "lucide-react";
import { useForms, useResponses, useMutation } from "../../hooks/useApi";
import { apiClient } from "../../api/client";
import { useNotification } from "../../context/NotificationContext";
import {
  downloadFormImportTemplate,
  downloadNestedFormImportTemplate,
  parseFormWorkbook
} from "../../utils/exportUtils";
import AnswerTemplateImport from "../AnswerTemplateImport";
import type { Question as FormQuestion } from "../../types";
import { Mail, MessageCircle, MessageSquare } from "lucide-react";
import EmailInviteModal from "../EmailInviteModal";
import WhatsAppInviteModal from "../WhatsAppInviteModal";
import SMSInviteModal from "../SMSInviteModal";
import ShareAnalyticsModal from "./ShareAnalyticsModal";
import AutoSendModal from "../forms/AutoSendModal";

import { useAuth } from "../../context/AuthContext";

// Add this interface for the dropdown options
interface TemplateOption {
  id: "flat" | "nested" | "linking";
  label: string;
  description: string;
}

interface FormItem {
  _id: string;
  id?: string;
  title: string;
  tenantId?: string;
  chassisTenantAssignments?: any[];
  description?: string;
  isVisible?: boolean;
  locationEnabled?: boolean;
  isActive?: boolean;
  viewType?: "section-wise" | "question-wise";
  sections?: any[];
  questions?: any[];
  createdAt?: string;
  createdBy?: any;
  responseCount?: number;
  parentFormId?: string | null;
  childForms?: Array<{
    formId: string;
    formTitle?: string;
    order?: number;
  }>;
}

interface ResponseData {
  responses: any[];
}

export default function FormsAnalytics() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const isInspector = user?.role === "inspector";
  const canManage = (user?.role === "admin" || user?.role === "superadmin" || user?.role === "subadmin") && !isInspector;
  const { showSuccess, showError, showConfirm } = useNotification();
  const [searchTerm, setSearchTerm] = useState("");
  const [openMenuId, setOpenMenuId] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [isAnswerTemplateOpen, setIsAnswerTemplateOpen] = useState(false);
  const [previewFormData, setPreviewFormData] = useState<FormQuestion | null>(
    null
  );
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [isSavingForm, setIsSavingForm] = useState(false);
   // Add these states for template dropdown
  const [isTemplateDropdownOpen, setIsTemplateDropdownOpen] = useState(false);
  const [selectedTemplate, setSelectedTemplate] = useState<TemplateOption | null>(null);

   const templateOptions: TemplateOption[] = [
    {
      id: "flat",
      label: "Follow-up Only",
      description: "Flat structure with unlimited main follow-ups (FU1-FU99)"
    },
    {
      id: "nested",
      label: "Nested Follow-up",
      description: "Hierarchical structure with nested follow-ups (FU1.1, FU1.1.1)"
    }
  ];

    useEffect(() => {
    if (templateOptions.length > 0 && !selectedTemplate) {
      setSelectedTemplate(templateOptions[0]);
    }
  }, []);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (templateDropdownRef.current && 
          !templateDropdownRef.current.contains(event.target as Node) &&
          menuRef.current && 
          !menuRef.current.contains(event.target as Node)) {
        setIsTemplateDropdownOpen(false);
      }
    };

    if (isTemplateDropdownOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isTemplateDropdownOpen]);


  const fileInputRef = useRef<HTMLInputElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const templateDropdownRef = useRef<HTMLDivElement>(null); 

  const {
    data: formsData,
    loading,
    error,
    execute: refetchForms,
  } = useForms(!isAnswerTemplateOpen);

  const {
    data: responsesData,
    loading: responsesLoading,
    execute: refetchResponses,
  } = useResponses();

  const deleteMutation = useMutation((id: string) => apiClient.deleteForm(id), {
    onSuccess: () => {
      refetchForms();
    },
  });

  const duplicateMutation = useMutation(
    (id: string) => apiClient.duplicateForm(id),
    {
      onSuccess: () => {
        refetchForms();
      },
    }
  );

  const visibilityMutation = useMutation(
    ({ id, isVisible }: { id: string; isVisible: boolean }) =>
      apiClient.updateFormVisibility(id, isVisible),
    {
      onSuccess: () => {
        refetchForms();
      },
    }
  );

  const locationMutation = useMutation(
    ({ id, locationEnabled }: { id: string; locationEnabled: boolean }) =>
      apiClient.updateFormLocationEnabled(id, locationEnabled),
    {
      onSuccess: () => {
        refetchForms();
      },
      onError: (error: any) => {
        showError(
          error.message || "Failed to update location setting",
          "Error"
        );
      },
    }
  );

  const viewTypeMutation = useMutation(
    ({ id, viewType }: { id: string; viewType: "section-wise" | "question-wise" }) =>
      apiClient.updateFormViewType(id, viewType),
    {
      onSuccess: () => {
        refetchForms();
        showSuccess("Form view type updated successfully");
      },
      onError: (error: any) => {
        console.error("View Type Update Error:", error);
        showError(
          typeof error === "string" ? error : error.message || "Failed to update view type setting",
          "Error"
        );
      },
    }
  );

  const forms = formsData?.forms || [];
  const parentForms = forms.filter((form: FormItem) => !form.parentFormId);
  const totalForms = parentForms.length;

  const [emailInviteModal, setEmailInviteModal] = useState<{
    open: boolean;
    formId: string | null;
    formTitle: string;
  }>({ open: false, formId: null, formTitle: "" });

  const [whatsappInviteModal, setWhatsappInviteModal] = useState<{
    open: boolean;
    formId: string | null;
    formTitle: string;
  }>({ open: false, formId: null, formTitle: "" });

  const [smsInviteModal, setSmsInviteModal] = useState<{
    open: boolean;
    formId: string | null;
    formTitle: string;
  }>({ open: false, formId: null, formTitle: "" });

  const [shareAnalyticsModal, setShareAnalyticsModal] = useState<{
    open: boolean;
    formId: string | null;
    formTitle: string;
  }>({ open: false, formId: null, formTitle: "" });

  const [autoSendModal, setAutoSendModal] = useState<{
    open: boolean;
    formId: string | null;
    formTitle: string;
  }>({ open: false, formId: null, formTitle: "" });

  const [inviteCounts, setInviteCounts] = useState<Record<string, number>>({});

  // Add these functions with your other handlers
  const openEmailInviteModal = (formId: string) => {
    const form = forms.find((f) => f.id === formId || f._id === formId);
    if (form) {
      setEmailInviteModal({
        open: true,
        formId,
        formTitle: form.title,
      });
    }
  };

  const openWhatsAppInviteModal = (formId: string) => {
    const form = forms.find((f) => f.id === formId || f._id === formId);
    if (form) {
      setWhatsappInviteModal({
        open: true,
        formId,
        formTitle: form.title,
      });
    }
  };

  const openSMSInviteModal = (formId: string) => {
    const form = forms.find((f) => f.id === formId || f._id === formId);
    if (form) {
      setSmsInviteModal({
        open: true,
        formId,
        formTitle: form.title,
      });
    }
  };

  const openShareAnalyticsModal = (formId: string) => {
    const form = forms.find((f) => f.id === formId || f._id === formId);
    if (form) {
      setShareAnalyticsModal({
        open: true,
        formId,
        formTitle: form.title,
      });
    }
  };

  const openAutoSendModal = (formId: string) => {
    const form = forms.find((f) => f.id === formId || f._id === formId);
    if (form) {
      setAutoSendModal({
        open: true,
        formId,
        formTitle: form.title,
      });
    }
  };

  useEffect(() => {
    const fetchInviteCounts = async () => {
      try {
        const counts: Record<string, number> = {};

        // Loop through forms and get invite stats only for owned forms
        for (const form of forms) {
          const formId = form.id || form._id;
          if (formId) {
            // Check if user can access this form's invite stats
            const ownerTenantId = typeof form.tenantId === 'object' ? form.tenantId?._id : form.tenantId;
            const isOwner = user?.role === "superadmin" || !form.tenantId || ownerTenantId === user?.tenantId;

            if (isOwner) {
              try {
                const response = await apiClient.getInviteStats(formId);
                if (response.success) {
                  counts[formId] = response.data.invites?.total || 0;
                }
              } catch (error) {
                // If access denied or other error, skip this form
                console.warn(`Failed to fetch invite stats for form ${formId}:`, error);
                counts[formId] = 0;
              }
            } else {
              // User doesn't own this form, set count to 0
              counts[formId] = 0;
            }
          }
        }

        setInviteCounts(counts);
      } catch (error) {
        console.error("Failed to fetch invite counts:", error);
      }
    };

    if (forms.length > 0) {
      fetchInviteCounts();
    }
  }, [forms, user]);

  console.log("DEBUG: Total forms from API:", forms.length);
  console.log("DEBUG: Parent forms (no parentFormId):", parentForms.length);
  console.log(
    "DEBUG: Child forms (with parentFormId):",
    forms.filter((f: FormItem) => f.parentFormId).length
  );
  console.log(
    "DEBUG: All forms data:",
    forms.map((f: FormItem) => ({
      id: f._id || f.id,
      title: f.title,
      parentFormId: f.parentFormId,
    }))
  );
  const activeFormsCount = parentForms.filter(
    (form: FormItem) => form.isActive === true
  ).length;
  const inactiveFormsCount = parentForms.filter(
    (form: FormItem) => form.isActive === false
  ).length;

  const formsMap = useMemo(() => {
    const map = new Map<string, FormItem>();
    forms.forEach((form) => {
      if (form._id) {
        map.set(form._id, form);
      }
      if (form.id) {
        map.set(form.id, form);
      }
    });
    return map;
  }, [forms]);

  const filteredForms = forms.filter((form: FormItem) => {
    const titleMatch = form.title
      ?.toLowerCase()
      .includes(searchTerm.toLowerCase());
    const descriptionMatch = form.description
      ?.toLowerCase()
      .includes(searchTerm.toLowerCase());
    return titleMatch || descriptionMatch;
  });

  const responseCounts = useMemo(() => {
    const allResponses =
      (responsesData as ResponseData | undefined)?.responses || [];
    
    // Filter responses for inspectors to only count their own
    const filteredResponses = user?.role === "inspector" 
      ? allResponses.filter((response: any) => {
          const creatorId = typeof response.createdBy === 'object' ? response.createdBy?._id || response.createdBy?.id : response.createdBy;
          const userId = user._id || user.id;
          const submittedBy = response.submittedBy || '';
          const submitterEmail = response.submitterContact?.email || '';
          const userEmail = user.email || '';
          const userUsername = user.username || '';
          
          // Match by createdBy or submittedBy or submitterContact.email
          return String(creatorId) === String(userId) ||
                 submittedBy === userEmail ||
                 submittedBy === userUsername ||
                 submitterEmail === userEmail;
        })
      : allResponses;
    
    return filteredResponses.reduce<Record<string, number>>((acc, response: any) => {
      if (response.questionId) {
        acc[response.questionId] = (acc[response.questionId] || 0) + 1;
      }
      return acc;
    }, {});
  }, [responsesData, user]);

  const groupedForms = useMemo(() => {
    const result = filteredForms.reduce((acc, form) => {
      const key = form.parentFormId || form.id || form._id;
      if (!key) {
        return acc;
      }

      if (!acc[key]) {
        acc[key] = {
          parent: form.parentFormId ? null : form,
          children: [],
        };
      }

      if (form.parentFormId) {
        const parentKey = form.parentFormId;
        acc[parentKey] = acc[parentKey] || {
          parent: null,
          children: [],
        };
        acc[parentKey].children.push(form);
      } else {
        acc[key].parent = form;
      }

      return acc;
    }, {} as Record<string, { parent: FormItem | null; children: FormItem[] }>);

    Object.values(result).forEach((group) => {
      const parent = group.parent;
      if (!parent) {
        return;
      }

      const childRefs = [...(parent.childForms || [])].sort(
        (a, b) => (a.order ?? 0) - (b.order ?? 0)
      );

      if (childRefs.length === 0) {
        return;
      }

      const existingChildrenMap = new Map<string, FormItem>();
      group.children.forEach((child) => {
        const childKey = child.id || child._id;
        if (childKey) {
          existingChildrenMap.set(childKey, child);
        }
      });

      const orderedChildren: FormItem[] = [];
      const usedChildIds = new Set<string>();

      childRefs.forEach((childRef, index) => {
        const childId = childRef.formId;
        if (!childId || usedChildIds.has(childId)) {
          return;
        }

        usedChildIds.add(childId);

        let child = existingChildrenMap.get(childId) || formsMap.get(childId);
        if (!child) {
          child = {
            _id: childId,
            id: childId,
            title: childRef.formTitle || "Linked Form",
            parentFormId: parent.id || parent._id || null,
          } as FormItem;
        }

        orderedChildren.push(child);
      });

      group.children.forEach((child) => {
        const childId = child.id || child._id;
        if (!childId || usedChildIds.has(childId)) {
          return;
        }
        orderedChildren.push(child);
      });

      group.children = orderedChildren;
    });

    return result;
  }, [filteredForms, formsMap]);

  const allForms = filteredForms.length;
  const totalResponses = filteredForms.reduce((sum, form) => {
    const formId = form.id || form._id;
    // For inspectors, only count their own responses
    const count = user?.role === "inspector" 
      ? (responseCounts[formId] || 0) 
      : (responseCounts[formId] || form.responseCount || 0);
    return sum + count;
  }, 0);

  const handleDelete = async (id: string, title: string) => {
    showConfirm(
      `Are you sure you want to delete "${title}"? This action cannot be undone.`,
      async () => {
        await deleteMutation.mutate(id);
        showSuccess("Form deleted successfully", "Success");
      },
      "Delete Form",
      "Delete",
      "Cancel"
    );
  };

  const handleDuplicate = async (id: string) => {
    await duplicateMutation.mutate(id);
  };

  const handleToggleVisibility = async (
    id: string,
    currentVisibility: boolean | undefined
  ) => {
    await visibilityMutation.mutate({
      id,
      isVisible: !currentVisibility,
    });
  };

  const handleToggleLocation = async (
    id: string,
    currentLocationEnabled: boolean | undefined
  ) => {
    const isCurrentlyEnabled = currentLocationEnabled !== false;
    await locationMutation.mutate({
      id,
      locationEnabled: !isCurrentlyEnabled,
    });
  };

  const handleToggleViewType = (
    id: string,
    currentViewType: "section-wise" | "question-wise" | undefined
  ) => {
    console.log("Toggling view type for ID:", id, "current:", currentViewType);
    const nextViewType =
      currentViewType === "question-wise" ? "section-wise" : "question-wise";
    
    viewTypeMutation.mutate({
      id,
      viewType: nextViewType,
    });
  };

  const handleExportTemplate = async (templateId?: "flat" | "nested") => {
    const templateToUse = templateId || (selectedTemplate?.id as "flat" | "nested");
    
    if (templateToUse === "nested") {
      downloadNestedFormImportTemplate();
      showSuccess("Nested Follow-up template downloaded", "Success");
    } else {
      await downloadFormImportTemplate();
      showSuccess("OPS template downloaded", "Success");
    }
    
    setIsTemplateDropdownOpen(false);
  };
  
   // Handle template selection
  const handleTemplateSelect = (template: TemplateOption) => {
    setSelectedTemplate(template);
    handleExportTemplate(template.id);
  };

  // Toggle template dropdown
  const toggleTemplateDropdown = () => {
    setIsTemplateDropdownOpen(!isTemplateDropdownOpen);
  };



  const handleFileInputChange = async (
    event: ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    const isValidType =
      file.type ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.name.toLowerCase().endsWith(".xlsx");
    if (!isValidType) {
      showError("Please select a valid .xlsx file", "Invalid File");
      return;
    }

    setIsImporting(true);

    try {
      const parsed = await parseFormWorkbook(file);
      const formPayload = {
        ...parsed,
        isVisible: parsed.isVisible ?? true,
        followUpQuestions: parsed.followUpQuestions || [],
      } as FormQuestion;

      setPreviewFormData(formPayload);
      setIsPreviewOpen(true);
    } catch (error: any) {
      showError(error?.message || "Failed to parse form", "Import Failed");
    } finally {
      setIsImporting(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const handleConfirmImport = async () => {
    if (!previewFormData) return;

    setIsSavingForm(true);

    try {
      await apiClient.createForm(previewFormData);
      showSuccess("Form imported successfully", "Import Complete");
      refetchForms();
      setIsPreviewOpen(false);
      setPreviewFormData(null);
    } catch (error: any) {
      showError(error?.message || "Failed to import form", "Import Failed");
    } finally {
      setIsSavingForm(false);
    }
  };

  const handleCancelImport = () => {
    setIsPreviewOpen(false);
    setPreviewFormData(null);
  };

  const handleImportClick = () => {
    if (isImporting) {
      return;
    }
    fileInputRef.current?.click();
  };

  const toggleMenu = (formId: string) => {
    setOpenMenuId(openMenuId === formId ? null : formId);
  };

  const handleManageChildForms = (formId: string) => {
    // Navigate to edit page where ChildFormsManager is available
    navigate(`/forms/${formId}/edit`);
    setOpenMenuId(null);
    // Optionally scroll to child forms section after a short delay
    setTimeout(() => {
      const childFormsSection = document.querySelector(
        '[data-section="child-forms"]'
      );
      if (childFormsSection) {
        childFormsSection.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
      }
    }, 500);
  };

  const handleLinkToParent = (formId: string) => {
    // Navigate to edit page where user can manage parent-child relationships
    navigate(`/forms/${formId}/edit`);
    setOpenMenuId(null);
    setTimeout(() => {
      const childFormsSection = document.querySelector(
        '[data-section="child-forms"]'
      );
      if (childFormsSection) {
        childFormsSection.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
      }
    }, 500);
  };

  const handleCopyShareLink = (formId: string, tenantSlug?: string) => {
    const baseUrl = window.location.origin;
    const shareLink = tenantSlug
      ? `${baseUrl}/${tenantSlug}/form/${formId}`
      : `${baseUrl}/form/${formId}`;

    navigator.clipboard.writeText(shareLink).then(() => {
      setCopiedId(formId);
      setTimeout(() => setCopiedId(null), 2000);
    });
    setOpenMenuId(null);
  };

  // Close menu when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setOpenMenuId(null);
      }
    };

    if (openMenuId) {
      document.addEventListener("mousedown", handleClickOutside);
    }

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [openMenuId]);

  if (loading || responsesLoading) {
    return (
      <div className="p-6">
        <div className="text-center py-12">
          <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary-600 mx-auto"></div>
          <p className="mt-4 text-primary-600">Loading forms...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-6">
        <div className="text-center py-12">
          <p className="text-red-600">Error loading forms: {error}</p>
          <button onClick={() => refetchForms()} className="mt-4 btn-primary">
            Try Again
          </button>
        </div>
      </div>
    );
  }

 return (
    <div className="p-6 space-y-6">
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        className="hidden"
        onChange={handleFileInputChange}
      />
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-xl sm:text-2xl font-bold text-primary-800">
            Service Analytics
          </h1>
          <p className="text-xs sm:text-sm text-primary-600">
            Create, edit, and analyze service request forms
          </p>
        </div>
        <div className="flex flex-col sm:flex-row sm:items-center gap-2 w-full sm:w-auto">
          {canManage && !isInspector && (
            <>
              {/* Updated Template Download Button with Dropdown */}
              <div 
                className="relative w-full sm:w-auto"
                ref={templateDropdownRef}
              >
                <button
                  onClick={toggleTemplateDropdown}
                  className="btn-secondary flex items-center justify-center w-full sm:min-w-[240px]"
                >
                  <Download className="w-4 h-4 mr-2" />
                  <span className="truncate">
                    {selectedTemplate ? `Download ${selectedTemplate.label} Template` : "Download Import Template"}
                  </span>
                  <ChevronDown className={`w-4 h-4 ml-2 flex-shrink-0 transition-transform ${isTemplateDropdownOpen ? 'rotate-180' : ''}`} />
                </button>
                
                {/* Dropdown Menu */}
                {isTemplateDropdownOpen && (
                  <div className="absolute top-full left-0 mt-1 w-full sm:w-72 bg-white dark:bg-gray-900 rounded-lg shadow-xl border border-primary-200 py-2 z-50 animate-fadeIn">
                    <div className="px-4 py-2 border-b border-primary-100">
                      <p className="text-xs font-medium text-primary-700">Select Template Type:</p>
                    </div>
                    
                    {templateOptions.map((template) => (
                      <button
                        key={template.id}
                        onClick={() => handleTemplateSelect(template)}
                        className={`w-full flex flex-col items-start px-4 py-3 text-left hover:bg-primary-50 transition-colors ${selectedTemplate?.id === template.id ? 'bg-primary-50 border-l-4 border-primary-600' : ''}`}
                      >
                        <div className="flex items-center w-full">
                          <div className={`p-1.5 rounded-lg mr-3 ${selectedTemplate?.id === template.id ? 'bg-primary-100' : 'bg-primary-50'}`}>
                            {template.id === "flat" ? (
                              <Layers className="w-4 h-4 text-primary-600" />
                            ) : (
                              <Layers className="w-4 h-4 text-purple-600" />
                            )}
                          </div>
                          <div className="flex-1">
                            <div className="font-medium text-primary-800 text-sm">
                              {template.label}
                            </div>
                            <div className="text-[10px] text-primary-600 mt-0.5">
                              {template.description}
                            </div>
                          </div>
                          {selectedTemplate?.id === template.id && (
                            <Check className="w-4 h-4 text-primary-600 ml-2" />
                          )}
                        </div>
                      </button>
                    ))}
                    
                    <div className="px-4 py-2 border-t border-primary-100 mt-1">
                      <p className="text-[10px] text-primary-500">
                        {selectedTemplate?.id === "flat" 
                          ? "Each question can have unlimited main-level follow-ups" 
                          : "Supports hierarchical follow-up questions with nesting"}
                      </p>
                    </div>
                  </div>
                )}
              </div>
              
              {/* Rest of your buttons remain the same */}
              <button
                onClick={handleImportClick}
                className="btn-secondary flex items-center justify-center w-full sm:w-auto"
                disabled={isImporting}
              >
                <Upload className="w-4 h-4 mr-2" />
                {isImporting ? "Importing..." : "Import Form (Excel)"}
              </button>
              <button
                onClick={() => setIsAnswerTemplateOpen(true)}
                className="btn-secondary flex items-center justify-center w-full sm:w-auto"
                title="Import answer templates for testing"
              >
                <Upload className="w-4 h-4 mr-2" />
                Import Answers
              </button>
              <button
                onClick={() =>
                  navigate("/forms/create", { state: { mode: "create" } })
                }
                className="btn-primary flex items-center justify-center w-full sm:w-auto"
              >
                <PlusCircle className="w-4 h-4 mr-2" />
                Create New Service Form
              </button>
            </>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-6">
        <div className="card p-6">
          <div className="flex items-center">
            <div className="p-3 bg-primary-50 rounded-lg mr-4">
              <FileText className="w-6 h-6 text-primary-600" />
            </div>
            <div>
              <div className="text-2xl font-medium text-primary-600">
                {totalForms}
              </div>
              <div className="text-sm text-primary-500">Total Forms</div>
              {/* <div className="mt-2 text-xs text-primary-500 space-x-2">
                <span className="inline-flex items-center px-2 py-1 bg-green-50 text-green-700 rounded-full">
                  Active: {activeFormsCount}
                </span>
                <span className="inline-flex items-center px-2 py-1 bg-red-50 text-red-700 rounded-full">
                  Inactive: {inactiveFormsCount}
                </span>
              </div> */}
            </div>
          </div>
        </div>
        <div className="card p-6">
          <div className="flex items-center">
            <div className="p-3 bg-primary-50 rounded-lg mr-4">
              <Users className="w-6 h-6 text-primary-600" />
            </div>
            <div>
              <div className="text-2xl font-medium text-primary-600">
                {totalResponses}
              </div>
              <div className="text-sm text-primary-500">Total Responses</div>
            </div>
          </div>
        </div>
        <div className="card p-6">
          <div className="flex items-center">
            <div className="p-3 bg-primary-50 rounded-lg mr-4">
              <Layers className="w-6 h-6 text-primary-600" />
            </div>
            <div>
              <div className="text-2xl font-medium text-primary-600">
                {Object.keys(groupedForms).length}
              </div>
              <div className="text-sm text-primary-500">Form Groups</div>
            </div>
          </div>
        </div>
      </div>

      <div className="relative">
        <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-primary-400 w-4 h-4" />
        <input
          type="text"
          placeholder="Search service forms..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className="pl-10 input-field"
        />
      </div>

      {filteredForms.length === 0 ? (
        <div className="text-center py-12 bg-white dark:bg-gray-900 rounded-lg border border-neutral-200 dark:border-gray-700">
          <FileText className="w-12 h-12 text-primary-300 mx-auto mb-4" />
          <h3 className="text-lg font-medium text-primary-600 mb-2">
            {searchTerm
              ? "No service forms found"
              : "No service forms created yet"}
          </h3>
          <p className="text-primary-500 mb-6">
            {searchTerm
              ? "Try adjusting your search criteria"
              : ""}
          </p>
          {!searchTerm && canManage && !isInspector && (
            <button
              onClick={() => navigate("/forms/create")}
              className="btn-primary"
            >
              <PlusCircle className="w-4 h-4 mr-2" />
              Create Your First Form
            </button>
          )}
        </div>
      ) : (
        <div className="space-y-6">
          {Object.values(groupedForms).map(({ parent, children }) => {
            if (!parent) return null;

            const formId = parent.id || parent._id;
            const responseCount = user?.role === "inspector" 
              ? (responseCounts[formId] || 0) 
              : (responseCounts[formId] || parent.responseCount || 0);
            const isLocationEnabled = parent.locationEnabled !== false;

            const ownerTenantId = typeof parent.tenantId === 'object' ? parent.tenantId?._id : parent.tenantId;
const isOwner = user?.role === "superadmin" || !parent.tenantId || ownerTenantId === user?.tenantId;

// Add this - check if user can edit (only admin and superadmin)
const canEdit = user?.role === "admin" || user?.role === "superadmin";

// Add this - check if user can delete (only admin and superadmin)
const canDelete = user?.role === "admin" || user?.role === "superadmin";

const tenantName = typeof parent.tenantId === 'object' ? (parent.tenantId?.companyName || parent.tenantId?.name) : null;
            return (
              <div
                key={formId}
                className="card p-6 hover:border-primary-300 transition-colors duration-200"
              >
                <div className="flex items-start justify-between mb-4">
                  <div className="flex-1">
                    <div className="flex items-center gap-2 mb-2">
                       <h3 className="font-medium text-primary-800 line-clamp-2 mb-0">
                        {parent.title}
                      </h3>
                      {tenantName && ownerTenantId !== user?.tenantId && (
                        <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold bg-blue-50 text-blue-600 border border-blue-100 uppercase tracking-wider">
                          {tenantName}
                        </span>
                      )}
                    </div>
                    <p className="text-sm text-primary-600 line-clamp-2">
                      {parent.description}
                    </p>
                  </div>
                </div>

                <div className="flex flex-col sm:flex-row sm:items-center justify-between text-xs text-primary-500 mb-6 gap-4">
                  <div className="flex flex-wrap items-center gap-4">
                    <div className="flex items-center bg-primary-50 dark:bg-gray-800 px-2 py-1 rounded-md">
                      <Users className="w-3.5 h-3.5 mr-1.5 text-primary-600" />
                      <span className="font-medium text-primary-700">{responseCount}</span>
                      <span className="ml-1 text-primary-500">responses</span>
                    </div>
                    {isOwner && (
                      <div className="flex items-center gap-1 bg-white dark:bg-gray-900 border border-primary-100 rounded-lg p-0.5">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            openEmailInviteModal(formId);
                          }}
                          title="Send Email Invites"
                          className="p-1.5 rounded-md hover:bg-blue-50 transition-colors group relative"
                        >
                          <Mail className="w-4 h-4 text-blue-600 group-hover:text-blue-700" />
                          {inviteCounts[formId] > 0 && (
                            <span className="absolute -top-1 -right-1 bg-blue-500 text-white text-[8px] rounded-full w-4 h-4 flex items-center justify-center font-bold">
                              {inviteCounts[formId]}
                            </span>
                          )}
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            openWhatsAppInviteModal(formId);
                          }}
                          title="Send WhatsApp Invites"
                          className="p-1.5 rounded-md hover:bg-green-50 transition-colors group"
                        >
                          <MessageCircle className="w-4 h-4 text-green-600 group-hover:text-green-700" />
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            openSMSInviteModal(formId);
                          }}
                          title="Send SMS Invites"
                          className="p-1.5 rounded-md hover:bg-purple-50 transition-colors group"
                        >
                          <MessageSquare className="w-4 h-4 text-purple-600 group-hover:text-purple-700" />
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center justify-between sm:justify-end gap-3 w-full sm:w-auto border-t sm:border-t-0 border-primary-50 pt-3 sm:pt-0">
                  <div className="relative">
  <button
    onClick={() => setOpenMenuId(openMenuId === formId ? null : formId)}
    className="flex items-center gap-1 px-3 py-1.5 text-sm font-medium text-primary-600 border border-primary-200 rounded-lg hover:bg-primary-50 transition-colors"
  >
    <MoreVertical className="w-4 h-4" />
    <span>Options</span>
  </button>

  {openMenuId === formId && (
    <>
      {/* Backdrop */}
      <div 
        className="fixed inset-0 z-40" 
        onClick={() => setOpenMenuId(null)} 
      />
      
      {/* Dropdown — fixed on mobile, absolute on desktop */}
      <div className="
        fixed z-50
        left-4 right-4
        sm:absolute sm:left-auto sm:right-0 sm:w-64
        top-1/2 -translate-y-1/2
        sm:top-10 sm:translate-y-0
        bg-white dark:bg-gray-900 
        rounded-xl shadow-2xl 
        border border-primary-100 
        py-2 
        animate-fadeIn 
        overflow-hidden
        max-h-[80vh] overflow-y-auto
      ">
        <div className="px-4 py-2 border-b border-primary-50 mb-1 flex items-center justify-between">
  <span className="text-[10px] font-bold text-primary-400 uppercase tracking-wider">Form Actions</span>
  <button
    onClick={() => setOpenMenuId(null)}
    className="p-1 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-700 text-gray-400 hover:text-gray-600 transition-colors"
    title="Close"
  >
    <X className="w-4 h-4" />
  </button>
</div>
        <button
          onClick={() => handleManageChildForms(formId)}
          className="w-full flex items-center gap-3 px-4 py-2.5 text-sm text-primary-700 hover:bg-primary-50 transition-colors"
        >
          <div className="p-1.5 bg-primary-50 rounded-lg">
            <Layers className="w-4 h-4 text-primary-600" />
          </div>
          <div className="text-left flex-1">
            <div className="font-medium">Manage Child Forms</div>
            <div className="text-[10px] text-primary-500">Link & organize forms</div>
          </div>
          {children.length > 0 && (
            <span className="px-2 py-0.5 bg-primary-600 text-white text-[10px] font-bold rounded-full">
              {children.length}
            </span>
          )}
        </button>

        <button
          onClick={() => handleLinkToParent(formId)}
          className="w-full flex items-center gap-3 px-4 py-2.5 text-sm text-primary-700 hover:bg-primary-50 transition-colors"
        >
          <div className="p-1.5 bg-blue-50 rounded-lg">
            <Link2 className="w-4 h-4 text-blue-600" />
          </div>
          <div className="text-left">
            <div className="font-medium">Link to Parent</div>
            <div className="text-[10px] text-primary-500">Connect to existing form</div>
          </div>
        </button>

        <button
          onClick={() => handleToggleViewType(formId, parent.viewType)}
          className="w-full flex items-center gap-3 px-4 py-2.5 text-sm text-primary-700 hover:bg-primary-50 transition-colors"
        >
          <div className="p-1.5 bg-orange-50 rounded-lg">
            {parent.viewType === "question-wise" ? (
              <Layout className="w-4 h-4 text-orange-600" />
            ) : (
              <Split className="w-4 h-4 text-orange-600" />
            )}
          </div>
          <div className="text-left">
            <div className="font-medium">
              {parent.viewType === "question-wise" ? "Section-wise View" : "Question-wise View"}
            </div>
            <div className="text-[10px] text-primary-500">Change display layout</div>
          </div>
        </button>

        <div className="border-t border-primary-50 my-1"></div>
        <div className="px-4 py-2">
          <span className="text-[10px] font-bold text-primary-400 uppercase tracking-wider">Sharing</span>
        </div>

        <button
          onClick={() => handleCopyShareLink(formId)}
          className="w-full flex items-center gap-3 px-4 py-2.5 text-sm text-primary-700 hover:bg-primary-50 transition-colors"
        >
          <div className="p-1.5 bg-green-50 rounded-lg">
            {copiedId === formId ? (
              <Check className="w-4 h-4 text-green-600" />
            ) : (
              <Link2 className="w-4 h-4 text-green-600" />
            )}
          </div>
          <div className="text-left">
            <div className="font-medium">
              {copiedId === formId ? "Link Copied!" : "Copy Form Link"}
            </div>
            <div className="text-[10px] text-primary-500">Share with responders</div>
          </div>
        </button>

        <button
          onClick={() => openShareAnalyticsModal(formId)}
          className="w-full flex items-center gap-3 px-4 py-2.5 text-sm text-primary-700 hover:bg-primary-50 transition-colors"
        >
          <div className="p-1.5 bg-indigo-50 rounded-lg">
            <Share2 className="w-4 h-4 text-indigo-600" />
          </div>
          <div className="text-left">
            <div className="font-medium">Share Analytics</div>
            <div className="text-[10px] text-primary-500">Invite external viewers</div>
          </div>
        </button>
      </div>
    </>
  )}
</div>

                    <div className="flex items-center text-primary-400 font-medium">
                      <Calendar className="w-3.5 h-3.5 mr-1.5" />
                      {parent.createdAt
                        ? new Date(parent.createdAt).toLocaleDateString()
                        : "Unknown"}
                    </div>
                  </div>
                </div>

                <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 mb-4">
                  <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
                    <span
                      className={`inline-flex items-center px-1.5 py-0.5 sm:px-2 sm:py-1 rounded-full text-[10px] sm:text-xs font-medium ${
                        parent.isVisible
                          ? "bg-green-100 text-green-800"
                          : "bg-yellow-100 text-yellow-800"
                      }`}
                    >
                      {parent.isVisible ? "Public" : "Private"}
                    </span>
                    <span
                      className={`inline-flex items-center gap-1 px-1.5 py-0.5 sm:px-2 sm:py-1 rounded-full text-[10px] sm:text-xs font-medium ${
                        isLocationEnabled
                          ? "bg-blue-100 text-blue-800"
                          : "bg-neutral-200 text-neutral-700"
                      }`}
                    >
                      <MapPin className="w-3 h-3" />
                      {isLocationEnabled
                        ? "Location Enabled"
                        : "Location Disabled"}
                    </span>
                    <span
                      className={`inline-flex items-center gap-1 px-1.5 py-0.5 sm:px-2.5 sm:py-0.5 rounded-full text-[9px] sm:text-[10px] font-medium border ${
                        parent.viewType === "question-wise"
                          ? "bg-orange-100 text-orange-800 border-orange-200"
                          : "bg-blue-100 text-blue-800 border-blue-200"
                      }`}
                    >
                      {parent.viewType === "question-wise" ? (
                        <>
                          <Split className="w-2.5 h-2.5 sm:w-3 sm:h-3" />
                          Question-wise
                        </>
                      ) : (
                        <>
                          <Layout className="w-2.5 h-2.5 sm:w-3 sm:h-3" />
                          Section-wise
                        </>
                      )}
                    </span>
                  </div>
                  <div className="flex items-center gap-2 sm:gap-3">
                    <button
                      onClick={() =>
                        handleToggleVisibility(formId, parent.isVisible)
                      }
                      disabled={visibilityMutation.loading}
                      className={`relative inline-flex h-5 w-9 sm:h-6 sm:w-11 items-center rounded-full transition-colors focus:outline-none focus:ring-2 focus:ring-offset-2 ${
                        parent.isVisible
                          ? "bg-green-500 focus:ring-green-500"
                          : "bg-red-500 focus:ring-red-500"
                      } ${
                        visibilityMutation.loading
                          ? "opacity-50 cursor-not-allowed"
                          : "cursor-pointer"
                      }`}
                      title={
                        parent.isVisible
                          ? "Active - Click to deactivate"
                          : "Inactive - Click to activate"
                      }
                    >
                      <span
                        className={`inline-block h-3 w-3 sm:h-4 sm:w-4 transform rounded-full bg-white transition-transform ${
                          parent.isVisible ? "translate-x-5 sm:translate-x-6" : "translate-x-1"
                        }`}
                      />
                    </button>
                    <button
                      onClick={() =>
                        handleToggleLocation(formId, parent.locationEnabled)
                      }
                      disabled={locationMutation.loading}
                      className={`relative inline-flex h-5 w-9 sm:h-6 sm:w-11 items-center rounded-full transition-colors focus:outline-none focus:ring-2 focus:ring-offset-2 ${
                        isLocationEnabled
                          ? "bg-primary-600 focus:ring-primary-600"
                          : "bg-neutral-400 focus:ring-neutral-400"
                      } ${
                        locationMutation.loading
                          ? "opacity-50 cursor-not-allowed"
                          : "cursor-pointer"
                      }`}
                      title={
                        isLocationEnabled
                          ? "Location enabled - Click to disable"
                          : "Location disabled - Click to enable"
                      }
                    >
                      <span
                        className={`inline-block h-3 w-3 sm:h-4 sm:w-4 transform rounded-full bg-white transition-transform ${
                          isLocationEnabled ? "translate-x-5 sm:translate-x-6" : "translate-x-1"
                        }`}
                      />
                    </button>
                    <button
                      onClick={() =>
                        handleToggleViewType(parent._id, parent.viewType)
                      }
                      disabled={viewTypeMutation.loading}
                      className={`relative inline-flex h-5 w-9 sm:h-6 sm:w-11 items-center rounded-full transition-colors focus:outline-none focus:ring-2 focus:ring-offset-2 ${
                        parent.viewType === "question-wise"
                          ? "bg-orange-500 focus:ring-orange-500"
                          : "bg-blue-500 focus:ring-blue-500"
                      } ${
                        viewTypeMutation.loading
                          ? "opacity-50 cursor-not-allowed"
                          : "cursor-pointer"
                      }`}
                      title={
                        parent.viewType === "question-wise"
                          ? "Question-wise view - Click for Section-wise"
                          : "Section-wise view - Click for Question-wise"
                      }
                    >
                      <span
                        className={`inline-block h-3 w-3 sm:h-4 sm:w-4 transform rounded-full bg-white transition-transform ${
                          parent.viewType === "question-wise"
                            ? "translate-x-5 sm:translate-x-6"
                            : "translate-x-1"
                        }`}
                      />
                    </button>
                  </div>
                </div>

                <div className="flex flex-col gap-4">
                  <div className="flex flex-wrap items-center gap-2">
                    {isOwner && (
                      <button
                        onClick={() => navigate(`/forms/${formId}/preview`)}
                        className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700 flex items-center justify-center gap-1.5"
                      >
                        <Eye className="w-3.5 h-3.5" />
                        <span>View</span>
                      </button>
                    )}
                    {isOwner ? (
                      <>
                        {canEdit && (
                          <button
                            onClick={() => navigate(`/forms/${formId}/edit`)}
                            className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700 flex items-center justify-center gap-1.5"
                            title="Edit form"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                            <span>Edit</span>
                          </button>
                        )}
                        <button
                          onClick={() => navigate(`/forms/${formId}/analytics`)}
                          className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700 flex items-center justify-center gap-1.5"
                          title="View analytics"
                        >
                          <BarChart3 className="w-3.5 h-3.5" />
                          <span>Analytics</span>
                        </button>
                        <button
                          onClick={() => navigate(`/forms/${formId}/uploads`)}
                          className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700 flex items-center justify-center gap-1.5"
                          title="View uploads"
                        >
                          <Folder className="w-3.5 h-3.5" />
                          <span>Uploads</span>
                        </button>
                      </>
                    ) : (
                      <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                        <button
                          onClick={() => navigate(`/forms/${formId}/analytics`)}
                          className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-indigo-600 rounded-lg transition-colors hover:bg-indigo-700 flex items-center justify-center gap-1.5"
                          title="View Shared Responses"
                        >
                          <List className="w-3.5 h-3.5" />
                          <span>Analytics</span>
                        </button>
                        <button
                          onClick={() => navigate(`/forms/${formId}/uploads`)}
                          className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-indigo-600 rounded-lg transition-colors hover:bg-indigo-700 flex items-center justify-center gap-1.5"
                          title="View Shared Uploads"
                        >
                          <Folder className="w-3.5 h-3.5" />
                          <span>Uploads</span>
                        </button>
                      </div>
                    )}

                    {isOwner && canDelete && (
                      <div className="flex gap-2 w-full sm:w-auto mt-2 sm:mt-0">
                        <button
                          onClick={() => handleDuplicate(formId)}
                          className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700 flex items-center justify-center gap-1.5"
                          title="Duplicate form"
                          disabled={duplicateMutation.loading}
                        >
                          <Copy className="w-3.5 h-3.5" />
                          <span>Duplicate</span>
                        </button>
                        <button
                          onClick={() => handleDelete(formId, parent.title)}
                          className="flex-1 sm:flex-none px-3 py-2 text-xs sm:text-sm font-medium text-white bg-red-600 rounded-lg transition-colors hover:bg-red-700 flex items-center justify-center gap-1.5"
                          title="Delete form"
                          disabled={deleteMutation.loading}
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>Delete</span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                {children.length > 0 && (
                  <div className="border-t border-neutral-200 dark:border-gray-700 pt-6 mt-6 bg-gradient-to-r from-primary-50/30 to-purple-50/30 -mx-6 px-6 pb-6 rounded-b-lg">
                    <div className="flex items-center justify-between mb-5">
                      <div className="flex items-center">
                        <div className="p-2 bg-gradient-to-br from-primary-500 to-purple-500 rounded-lg mr-3 shadow-sm">
                          <Layers className="w-5 h-5 text-white" />
                        </div>
                        <div>
                          <h4 className="text-lg font-semibold text-primary-800 flex items-center">
                            Child Forms
                            <span className="ml-2 px-2.5 py-0.5 text-xs font-bold bg-gradient-to-r from-primary-500 to-purple-500 text-white rounded-full shadow-sm">
                              {children.length}
                            </span>
                          </h4>
                          <p className="text-xs text-primary-600 mt-0.5">
                            Connected follow-up forms
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      {children.map((child, index) => {
                        const childId = child.id || child._id;
                        // For inspectors, only count their own responses
                        const childResponseCount = childId
                          ? (user?.role === "inspector" ? (responseCounts[childId] || 0) : (responseCounts[childId] || child.responseCount || 0))
                          : (user?.role === "inspector" ? 0 : (child.responseCount || 0));

                        return (
                          <div
                            key={childId}
                            className="relative bg-white dark:bg-gray-900 rounded-xl p-4 border-2 border-primary-100 hover:border-primary-300 hover:shadow-lg transition-all duration-300 group transform hover:-translate-y-1"
                            style={{
                              animationDelay: `${index * 50}ms`,
                              animation: "fadeInUp 0.5s ease-out forwards",
                            }}
                          >
                            {/* Corner decoration */}
                            <div className="absolute top-0 right-0 w-16 h-16 bg-gradient-to-br from-primary-100 to-purple-100 rounded-bl-full opacity-50"></div>

                            <div className="relative">
                              <div className="flex items-start justify-between mb-3">
                                <div className="flex items-center space-x-2">
                                  <div className="p-2.5 bg-gradient-to-br from-primary-500 to-purple-500 rounded-lg shadow-md group-hover:scale-110 transition-transform duration-300">
                                    <FileText className="w-4 h-4 text-white" />
                                  </div>
                                  <span className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold bg-gradient-to-r from-primary-100 to-purple-100 text-primary-700 border border-primary-200">
                                    ✦ Child
                                  </span>
                                </div>
                              </div>

                              <h5 className="font-semibold text-primary-800 mb-2 line-clamp-2 text-sm group-hover:text-primary-600 transition-colors">
                                {child.title}
                              </h5>

                              {child.description && (
                                <p className="text-xs text-primary-600 mb-3 line-clamp-2">
                                  {child.description}
                                </p>
                              )}

                              <div className="flex items-center justify-between text-xs text-primary-600 mb-3 pb-3 border-b border-primary-100">
                                <div className="flex items-center space-x-1">
                                  <Users className="w-3.5 h-3.5 text-primary-500" />
                                  <span className="font-medium">
                                    {childResponseCount}
                                  </span>
                                  <span className="text-primary-500">
                                    responses
                                  </span>
                                </div>
                                {child.createdAt && (
                                  <div className="flex items-center space-x-1 text-primary-500">
                                    <Calendar className="w-3.5 h-3.5" />
                                    <span>
                                      {new Date(
                                        child.createdAt
                                      ).toLocaleDateString("en-US", {
                                        month: "short",
                                        day: "numeric",
                                      })}
                                    </span>
                                  </div>
                                )}
                              </div>

                              {/* Quick action buttons */}
                              <div className="flex flex-wrap items-center justify-between gap-2 mt-auto pt-3 border-t border-primary-100">
                                {isOwner && (
                                  <>
                                    <button
                                      onClick={() =>
                                        navigate(`/forms/${childId}/preview`)
                                      }
                                      className="flex-1 min-w-[60px] px-2 py-1.5 text-[10px] sm:text-xs font-medium rounded-lg bg-gradient-to-r from-primary-500 to-primary-600 text-white hover:from-primary-600 hover:to-primary-700 transition-all duration-200 shadow-sm hover:shadow-md flex items-center justify-center gap-1"
                                      title="View form"
                                    >
                                      <Eye className="w-3 h-3" />
                                      View
                                    </button>
                                    <button
                                      onClick={() =>
                                        navigate(`/forms/${childId}/edit`)
                                      }
                                      className="p-1.5 rounded-lg border border-primary-200 text-primary-600 hover:bg-primary-50 transition-colors flex items-center justify-center"
                                      title="Edit form"
                                    >
                                      <Edit2 className="w-3.5 h-3.5" />
                                    </button>
                                  </>
                                )}
                                <button
                                  onClick={() =>
                                    navigate(`/forms/${childId}/analytics`)
                                  }
                                  className={`${isOwner ? 'p-1.5 border border-primary-200 text-primary-600' : 'flex-1 px-2 py-1.5 text-xs font-medium rounded-lg bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm flex items-center justify-center gap-1'} transition-all flex items-center justify-center rounded-lg`}
                                  title="Analytics"
                                >
                                  <BarChart3 className="w-3.5 h-3.5" />
                                  {!isOwner && <span className="ml-1">Analytics</span>}
                                </button>
                                <button
                                  onClick={() =>
                                    navigate(`/forms/${childId}/responses`)
                                  }
                                  className={`${isOwner ? 'p-1.5 border border-primary-200 text-primary-600' : 'p-1.5 border border-indigo-200 text-indigo-600 hover:bg-indigo-50'} transition-all rounded-lg flex items-center justify-center`}
                                  title="Responses"
                                >
                                  <List className="w-3.5 h-3.5" />
                                </button>
                                {(isOwner) && (
                                  <button
                                    onClick={() =>
                                      handleDelete(childId, child.title || "")
                                    }
                                    className="p-1.5 rounded-lg border border-red-200 text-red-600 hover:bg-red-50 transition-colors flex items-center justify-center"
                                    title="Delete"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                )}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      <AnswerTemplateImport
        isOpen={isAnswerTemplateOpen}
        onClose={() => setIsAnswerTemplateOpen(false)}
        onSuccess={() => {
          refetchForms();
          refetchResponses();
        }}
      />

      {isPreviewOpen && previewFormData && (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-900 rounded-lg shadow-xl max-w-4xl w-full max-h-[95vh] overflow-y-auto">
            <div className="sticky top-0 bg-white dark:bg-gray-900 border-b border-gray-200 dark:border-gray-700 p-6 flex items-center justify-between">
              <div>
                <h2 className="text-xl font-bold text-primary-800 dark:text-primary-100">
                  Edit Imported Form
                </h2>
                <p className="text-sm text-primary-600 dark:text-primary-400">
                  Modify form details and then save
                </p>
              </div>
              <button
                onClick={handleCancelImport}
                className="p-1 hover:bg-gray-100 dark:hover:bg-gray-800 rounded-lg transition-colors"
              >
                <X className="w-6 h-6 text-gray-500" />
              </button>
            </div>

            <div className="p-6 space-y-6">
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-primary-700 dark:text-primary-300 mb-2">
                    Form Title
                  </label>
                  <input
                    type="text"
                    value={previewFormData.title || ""}
                    onChange={(e) =>
                      setPreviewFormData({
                        ...previewFormData,
                        title: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 focus:outline-none focus:border-primary-500"
                    placeholder="Enter form title"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-primary-700 dark:text-primary-300 mb-2">
                    Description
                  </label>
                  <textarea
                    value={previewFormData.description || ""}
                    onChange={(e) =>
                      setPreviewFormData({
                        ...previewFormData,
                        description: e.target.value,
                      })
                    }
                    rows={3}
                    className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 focus:outline-none focus:border-primary-500"
                    placeholder="Enter form description (optional)"
                  />
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-primary-700 dark:text-primary-300 mb-2">
                      Sections
                    </label>
                    <p className="text-gray-800 dark:text-gray-200 bg-gray-50 dark:bg-gray-800 p-3 rounded-lg font-medium">
                      {previewFormData.sections?.length || 0} section(s)
                    </p>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-primary-700 dark:text-primary-300 mb-2">
                      Total Questions
                    </label>
                    <p className="text-gray-800 dark:text-gray-200 bg-gray-50 dark:bg-gray-800 p-3 rounded-lg font-medium">
                      {previewFormData.sections?.reduce(
                        (sum, s) => sum + (s.questions?.length || 0),
                        0
                      ) || 0}{" "}
                      question(s)
                    </p>
                  </div>
                </div>

                {previewFormData.sections &&
                  previewFormData.sections.length > 0 && (
                    <div>
                      <label className="block text-sm font-medium text-primary-700 dark:text-primary-300 mb-3">
                        Sections & Questions
                      </label>
                      <div className="space-y-3 max-h-96 overflow-y-auto border border-gray-200 dark:border-gray-700 rounded-lg p-3 bg-gray-50 dark:bg-gray-800">
                        {previewFormData.sections.map((section, idx) => (
                          <div
                            key={idx}
                            className="bg-white dark:bg-gray-900 p-4 rounded-lg border border-gray-200 dark:border-gray-700"
                          >
                            <div className="mb-3">
                              <label className="text-xs font-medium text-primary-600 dark:text-primary-400 block mb-1">
                                Section {idx + 1} Title
                              </label>
                              <input
                                type="text"
                                value={section.title || ""}
                                onChange={(e) => {
                                  const updatedSections = [
                                    ...(previewFormData.sections || []),
                                  ];
                                  updatedSections[idx] = {
                                    ...updatedSections[idx],
                                    title: e.target.value,
                                  };
                                  setPreviewFormData({
                                    ...previewFormData,
                                    sections: updatedSections,
                                  });
                                }}
                                className="w-full px-2 py-1 text-sm border border-gray-300 dark:border-gray-600 rounded bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 focus:outline-none focus:border-primary-500"
                              />
                            </div>
                            <div className="text-sm text-gray-600 dark:text-gray-400 space-y-1">
                              <p className="font-medium">
                                Questions ({section.questions?.length || 0}):
                              </p>
                              {section.questions &&
                              section.questions.length > 0 ? (
                                <ul className="space-y-1 ml-2">
                                  {section.questions.map((q, qIdx) => (
                                    <li
                                      key={qIdx}
                                      className="text-xs text-gray-600 dark:text-gray-400 flex items-start"
                                    >
                                      <span className="mr-2">•</span>
                                      <span className="break-words">
                                        {q.text || `Question ${qIdx + 1}`}
                                      </span>
                                    </li>
                                  ))}
                                </ul>
                              ) : (
                                <p className="text-xs text-gray-500 ml-2">
                                  No questions
                                </p>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
              </div>

              <div className="flex gap-3 pt-4 border-t border-gray-200 dark:border-gray-700">
                <button
                  onClick={handleCancelImport}
                  disabled={isSavingForm}
                  className="flex-1 px-4 py-2 bg-gray-200 dark:bg-gray-700 text-gray-800 dark:text-white rounded-lg hover:bg-gray-300 dark:hover:bg-gray-600 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  Cancel
                </button>
                <button
                  onClick={handleConfirmImport}
                  disabled={isSavingForm || !previewFormData.title?.trim()}
                  className="flex-1 px-4 py-2 bg-primary-600 text-white rounded-lg hover:bg-primary-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                >
                  {isSavingForm ? (
                    <>
                      <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                      Saving...
                    </>
                  ) : (
                    <>
                      <Save className="w-4 h-4" />
                      Save Form
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      <EmailInviteModal
        isOpen={emailInviteModal.open}
        onClose={() =>
          setEmailInviteModal((prev) => ({ ...prev, open: false }))
        }
        formId={emailInviteModal.formId || ""}
        formTitle={emailInviteModal.formTitle}
      />
      <WhatsAppInviteModal
        isOpen={whatsappInviteModal.open}
        onClose={() =>
          setWhatsappInviteModal((prev) => ({ ...prev, open: false }))
        }
        formId={whatsappInviteModal.formId || ""}
        formTitle={whatsappInviteModal.formTitle}
      />
      <SMSInviteModal
        isOpen={smsInviteModal.open}
        onClose={() =>
          setSmsInviteModal((prev) => ({ ...prev, open: false }))
        }
        formId={smsInviteModal.formId || ""}
        formTitle={smsInviteModal.formTitle}
      />
      <ShareAnalyticsModal
        isOpen={shareAnalyticsModal.open}
        onClose={() =>
          setShareAnalyticsModal((prev) => ({ ...prev, open: false }))
        }
        formId={shareAnalyticsModal.formId || ""}
        formTitle={shareAnalyticsModal.formTitle}
      />
    </div>
  );
}