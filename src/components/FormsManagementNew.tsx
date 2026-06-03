import React, { useState, useEffect, useRef, ChangeEvent } from "react";
import { useNavigate } from "react-router-dom";
import {
  PlusCircle,
  Search,
  Edit2,
  Trash2,
  Eye,
  FileText,
  MessageSquarePlus,
  Copy,
  BarChart3,
  TrendingUp,
  Users,
  Calendar,
  MoreVertical,
  List,
  Link as LinkIcon,
  ExternalLink,
  X,
  Check,
  Upload,
  Download,
  Database,
  Send,
  Share2,
} from "lucide-react";
import { useForms, useMutation } from "../hooks/useApi";
import { apiClient } from "../api/client";
import { useNotification } from "../context/NotificationContext";
import { useAuth } from "../context/AuthContext";
import {
  exportFormStructureToExcel,
  downloadFormImportTemplate,
  parseFormWorkbook,
  loadSampleFormData,
} from "../utils/exportUtils";
import ImportPreviewModal from "./ImportPreviewModal";
import ShareAnalyticsModal from "./analytics/ShareAnalyticsModal";
import AutoSendModal from "./forms/AutoSendModal";
import type {
  Question as FormSchema,
  Section as FormSection,
} from "../types/forms";
import type { Question as FormQuestion } from "../types";

interface Form {
  _id: string;
  id?: string;
  title: string;
  description: string;
  isVisible: boolean;
  isActive: boolean;
  sections: any[];
  questions: any[];
  createdAt: string;
  createdBy: any;
  responseCount?: number;
  childForms?: Array<{ formId: string; formTitle: string; order: number }>;
  parentFormId?: string;
  parentFormTitle?: string;
}

export default function FormsManagementNew() {
  const navigate = useNavigate();
  const { user, tenant } = useAuth();
  const isInspector = user?.role === "inspector";
  const isSubAdmin = user?.role === "subadmin";
  const canManage = user?.role === "admin" || user?.role === "superadmin" || user?.role === "subadmin";
  const canEdit = canManage && !isInspector;
  const [searchTerm, setSearchTerm] = useState("");
  const [openDropdownId, setOpenDropdownId] = useState<string | null>(null);
  const [showChildFormsModal, setShowChildFormsModal] = useState(false);
  const [selectedParentForm, setSelectedParentForm] = useState<Form | null>(
    null,
  );
  const [selectedChildFormIds, setSelectedChildFormIds] = useState<string[]>(
    [],
  );
  const dropdownRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const importFileRef = useRef<File | null>(null);
  const { data: formsData, loading, error, execute: refetchForms } = useForms();
  const { showSuccess, showError, showConfirm } = useNotification();
  const [isImporting, setIsImporting] = useState(false);
  const [exportingFormId, setExportingFormId] = useState<string | null>(null);
  const [showImportPreview, setShowImportPreview] = useState(false);
  const [importPreviewData, setImportPreviewData] = useState<any>(null);
  const [importParameters, setImportParameters] = useState<any[]>([]);
  const [isConfirmingImport, setIsConfirmingImport] = useState(false);
  const [showShareModal, setShowShareModal] = useState(false);
  const [selectedFormForShare, setSelectedFormForShare] = useState<Form | null>(null);
  const [showAutoSendModal, setShowAutoSendModal] = useState(false);
  const [selectedFormForAutoSend, setSelectedFormForAutoSend] = useState<Form | null>(null);

  const deleteMutation = useMutation((id: string) => apiClient.deleteForm(id), {
    onSuccess: () => {
      refetchForms();
      showSuccess("Form deleted successfully", "Success");
    },
    onError: (error: any) => {
      showError(error.message || "Failed to delete form", "Error");
    },
  });

  const duplicateMutation = useMutation(
    (id: string) => apiClient.duplicateForm(id),
    {
      onSuccess: () => {
        refetchForms();
        showSuccess("Form duplicated successfully", "Success");
      },
      onError: (error: any) => {
        showError(error.message || "Failed to duplicate form", "Error");
      },
    },
  );

  const visibilityMutation = useMutation(
    ({ id, isVisible }: { id: string; isVisible: boolean }) =>
      apiClient.updateFormVisibility(id, isVisible),
    {
      onSuccess: (data, variables) => {
        refetchForms();
        showSuccess(
          `Form is now ${variables.isVisible ? "public" : "private"}`,
          "Visibility Updated",
        );
      },
      onError: (error: any) => {
        showError(error.message || "Failed to update visibility", "Error");
      },
    },
  );

  const activeMutation = useMutation(
    ({ id, isActive }: { id: string; isActive: boolean }) =>
      apiClient.updateFormActiveStatus(id, isActive),
    {
      onSuccess: (data, variables) => {
        refetchForms();
        showSuccess(
          `Form is now ${variables.isActive ? "active" : "inactive"}`,
          "Status Updated",
        );
      },
      onError: (error: any) => {
        showError(error.message || "Failed to update active status", "Error");
      },
    },
  );

  const linkChildFormMutation = useMutation(
    ({
      parentFormId,
      childFormId,
    }: {
      parentFormId: string;
      childFormId: string;
    }) => apiClient.linkChildForm(parentFormId, childFormId),
    {
      onSuccess: () => {
        refetchForms();
        showSuccess("Child form linked successfully", "Success");
      },
      onError: (error: any) => {
        showError(error.message || "Failed to link child form", "Error");
      },
    },
  );

  const unlinkChildFormMutation = useMutation(
    ({
      parentFormId,
      childFormId,
    }: {
      parentFormId: string;
      childFormId: string;
    }) => apiClient.unlinkChildForm(parentFormId, childFormId),
    {
      onSuccess: () => {
        refetchForms();
        showSuccess("Child form unlinked successfully", "Success");
      },
      onError: (error: any) => {
        showError(error.message || "Failed to unlink child form", "Error");
      },
    },
  );

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(event.target as Node)
      ) {
        setOpenDropdownId(null);
      }
    };

    if (openDropdownId) {
      document.addEventListener("mousedown", handleClickOutside);
    }

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [openDropdownId]);

  const forms = formsData?.forms || [];

  const searchValue = searchTerm.toLowerCase();
  const activeForms = React.useMemo(
    () => forms.filter((form: Form) => form.isActive),
    [forms],
  );

  // Separate parent/standalone forms and child forms
  const parentAndStandaloneForms = React.useMemo(
    () => activeForms.filter((form: Form) => !form.parentFormId),
    [activeForms],
  );

  // Filter forms based on search (parent/standalone only, children will be shown under parents)
  const filteredForms = React.useMemo(
    () =>
      parentAndStandaloneForms.filter(
        (form: Form) =>
          form.title.toLowerCase().includes(searchValue) ||
          form.description.toLowerCase().includes(searchValue),
      ),
    [parentAndStandaloneForms, searchValue],
  );

  const hasAnyForms = forms.length > 0;
  const hasActiveForms = activeForms.length > 0;

  // Get child forms for a parent - memoized to avoid recalculating on every render
  const childFormsMap = React.useMemo(() => {
    const map = new Map<string, Form[]>();
    activeForms.forEach((form: Form) => {
      if (form.parentFormId) {
        if (!map.has(form.parentFormId)) {
          map.set(form.parentFormId, []);
        }
        map.get(form.parentFormId)!.push(form);
      }
    });
    return map;
  }, [activeForms]);

  const getChildFormsForParent = (parentFormId: string) => {
    return childFormsMap.get(parentFormId) || [];
  };

  // Render a single form card (used for both parent and child forms)
  const renderFormCard = (form: Form, isChild: boolean = false) => (
    <div
      key={form._id}
      className={`rounded-lg border p-4 hover:shadow-sm transition-shadow bg-white dark:bg-gray-900 border-neutral-200 dark:border-gray-700 ${
        isChild
          ? "ml-8 border-l-4 border-l-blue-400 bg-blue-50/30 dark:border-l-blue-500 dark:bg-blue-900/40"
          : ""
      }`}
    >
      <div className="flex items-start justify-between mb-4">
        <div className="flex-1">
          <div className="flex items-center gap-2 mb-1">
            {isChild && (
              <LinkIcon className="w-4 h-4 text-blue-600 flex-shrink-0" />
            )}
            <h3
              className={`font-medium line-clamp-2 ${isChild ? "text-sm" : ""}`}
            >
              {form.title}
            </h3>
            {/* Parent Form Indicator */}
            {!isChild && form.childForms && form.childForms.length > 0 && (
              <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-purple-100 text-purple-800 border border-purple-200">
                <LinkIcon className="w-3 h-3 mr-1" />
                Parent ({form.childForms.length} child
                {form.childForms.length !== 1 ? "s" : ""})
              </span>
            )}
            {/* Child Form Indicator */}
            {isChild && (
              <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800 border border-blue-200">
                Child Form
              </span>
            )}
          </div>
          <p
            className={`line-clamp-2 ${
              isChild ? "text-xs text-primary-600" : "text-sm"
            }`}
          >
            {form.description}
          </p>
        </div>
        <div
          className="relative ml-2"
          ref={openDropdownId === form._id ? dropdownRef : null}
        >
          <button
            onClick={() => handleToggleDropdown(form._id)}
            className="p-1 hover:bg-neutral-100 rounded"
          >
            <MoreVertical className="w-4 h-4 text-primary-400" />
          </button>

          {/* Dropdown Menu */}
          {openDropdownId === form._id && (
            <div className="absolute right-0 mt-1 w-48 bg-white dark:bg-gray-900 rounded-lg shadow-lg border border-neutral-200 dark:border-gray-700 py-1 z-10">
              {!isChild && (
                <button
                  onClick={() => handleOpenChildFormsModal(form)}
                  className="w-full text-left px-4 py-2 text-sm text-primary-700 hover:bg-primary-50 flex items-center"
                >
                  <LinkIcon className="w-4 h-4 mr-2" />
                  Manage Child Forms
                </button>
              )}
              <button
                onClick={() => {
                  handleCopyFormLink(form.id || form._id, form.title);
                  setOpenDropdownId(null);
                }}
                className="w-full text-left px-4 py-2 text-sm text-primary-700 hover:bg-primary-50 flex items-center"
              >
                <Copy className="w-4 h-4 mr-2" />
                Copy Customer Link
              </button>
              <button
                onClick={() => {
                  handleOpenFormLink(form.id || form._id);
                  setOpenDropdownId(null);
                }}
                className="w-full text-left px-4 py-2 text-sm text-primary-700 hover:bg-primary-50 flex items-center"
              >
                <ExternalLink className="w-4 h-4 mr-2" />
                Open Customer Link
              </button>
              <button
                onClick={() => {
                  setSelectedFormForShare(form);
                  setShowShareModal(true);
                  setOpenDropdownId(null);
                }}
                className="w-full text-left px-4 py-2 text-sm text-primary-700 hover:bg-primary-50 flex items-center"
              >
                <Share2 className="w-4 h-4 mr-2" />
                Share via WhatsApp/Email
              </button>
              <button
                onClick={() => {
                  setSelectedFormForAutoSend(form);
                  setShowAutoSendModal(true);
                  setOpenDropdownId(null);
                }}
                className="w-full text-left px-4 py-2 text-sm text-primary-700 hover:bg-primary-50 flex items-center"
              >
                <Send className="w-4 h-4 mr-2" />
                Email Automation
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Form Stats */}
      <div
        className={`flex items-center justify-between text-primary-500 mb-4 ${
          isChild ? "text-xs" : "text-xs"
        }`}
      >
        <div className="flex items-center">
          <Users className="w-3 h-3 mr-1" />
          {form.responseCount || 0} responses
        </div>
        <div className="flex items-center">
          <Calendar className="w-3 h-3 mr-1" />
          {new Date(form.createdAt).toLocaleDateString()}
        </div>
      </div>

      {/* Active Status */}
      <div className="flex items-center justify-between mb-4">
        <span
          className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${
            form.isActive
              ? "bg-green-100 text-green-800"
              : "bg-gray-100 text-gray-800"
          }`}
        >
          {form.isActive ? "Active" : "Inactive"}
        </span>
        {canEdit ? (
          <button
            onClick={() => handleToggleActive(form._id, form.isActive)}
            className="text-xs text-primary-600 hover:text-primary-800"
            disabled={activeMutation.loading}
          >
            {activeMutation.loading ? "..." : "Toggle"}
          </button>
        ) : (
          <span className="text-xs text-gray-400 italic">View only</span>
        )}
      </div>

      {/* Actions */}
      <div className="flex items-center justify-between">
        <div className="flex space-x-2">
          <button
            onClick={() => handlePreviewForm(form.id || form._id)}
            className={`px-3 py-2 font-medium text-white bg-primary-600 rounded-lg transition-colors hover:bg-primary-700 ${
              isChild ? "text-xs" : "text-sm"
            }`}
          >
            View
          </button>
          {canEdit && (
            <button
              onClick={() => handleEditForm(form._id)}
              className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
              title="Edit form"
            >
              <Edit2 className={isChild ? "w-3 h-3" : "w-4 h-4"} />
            </button>
          )}
          <button
            onClick={() => handleViewResponses(form._id)}
            className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
            title="View responses"
          >
            <List className={isChild ? "w-3 h-3" : "w-4 h-4"} />
          </button>
        </div>
        <div className="flex space-x-2">
          <button
            onClick={() => handleExportForm(form)}
            className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
            title="Export form structure"
            disabled={exportingFormId === form._id}
          >
            {exportingFormId === form._id ? (
              <div className="w-4 h-4 border-2 border-primary-200 border-t-primary-600 rounded-full animate-spin" />
            ) : (
              <Download className={isChild ? "w-3 h-3" : "w-4 h-4"} />
            )}
          </button>
          {canEdit && (
            <button
              onClick={() => handleDuplicate(form._id)}
              className="p-2 text-primary-600 hover:text-primary-800 hover:bg-primary-50 rounded-lg transition-colors"
              title="Duplicate form"
              disabled={duplicateMutation.loading}
            >
              <Copy className={isChild ? "w-3 h-3" : "w-4 h-4"} />
            </button>
          )}
          {canEdit && (
            <button
              onClick={() => handleDelete(form._id, form.title)}
              className="p-2 text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-colors"
              title="Delete form"
              disabled={deleteMutation.loading}
            >
              <Trash2 className={isChild ? "w-3 h-3" : "w-4 h-4"} />
            </button>
          )}
        </div>
      </div>
    </div>
  );

  const handleDelete = async (id: string, title: string) => {
    showConfirm(
      `Are you sure you want to delete "${title}"? This action cannot be undone.`,
      async () => {
        await deleteMutation.mutate(id);
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
    currentVisibility: boolean,
  ) => {
    await visibilityMutation.mutate({
      id,
      isVisible: !currentVisibility,
    });
  };

  const handleToggleActive = async (id: string, currentActive: boolean) => {
    await activeMutation.mutate({
      id,
      isActive: !currentActive,
    });
  };

  const handleCreateForm = () => {
    navigate("/forms/create");
  };

  const handleExportTemplate = async () => {
    await downloadFormImportTemplate();
  };

  const handleLoadSampleData = async () => {
    try {
      setIsImporting(true);
      const sampleFormData = await loadSampleFormData();

      // Navigate to form creation with sample data pre-loaded
      navigate("/forms/create", {
        state: {
          mode: "create",
          formData: {
            title:
              sampleFormData.title || "Sample Customer Service Feedback Form",
            description:
              sampleFormData.description || "Demo form with sub-parameters",
            sections: sampleFormData.sections,
            followUpQuestions: sampleFormData.followUpQuestions || [],
            isVisible: true,
            locationEnabled: true,
          },
        },
      });
    } catch (error) {
      console.error("Error loading sample data:", error);
      showError("Failed to load sample data", "Error");
    } finally {
      setIsImporting(false);
    }
  };

  const handleExportForm = (form: Form) => {
    const normalized: FormSchema = {
      id: form.id || form._id,
      title: form.title,
      description: form.description,
      sections: (form.sections || []) as FormSection[],
      followUpQuestions: [],
    } as FormSchema;
    setExportingFormId(form._id);
    try {
      exportFormStructureToExcel(normalized as unknown as FormQuestion);
      showSuccess(`Exported form "${form.title}"`, "Export Complete");
    } catch (error: any) {
      showError(error?.message || "Failed to export form", "Export Failed");
    } finally {
      setExportingFormId(null);
    }
  };

  const handleFileInputChange = async (
    event: ChangeEvent<HTMLInputElement>,
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

      // Save file for later use in confirmation
      importFileRef.current = file;

      // Prepare preview data
      const previewSections = (parsed.sections || []).map((section: any) => ({
        title: section.title,
        description: section.description,
        questions: section.questions.map((q: any) => ({
          text: q.text,
          type: q.type,
          required: q.required,
          subParam1: q.subParam1,
          subParam2: q.subParam2,
          options: q.options,
          hasFollowUps: q.followUpQuestions && q.followUpQuestions.length > 0,
        })),
      }));

      setImportPreviewData({
        title: parsed.title,
        description: parsed.description,
        sections: previewSections,
      });

      setImportParameters(parsed.parametersToCreate || []);
      setShowImportPreview(true);
    } catch (error: any) {
      showError(error?.message || "Failed to parse form", "Parse Failed");
      // Clear file input on error
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } finally {
      setIsImporting(false);
    }
  };

  const handleConfirmImport = async () => {
    if (!importPreviewData || !importFileRef.current) return;

    setIsConfirmingImport(true);

    try {
      // Parse the file again to get the full data
      const file = importFileRef.current;
      const parsed = await parseFormWorkbook(file);

      const sectionBranching: any[] = [];
      parsed.sections?.forEach((section: any) => {
        section.questions?.forEach((question: any) => {
          if (question.branchingRules && question.branchingRules.length > 0) {
            question.branchingRules.forEach((rule: any) => {
              sectionBranching.push({
                questionId: question.id,
                sectionId: section.id,
                optionLabel: rule.optionLabel,
                targetSectionId: rule.targetSectionId,
                isOtherOption: false,
              });
            });
          }
        });
      });

      const formPayload = {
        ...parsed,
        sections: parsed.sections || [],
        isVisible: true,
        followUpQuestions: parsed.followUpQuestions || [],
        sectionBranching: sectionBranching,
      } as any;

      const created = await apiClient.createForm(formPayload);

      // Create parameters if any are specified in the Excel
      if (
        parsed.parametersToCreate &&
        parsed.parametersToCreate.length > 0 &&
        created?.form?._id
      ) {
        const parameterPromises = parsed.parametersToCreate.map((param) =>
          apiClient.createParameter({
            ...param,
            formId: created.form._id,
          }),
        );
        await Promise.all(parameterPromises);
        showSuccess(
          `Form imported with ${parsed.parametersToCreate.length} parameter(s)`,
          "Import Complete",
        );
      } else {
        showSuccess("Form imported successfully", "Import Complete");
      }

      if (created?.form?._id) {
        navigate(`/forms/${created.form._id}/edit`);
      }
      refetchForms();
      setShowImportPreview(false);
      importFileRef.current = null;
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } catch (error: any) {
      showError(error?.message || "Failed to import form", "Import Failed");
    } finally {
      setIsConfirmingImport(false);
    }
  };

  const handleImportClick = () => {
    if (isImporting) {
      return;
    }
    fileInputRef.current?.click();
  };

  const handleEditForm = (id: string) => {
    navigate(`/forms/${id}/edit`);
  };

  const handleViewResponses = (id: string) => {
    navigate(`/forms/${id}/responses`);
  };

  const handlePreviewForm = (id: string) => {
    navigate(`/forms/${id}/preview`);
  };

  const handleViewAnalytics = (id: string) => {
    navigate(`/forms/${id}/analytics`);
  };

  const getCustomerFormUrl = (formId: string) => {
    const tenantSlug = user?.tenant?.slug || "default";
    // For production, use your actual customer frontend URL
    const baseUrl = window.location.origin.includes("localhost")
      ? "http://localhost:5174"
      : "https://3wheelertvs.focusengineeringapp.com/";
    // : "https://formsuser.focusengineeringapp.com";
    return `${baseUrl}${tenantSlug}/forms/${formId}`;
  };

  const handleCopyFormLink = (formId: string, formTitle: string) => {
    const url = getCustomerFormUrl(formId);
    navigator.clipboard
      .writeText(url)
      .then(() => {
        showSuccess(`Link copied for "${formTitle}"`, "Link Copied!");
      })
      .catch(() => {
        showError("Failed to copy link", "Error");
      });
  };

  const handleOpenFormLink = (formId: string) => {
    const url = getCustomerFormUrl(formId);
    window.open(url, "_blank");
  };

  const handleToggleDropdown = (formId: string) => {
    setOpenDropdownId(openDropdownId === formId ? null : formId);
  };

  const handleOpenChildFormsModal = (form: Form) => {
    setSelectedParentForm(form);
    // Get currently linked child form IDs
    const linkedChildIds = form.childForms?.map((cf) => cf.formId) || [];
    setSelectedChildFormIds(linkedChildIds);
    setShowChildFormsModal(true);
    setOpenDropdownId(null);
  };

  const handleCloseChildFormsModal = () => {
    setShowChildFormsModal(false);
    setSelectedParentForm(null);
    setSelectedChildFormIds([]);
  };

  const handleToggleChildForm = (childFormId: string) => {
    setSelectedChildFormIds((prev) => {
      if (prev.includes(childFormId)) {
        return prev.filter((id) => id !== childFormId);
      } else {
        return [...prev, childFormId];
      }
    });
  };

  const handleSaveChildForms = async () => {
    if (!selectedParentForm) return;

    const parentFormId = selectedParentForm._id;
    const currentLinkedIds =
      selectedParentForm.childForms?.map((cf) => cf.formId) || [];

    // Find forms to link (newly selected)
    const toLink = selectedChildFormIds.filter(
      (id) => !currentLinkedIds.includes(id),
    );

    // Find forms to unlink (deselected)
    const toUnlink = currentLinkedIds.filter(
      (id) => !selectedChildFormIds.includes(id),
    );

    try {
      // Link new child forms
      for (const childFormId of toLink) {
        await linkChildFormMutation.mutate({ parentFormId, childFormId });
      }

      // Unlink removed child forms
      for (const childFormId of toUnlink) {
        await unlinkChildFormMutation.mutate({ parentFormId, childFormId });
      }

      handleCloseChildFormsModal();
    } catch (error) {
      // Error handling is done in mutation callbacks
    }
  };

  // Get available forms for linking (exclude the parent itself and forms that are children of OTHER parents)
  const getAvailableChildForms = () => {
    if (!selectedParentForm) return [];

    return forms.filter(
      (form: Form) =>
        form._id !== selectedParentForm._id && // Not the parent itself
        (!form.parentFormId || form.parentFormId === selectedParentForm._id), // Not a child of another parent (but allow children of THIS parent)
    );
  };

  if (loading) {
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
    <>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        className="hidden"
        onChange={handleFileInputChange}
      />
      <div className="p-6">
        <div className="mb-4">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
            <div>
              <h1 className="font-bold">Service Request Management</h1>
              <p>Create, edit, and manage service request forms</p>
            </div>
            <div className="flex flex-col sm:flex-row sm:items-center gap-2">
              {canManage && !isInspector && (
                <>
                  <button
                    onClick={handleExportTemplate}
                    className="btn-secondary flex items-center justify-center"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download Import Template
                  </button>
                  <button
                    onClick={handleLoadSampleData}
                    className="btn-secondary flex items-center justify-center"
                    disabled={isImporting}
                  >
                    <Database className="w-4 h-4 mr-2" />
                    {isImporting ? "Loading..." : "Load Sample Data"}
                  </button>
                  <button
                    onClick={handleImportClick}
                    className="btn-secondary flex items-center justify-center"
                    disabled={isImporting}
                  >
                    <Upload className="w-4 h-4 mr-2" />
                    {isImporting ? "Importing..." : "Import Form (Excel)"}
                  </button>
                  <button
                    onClick={handleCreateForm}
                    className="btn-primary flex items-center justify-center"
                  >
                    <PlusCircle className="w-3 h-3 mr-1" />
                    Create New Service Form
                  </button>
                </>
              )}
            </div>
          </div>
        </div>

        {/* Search */}
        <div className="mb-6">
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
        </div>

        {/* Forms List */}
        {filteredForms.length === 0 ? (
          <div className="text-center py-12 bg-white dark:bg-gray-900 rounded-lg border border-neutral-200 dark:border-gray-700">
            <FileText className="w-12 h-12 text-primary-300 mx-auto mb-4" />
            <h3 className="text-lg font-medium text-primary-600 mb-2">
              {(() => {
                if (!hasAnyForms) {
                  return "No service forms created yet";
                }
                if (!hasActiveForms) {
                  return "No active service forms";
                }
                return "No service forms found";
              })()}
            </h3>
            <p className="text-primary-500 mb-6">
              {(() => {
                if (!hasAnyForms) {
                  return "Create your first service form to get started";
                }
                if (!hasActiveForms) {
                  return "Activate a form to make it available here";
                }
                return "Try adjusting your search criteria";
              })()}
            </p>
            {!hasAnyForms && canEdit && (
              <button onClick={handleCreateForm} className="btn-primary">
                <PlusCircle className="w-4 h-4 mr-2" />
                Create Your First Form
              </button>
            )}
          </div>
        ) : (
          <div className="space-y-4">
            {filteredForms.map((form: Form) => {
              const childForms = getChildFormsForParent(form._id);
              return (
                <div key={form._id} className="space-y-2">
                  {/* Parent/Standalone Form */}
                  {renderFormCard(form, false)}

                  {/* Child Forms (nested under parent) */}
                  {childForms.length > 0 && (
                    <div className="space-y-2">
                      {childForms.map((childForm: Form) =>
                        renderFormCard(childForm, true),
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* Child Forms Management Modal */}
        {showChildFormsModal && selectedParentForm && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
            <div className="bg-white dark:bg-gray-900 rounded-lg shadow-xl max-w-2xl w-full max-h-[80vh] overflow-hidden flex flex-col">
              {/* Modal Header */}
              <div className="px-6 py-4 border-b border-neutral-200 dark:border-gray-700 flex items-center justify-between">
                <div>
                  <h2 className="text-xl font-semibold text-primary-900">
                    Manage Child Forms
                  </h2>
                  <p className="text-sm text-primary-600 mt-1">
                    Parent: {selectedParentForm.title}
                  </p>
                </div>
                <button
                  onClick={handleCloseChildFormsModal}
                  className="p-1 hover:bg-neutral-100 rounded"
                >
                  <X className="w-5 h-5 text-primary-400" />
                </button>
              </div>

              {/* Modal Body */}
              <div className="px-6 py-4 overflow-y-auto flex-1">
                <p className="text-sm text-primary-600 mb-4">
                  Select forms to link as child forms. Child forms will appear
                  after users complete this parent form.
                </p>

                {getAvailableChildForms().length === 0 ? (
                  <div className="text-center py-8">
                    <FileText className="w-12 h-12 text-primary-300 mx-auto mb-3" />
                    <p className="text-primary-600">
                      No available forms to link as child forms.
                    </p>
                    <p className="text-sm text-primary-500 mt-2">
                      Create more forms or unlink existing child forms from
                      other parents.
                    </p>
                  </div>
                ) : (
                  <div className="space-y-2">
                    {getAvailableChildForms().map((form: Form) => {
                      const isSelected = selectedChildFormIds.includes(
                        form._id,
                      );
                      return (
                        <div
                          key={form._id}
                          className={`border rounded-lg p-4 cursor-pointer transition-all ${
                            isSelected
                              ? "border-primary-500 bg-primary-50"
                              : "border-neutral-200 hover:border-primary-300 hover:bg-primary-50"
                          }`}
                          onClick={() => handleToggleChildForm(form._id)}
                        >
                          <div className="flex items-start justify-between">
                            <div className="flex-1">
                              <div className="flex items-center gap-2">
                                <h3 className="font-medium text-primary-900">
                                  {form.title}
                                </h3>
                                {isSelected && (
                                  <Check className="w-5 h-5 text-primary-600" />
                                )}
                              </div>
                              <p className="text-sm text-primary-600 mt-1 line-clamp-2">
                                {form.description}
                              </p>
                              <div className="flex items-center gap-4 mt-2 text-xs text-primary-500">
                                <span className="flex items-center">
                                  <Users className="w-3 h-3 mr-1" />
                                  {form.responseCount || 0} responses
                                </span>
                                <span className="flex items-center">
                                  <Calendar className="w-3 h-3 mr-1" />
                                  {new Date(
                                    form.createdAt,
                                  ).toLocaleDateString()}
                                </span>
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Modal Footer */}
              <div className="px-6 py-4 border-t border-neutral-200 dark:border-gray-700 flex items-center justify-between">
                <div className="text-sm text-primary-600">
                  {selectedChildFormIds.length} form
                  {selectedChildFormIds.length !== 1 ? "s" : ""} selected
                </div>
                <div className="flex space-x-2">
                  <button
                    onClick={handleCloseChildFormsModal}
                    className="px-4 py-2 text-sm font-medium text-primary-700 hover:bg-neutral-100 rounded-lg transition-colors"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleSaveChildForms}
                    disabled={
                      linkChildFormMutation.loading ||
                      unlinkChildFormMutation.loading
                    }
                    className="px-4 py-2 text-sm font-medium text-white bg-primary-600 rounded-lg hover:bg-primary-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {linkChildFormMutation.loading ||
                    unlinkChildFormMutation.loading
                      ? "Saving..."
                      : "Save Changes"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Import Preview Modal */}
        {showImportPreview && importPreviewData && (
          <ImportPreviewModal
            isOpen={showImportPreview}
            onClose={() => {
              setShowImportPreview(false);
              setImportPreviewData(null);
              setImportParameters([]);
              importFileRef.current = null;
              if (fileInputRef.current) {
                fileInputRef.current.value = "";
              }
            }}
            onConfirm={handleConfirmImport}
            formData={importPreviewData}
            parameters={importParameters}
            isLoading={isConfirmingImport}
          />
        )}

        {showShareModal && selectedFormForShare && (
          <ShareAnalyticsModal
            isOpen={showShareModal}
            onClose={() => {
              setShowShareModal(false);
              setSelectedFormForShare(null);
            }}
            formId={selectedFormForShare.id || selectedFormForShare._id}
            formTitle={selectedFormForShare.title}
          />
        )}

        {showAutoSendModal && selectedFormForAutoSend && (
          <AutoSendModal
            isOpen={showAutoSendModal}
            onClose={() => {
              setShowAutoSendModal(false);
              setSelectedFormForAutoSend(null);
            }}
            formId={selectedFormForAutoSend.id || selectedFormForAutoSend._id}
            formTitle={selectedFormForAutoSend.title}
          />
        )}
      </div>
    </>
  );
}
