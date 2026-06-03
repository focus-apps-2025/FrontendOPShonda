import type { QuestionType } from "./questionTypes";

export interface GridOption {
  rows: string[];
  columns: string[];
}

export interface ShowWhen {
  questionId: string;
  value: string | number;
}

export interface Section {
  id: string;
  title: string;
  description?: string;
  nextSectionId?: string;
  questions: FollowUpQuestion[];
  linkedToOption?: string;
  linkedToQuestionId?: string;
  merging?: string;
}

export interface FollowUpQuestion {
  id: string;
  text: string;
  type: QuestionType;
  required?: boolean;
  options?: string[];
  allowedFileTypes?: string[];
  correctAnswer?: string; // Single correct answer (backward compatibility)
  correctAnswers?: string[]; // Multiple correct answers
  gridOptions?: GridOption;
  min?: number;
  max?: number;
  minLabel?: string;
  maxLabel?: string;
  step?: number;
  showWhen?: ShowWhen;
  parentId?: string;
  imageUrl?: string;
  description?: string;
  suggestion?: string;
  sectionId?: string;
  subParam1?: string;
  subParam2?: string;
  followUpQuestions?: FollowUpQuestion[]; // Support nested follow-ups
  trackResponseRank?: boolean;
  trackResponseRankLabel?: string;
  trackResponseRankType?: string;
  trackResponseQuestion?: boolean;
  trackResponseQuestionType?: string;
  trackResponseQuestionLabel?: string;
  branchingRules?: Array<{
    optionLabel: string;
    targetSectionId: string;
    isOtherOption?: boolean;
  }>;
  requireFollowUp?: boolean; // Make follow-up mandatory for certain question types
}

export interface OpsTemplateConfig {
  basicInfoLabels?: {
    deptSection?: string;
    lineZone?: string;
    model?: string;
    processStation?: string;
    formatNo?: string;
    controlNo?: string;
  };
  abnormalityHandlingRoute?: string;
  abnormalityDetailsLabel?: string;
  emsGuidelines?: string[];
  fiveSGuidelines?: string[];
  processInstructions?: string[];
  generalInstructions?: string[];
  troubleTasks?: Array<{
    sno: number;
    trouble: string;
    task: string;
  }>;
  rejectionHandling?: string;
  measuringInstruments?: string[];
  tableHeaders?: Record<string, string>;
  associateSignArea?: {
    title1?: string;
    title2?: string;
  };
  illustrationsQuestion?: string;
}

export interface Question {
  id: string;
  title: string;
  description: string;
  logoUrl?: string;
  imageUrl?: string;
  sections: Section[];
  followUpQuestions: FollowUpQuestion[];
  parentFormId?: string;
  parentFormTitle?: string;
  locationEnabled?: boolean;
  viewType?: "section-wise" | "question-wise";
  followUpConfig?: Record<
    string,
    {
      hasFollowUp: boolean;
      required: boolean;
      linkedSectionId?: string;
      linkedFormId?: string;
    }
  >;
  chassisNumbers?: string[];
  chassisTenantAssignments?: Array<{
    chassisNumber: string;
    assignedTenants: string[];
  }>;
  opsTemplateConfig?: OpsTemplateConfig;
}

export interface Response {
  id: string;
  questionId: string;
  answers: Record<string, any>;
  timestamp: string;
  parentResponseId?: string;
  assignedTo?: string;
  assignedAt?: string;
  verifiedBy?: string;
  verifiedAt?: string;
  status?: "pending" | "verified" | "rejected";
  notes?: string;
  score?: {
    correct: number;
    total: number;
  };
  submissionMetadata?: {
    ipAddress?: string;
    userAgent?: string;
    browser?: string;
    device?: string;
    os?: string;
    location?: {
      country?: string;
      countryCode?: string;
      region?: string;
      city?: string;
      latitude?: number;
      longitude?: number;
      timezone?: string;
      isp?: string;
    };
    capturedLocation?: {
      latitude?: number;
      longitude?: number;
      accuracy?: number;
      source?: "browser" | "ip" | "manual" | "unknown";
      capturedAt?: string;
    };
    submittedAt?: string;
    source?: string;
    formSessionId?: string | null;
  };
  totalTimeSpent?: number;
  timeSpent?: number; // Backend uses this field name
  questionTimings?: Array<{
    questionId: string;
    timeSpent: number;
  }>;
  isDispatched?: boolean;
  dispatchedAt?: string;
}

export interface Profile {
  name: string;
  email: string;
  phone: string;
  joinDate: string;
  avatar?: string;
  userId: string;
  username?: string;
}
