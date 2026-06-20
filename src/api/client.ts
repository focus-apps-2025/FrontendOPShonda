// Automatically detect environment and set API base URL
const API_BASE_URL = (() => {
  const hostname = window.location.hostname;
  const isLocal =
    hostname === "localhost" ||
    hostname === "127.0.0.1" ||
    hostname.startsWith("192.168.") ||
    hostname.startsWith("10.") ||
    hostname.startsWith("172.");

  const baseUrl = isLocal
    ? "http://127.0.0.1:5000/api"
    : "https://backendops.focusengineeringapp.com/api";

  console.log(
    `🔗 API Base URL: ${baseUrl} (Environment: ${isLocal ? "Local" : "Production"
    })`,
  );
  return baseUrl;
})();
interface ApiResponse<T> {
  success: boolean;
  message?: string;
  data?: T;
  errors?: string[];
}
interface ApiResponse<T> {
  success: boolean;
  message?: string;
  data?: T;
  errors?: string[];
}

class ApiError extends Error {
  constructor(
    public status: number,
    public response: any,
    message?: string,
  ) {
    super(message || "API Error");
    this.name = "ApiError";
  }
}

class ApiClient {
  private baseUrl: string;
  private token: string | null = null;

  constructor(baseUrl: string = API_BASE_URL) {
    this.baseUrl = baseUrl;
    this.token = localStorage.getItem("auth_token");
  }

  setToken(token: string) {
    this.token = token;
    localStorage.setItem("auth_token", token);
  }

  clearToken() {
    this.token = null;
    localStorage.removeItem("auth_token");
  }

  public getBaseUrl() {
    return this.baseUrl;
  }

  public async request<T>(
    endpoint: string,
    options: RequestInit = {},
  ): Promise<T> {
    const url = `${this.baseUrl}${endpoint}`;

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      "X-App-Type": "website",
      ...(options.headers as Record<string, string>),
    };

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    } else {
      const guestToken = localStorage.getItem("guest_auth_token");
      if (guestToken) {
        headers.Authorization = `Bearer ${guestToken}`;
      }
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);

    const config: RequestInit = {
      ...options,
      headers,
      signal: controller.signal,
    };

    try {
      const response = await fetch(url, config);
      clearTimeout(timeoutId);
      const data: ApiResponse<T> = await response.json();

      if (!response.ok) {
        let errorMessage = data.message || "Request failed";
        if (data.errors && Array.isArray(data.errors)) {
          const errorDetails = data.errors
            .map((err: any) =>
              typeof err === "string" ? err : `${err.field}: ${err.message}`,
            )
            .join(", ");
          errorMessage = `${errorMessage}: ${errorDetails}`;
        }
        throw new ApiError(response.status, data, errorMessage);
      }

      if (!data.success) {
        let errorMessage = data.message || "Request failed";
        if (data.errors && Array.isArray(data.errors)) {
          const errorDetails = data.errors
            .map((err: any) =>
              typeof err === "string" ? err : `${err.field}: ${err.message}`,
            )
            .join(", ");
          errorMessage = `${errorMessage}: ${errorDetails}`;
        }
        throw new ApiError(response.status, data, errorMessage);
      }

      return data.data as T;
    } catch (error) {
      clearTimeout(timeoutId);
      if (error instanceof ApiError) {
        throw error;
      }
      if (error instanceof DOMException && error.name === "AbortError") {
        throw new ApiError(
          408,
          null,
          "Request timed out. The server took too long to respond.",
        );
      }
      throw new ApiError(500, null, "Network error or server unavailable");
    }
  }

  // HTTP Helper Methods
  public async get<T>(
    endpoint: string,
    options: RequestInit = {},
  ): Promise<{ success: boolean; data: T; message?: string }> {
    const data = await this.request<T>(endpoint, { ...options, method: "GET" });
    return { success: true, data };
  }

  public async post<T>(
    endpoint: string,
    body: any,
    options: RequestInit = {},
  ): Promise<{ success: boolean; data: T; message?: string }> {
    const data = await this.request<T>(endpoint, {
      ...options,
      method: "POST",
      body: JSON.stringify(body),
    });
    return { success: true, data };
  }

  public async put<T>(
    endpoint: string,
    body: any,
    options: RequestInit = {},
  ): Promise<{ success: boolean; data: T; message?: string }> {
    const data = await this.request<T>(endpoint, {
      ...options,
      method: "PUT",
      body: JSON.stringify(body),
    });
    return { success: true, data };
  }

  public async delete<T>(
    endpoint: string,
    options: RequestInit = {},
  ): Promise<{ success: boolean; data: T; message?: string }> {
    const data = await this.request<T>(endpoint, {
      ...options,
      method: "DELETE",
    });
    return { success: true, data };
  }

  private ensureAbsoluteFileUrl(value: string) {
    if (!value) {
      return "";
    }
    if (
      value.startsWith("http://") ||
      value.startsWith("https://") ||
      value.startsWith("data:")
    ) {
      return value;
    }
    const normalized = value.startsWith("/") ? value : `/${value}`;
    try {
      const url = new URL(this.baseUrl);
      return `${url.origin}${normalized}`;
    } catch {
      return normalized;
    }
  }

  // Authentication
  async login(credentials: {
    email: string;
    password: string;
    tenantSlug?: string;
    location?: any;
  }) {
    const data = await this.request<{
      user: any;
      token: string;
      tenant?: any;
      sessionLogId?: string;
    }>("/auth/login", {
      method: "POST",
      body: JSON.stringify(credentials),
    });

    this.setToken(data.token);
    return data;
  }

  async signup(signupData: {
    name: string;
    slug: string;
    companyName: string;
    adminEmail: string;
    adminPassword: string;
    adminFirstName: string;
    adminLastName: string;
  }) {
    return this.request<{ tenant: any; admin: any }>("/auth/signup", {
      method: "POST",
      body: JSON.stringify(signupData),
    });
  }

  async logout(sessionLogId?: string) {
    try {
      await this.request("/auth/logout", {
        method: "POST",
        body: JSON.stringify({ sessionLogId }),
      });
    } catch (e) {
      console.warn("Logout request failed:", e);
    } finally {
      this.clearToken();
    }
  }

  async getProfile() {
    return this.request<{ user: any }>("/auth/profile");
  }

  async changePassword(passwords: {
    currentPassword: string;
    newPassword: string;
  }) {
    return this.request("/auth/change-password", {
      method: "PUT",
      body: JSON.stringify(passwords),
    });
  }

  // Forgot Password - Step 1: Request OTP via email (triggers SMS to registered mobile)
  async forgotPassword(email: string) {
    return this.request<{
      success: boolean;
      message: string;
      data?: any;
    }>("/auth/forgot-password", {
      method: "POST",
      body: JSON.stringify({ email: email.toLowerCase().trim() }),
    });
  }

  // Forgot Password - Resend OTP
  async resendForgotOtp(email: string) {
    return this.request<{
      success: boolean;
      message: string;
      data?: any;
    }>("/auth/resend-otp", {
      method: "POST",
      body: JSON.stringify({ email: email.toLowerCase().trim() }),
    });
  }

  // Forgot Password - Step 2: Verify OTP
  async verifyForgotOtp(email: string, otp: string) {
    return this.request<{
      success: boolean;
      message: string;
      verificationId?: string;
      data?: { verificationId?: string };
    }>("/auth/verify-forgot-otp", {
      method: "POST",
      body: JSON.stringify({ email: email.toLowerCase().trim(), otp }),
    });
  }

  // Forgot Password - Step 3: Reset password
  async resetPassword(email: string, verificationId: string, newPassword: string) {
    return this.request<{
      success: boolean;
      message: string;
    }>("/auth/reset-password", {
      method: "POST",
      body: JSON.stringify({ email: email.toLowerCase().trim(), verificationId, newPassword }),
    });
  }

  // Users
  async getUsersHierarchy(params?: { role?: string }) {
    return this.request<{ users: any[] }>(
      `/users/hierarchy${params?.role ? `?role=${params.role}` : ""}`,
    );
  }

  async getUserActivityLogs(params?: {
    page?: number;
    limit?: number;
    search?: string;
    startDate?: string;
    endDate?: string;
  }) {
    const query = new URLSearchParams();
    if (params?.page) query.set("page", params.page.toString());
    if (params?.limit) query.set("limit", params.limit.toString());
    if (params?.search) query.set("search", params.search);
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);

    return this.request<{ logs: any[]; pagination: any }>(
      `/users/activity-logs?${query.toString()}`,
    );
  }

  // SuperAdmin: Get all tenants performance
  async getAllTenantsPerformance(params?: {
    startDate?: string;
    endDate?: string;
    search?: string;
    page?: number;
    limit?: number;
  }) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.search) query.set("search", params.search);
    if (params?.page) query.set("page", params.page.toString());
    if (params?.limit) query.set("limit", params.limit.toString());

    return this.request<{
      users: any[];
      pagination: {
        currentPage: number;
        totalPages: number;
        totalUsers: number;
        hasNextPage: boolean;
        hasPrevPage: boolean;
      };
    }>(`/users/all-tenants-performance?${query.toString()}`);
  }

  // SuperAdmin: Get response details for a specific tenant
  async getTenantResponseDetails(
    tenantId: string,
    params?: { startDate?: string; endDate?: string },
  ) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);

    const endpoint = `/analytics/superadmin/tenant/${tenantId}/response-details${query.toString() ? `?${query.toString()}` : ""}`;
    return this.request<{
      totalResponses: number;
      statusBreakdown: { pending: number; verified: number; rejected: number };
      yesNoNA: { yes: number; no: number; na: number };
      formBreakdown: {
        formId: string;
        formTitle: string;
        yes: number;
        no: number;
        na: number;
        responseCount: number;
        avgTimeSpent?: number;
        totalTimeSpent?: number;
      }[];
    }>(endpoint);
  }

  // Users
  async getUsers(params?: {
    role?: string;
    search?: string;
    page?: number;
    limit?: number;
    tenantId?: string;
  }) {
    const query = new URLSearchParams();

    if (params?.role) {
      query.set("role", params.role);
    }

    if (params?.search) {
      query.set("search", params.search);
    }

    if (typeof params?.page === "number") {
      query.set("page", params.page.toString());
    }

    if (typeof params?.limit === "number") {
      query.set("limit", params.limit.toString());
    }

    if (params?.tenantId) {
      query.set("tenantId", params.tenantId);
    }

    const endpoint = `/users${query.toString() ? `?${query.toString()}` : ""}`;

    return this.request<{ users: any[]; pagination: any }>(endpoint);
  }

  async createUser(userData: any) {
    return this.request<{ user: any }>("/users", {
      method: "POST",
      body: JSON.stringify(userData),
    });
  }

  async updateUser(id: string, userData: any) {
    return this.request<{ user: any }>(`/users/${id}`, {
      method: "PUT",
      body: JSON.stringify(userData),
    });
  }

  async deleteUser(id: string) {
    return this.request(`/users/${id}`, {
      method: "DELETE",
    });
  }

  async resetUserPassword(id: string, newPassword: string) {
    return this.request(`/users/${id}/reset-password`, {
      method: "PUT",
      body: JSON.stringify({ newPassword }),
    });
  }

  // Forms
  // In client.ts
  async getForms(params?: { isGlobal?: boolean; search?: string }) {
    const query = new URLSearchParams();
    if (params?.isGlobal !== undefined) {
      query.set("isGlobal", params.isGlobal.toString());
    }
    if (params?.search) {
      query.set("search", params.search);
    }
    const endpoint = `/forms${query.toString() ? `?${query.toString()}` : ""}`;
    const result = await this.request<{ forms: any[] }>(endpoint);

    // DEBUG: Log the response
    console.log("Forms API Response:", result);
    result.forms.forEach((form) => {
      console.log(`Form "${form.title}" responseCount: ${form.responseCount}`);
    });

    return result;
  }

  async getPublicForms(tenantSlug?: string) {
    const endpoint = tenantSlug
      ? `/forms/public/${tenantSlug}`
      : "/forms/public";
    return this.request<{ forms: any[] }>(endpoint);
  }

  async getForm(id: string) {
    return this.request<{ form: any }>(`/forms/${id}`);
  }

  async getFormById(id: string) {
    return this.request<{ form: any }>(`/forms/${id}`);
  }

  async getPublicForm(id: string, tenantSlug?: string) {
    // IMPORTANT: tenantSlug is REQUIRED for public access
    if (!tenantSlug) {
      // You need to get the tenantSlug from somewhere
      // Maybe from the URL params or a default value
      console.error("tenantSlug is required for public form access");
      throw new Error("tenantSlug is required");
    }

    // This matches your backend route exactly: /forms/:id/public/:tenantSlug
    const endpoint = `/forms/${id}/public/${tenantSlug}`;
    const url = `${this.baseUrl}${endpoint}`;

    const response = await fetch(url);
    const data: ApiResponse<{ form: any }> = await response.json();

    if (!response.ok) {
      throw new ApiError(response.status, data, data.message);
    }

    if (!data.success) {
      throw new ApiError(response.status, data, data.message);
    }

    return data.data;
  }
  async submitPublicResponse(
    formId: string,
    data: {
      inviteId: string;
      answers: any;
      location?: any;
      submissionMetadata?: any;
      startedAt?: Date | string;
      completedAt?: Date | string;
      sessionId?: string;
      isSectionSubmit?: boolean;
      sectionIndex?: number;
    },
  ) {
    const url = `${this.baseUrl}/forms/${formId}/public/submit`;

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };

    // Include auth token if available
    if (this.token) {
      headers["Authorization"] = `Bearer ${this.token}`;
    }

    const response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(data),
    });

    const result = await response.json();

    console.log("Submit response status:", response.status);
    console.log("Submit response body:", result);

    if (!response.ok) {
      throw new ApiError(
        response.status,
        result,
        result.message || "Submission failed",
      );
    }

    if (!result.success) {
      throw new ApiError(
        response.status,
        result,
        result.message || "Submission failed",
      );
    }

    // Log the data being returned
    console.log("submitPublicResponse result.data:", result.data);

    return result.data;
  }
  async createForm(formData: any) {
    return this.request<{ form: any }>("/forms", {
      method: "POST",
      body: JSON.stringify(formData),
    });
  }

  async updateForm(id: string, formData: any) {
    return this.request<{ form: any }>(`/forms/${id}`, {
      method: "PUT",
      body: JSON.stringify(formData),
    });
  }

  async deleteForm(id: string) {
    return this.request(`/forms/${id}`, {
      method: "DELETE",
    });
  }

  async updateFormVisibility(id: string, isVisible: boolean) {
    return this.request(`/forms/${id}/visibility`, {
      method: "PATCH",
      body: JSON.stringify({ isVisible }),
    });
  }

  async updateFormLocationEnabled(id: string, locationEnabled: boolean) {
    return this.request(`/forms/${id}/location`, {
      method: "PATCH",
      body: JSON.stringify({ locationEnabled }),
    });
  }

  async updateFormActiveStatus(id: string, isActive: boolean) {
    return this.request(`/forms/${id}/active`, {
      method: "PATCH",
      body: JSON.stringify({ isActive }),
    });
  }

  async updateFormViewType(
    id: string,
    viewType: "section-wise" | "question-wise",
  ) {
    return this.request(`/forms/${id}/view-type`, {
      method: "PATCH",
      body: JSON.stringify({ viewType }),
    });
  }

  async duplicateForm(id: string) {
    return this.request<{ form: any }>(`/forms/${id}/duplicate`, {
      method: "POST",
    });
  }

  async getAutoSendConfig(formId: string) {
    return this.get<any>(`/forms/${formId}/autosend/config`);
  }

  async updateAutoSendConfig(formId: string, config: any) {
    return this.put<any>(`/forms/${formId}/autosend/config`, {
      autoSendConfig: config,
    });
  }

  async getAutoSendHistory(formId: string, page: number = 1) {
    return this.get<any>(`/forms/${formId}/autosend/history?page=${page}`);
  }

  // Child Form Management (Parent-Child Relationships)
  async linkChildForm(parentFormId: string, childFormId: string) {
    return this.request<{ parentForm: any; childForm: any }>(
      `/forms/${parentFormId}/child-forms`,
      {
        method: "POST",
        body: JSON.stringify({ childFormId }),
      },
    );
  }

  async unlinkChildForm(parentFormId: string, childFormId: string) {
    return this.request(`/forms/${parentFormId}/child-forms/${childFormId}`, {
      method: "DELETE",
    });
  }

  async getChildForms(parentFormId: string) {
    return this.request<{
      parentFormId: string;
      parentFormTitle: string;
      childForms: any[];
    }>(`/forms/${parentFormId}/child-forms`);
  }

  async reorderChildForms(parentFormId: string, childFormOrder: string[]) {
    return this.request(`/forms/${parentFormId}/child-forms/reorder`, {
      method: "PUT",
      body: JSON.stringify({ childFormOrder }),
    });
  }

  // Responses
  async getResponses() {
    return this.request<{ responses: any[]; pagination?: any }>(
      "/responses?limit=1000",
    );
  }

  async getFormResponses(formId: string, options?: { analytics?: boolean }) {
    const query = options?.analytics ? "?analytics=true" : "";
    return this.request<{ responses: any[] }>(
      `/responses/form/${formId}${query}`,
    );
  }

  async getResponse(id: string) {
    return this.request<{ response: any }>(`/responses/${id}`);
  }

  async createResponse(responseData: any) {
    const tenantSlug = responseData.tenantSlug;
    const formId = responseData.questionId || responseData.formId;

    let url = `${this.baseUrl}/responses`;
    if (tenantSlug && formId) {
      url = `${this.baseUrl}/responses/${tenantSlug}/forms/${formId}/responses`;
    }

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };
    if (this.token) headers.Authorization = `Bearer ${this.token}`;

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);

    try {
      const res = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify(responseData),
        signal: controller.signal,
      });

      clearTimeout(timeoutId);
      const json = await res.json();

      if (!res.ok) {
        throw new ApiError(
          res.status,
          json,
          json.message || "Failed to create response",
        );
      }

      return json;
    } catch (err: any) {
      clearTimeout(timeoutId);
      if (err.name === "AbortError") {
        throw new ApiError(
          408,
          null,
          "Request timed out. The server took too long to respond.",
        );
      }
      if (err instanceof ApiError) throw err;
      throw new ApiError(500, null, err.message || "Failed to create response");
    }
  }

  async batchImportResponses(batchData: any) {
    return this.request<any>("/responses/batch/import", {
      method: "POST",
      body: JSON.stringify(batchData),
    });
  }

  async processImages(answers: any, submissionId?: string) {
    return this.request<any>("/responses/process-images", {
      method: "POST",
      body: JSON.stringify({ answers, submissionId }),
    });
  }

  async convertImageUrl(imageUrl: string) {
    return this.request<{ cloudinaryUrl: string }>("/responses/convert-image", {
      method: "POST",
      body: JSON.stringify({ imageUrl }),
    });
  }

  async submitResponse(
    formId: string,
    responseData: any & {
      startedAt?: Date | string;
      completedAt?: Date | string;
      sessionId?: string;
      isSectionSubmit?: boolean;
      sectionIndex?: number;
    },
  ) {
    return this.request<{ response: any }>(`/responses/${formId}`, {
      method: "POST",
      body: JSON.stringify(responseData),
    });
  }

  async getResponsesByModel(
    formId: string,
    tenantSlug: string | undefined,
    modelQuestionId: string,
    modelNumber: string,
  ): Promise<any[]> {
    try {
      // Try the dedicated endpoint first (local)
      let allResponses = [];
      let endpoint = tenantSlug
        ? `/responses/${tenantSlug}/forms/${formId}/responses/by-model`
        : `/responses/form/${formId}/responses/by-model`;

      try {
        console.log('📡 Trying dedicated endpoint:', endpoint);
        const data = await this.request<any>(`${endpoint}?modelNumber=${modelNumber}&modelQuestionId=${modelQuestionId}`);
        allResponses = data?.data || data || [];
        console.log('✅ Dedicated endpoint worked!');
      } catch (error: any) {
        // If 404, fallback to byForm endpoint
        if (error?.response?.status === 404) {
          console.log('🔄 Dedicated endpoint not found, falling back to byForm');
          const fallbackEndpoint = `/responses/form/${formId}${tenantSlug ? `?tenant=${tenantSlug}` : ''}`;
          const data = await this.request<any>(fallbackEndpoint);
          const allFormResponses = data?.data || data || [];

          // Filter by model number
          allResponses = allFormResponses.filter((response: any) => {
            const modelAnswer = response.answers?.[modelQuestionId] || response[modelQuestionId];
            return String(modelAnswer).toLowerCase() === String(modelNumber).toLowerCase().trim();
          });
        } else {
          throw error;
        }
      }

      console.log(`📊 Found ${allResponses.length} historical responses`);
      return allResponses;

    } catch (error) {
      console.error('❌ Error fetching responses by model:', error);
      return [];
    }
  }

  async updateResponse(id: string, responseData: any) {
    console.log("API Client: Updating response", id, responseData);
    const result = await this.request<{ response: any }>(`/responses/${id}`, {
      method: "PUT",
      body: JSON.stringify(responseData),
    });
    console.log("API Client: Update response result", result);
    return result;
  }

  async deleteResponse(id: string) {
    return this.request(`/responses/${id}`, {
      method: "DELETE",
    });
  }

  async assignResponse(id: string, assignedTo: string) {
    return this.request(`/responses/${id}/assign`, {
      method: "PATCH",
      body: JSON.stringify({ assignedTo }),
    });
  }

  async exportFormResponses(formId: string, format: "excel" | "csv" = "excel") {
    const headers: Record<string, string> = {};

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    const response = await fetch(
      `${this.baseUrl}/responses/form/${formId}/export?format=${format}`,
      {
        headers,
      },
    );

    if (!response.ok) {
      throw new ApiError(response.status, null, "Export failed");
    }

    return response.blob();
  }

  // Analytics
  async getDashboardAnalytics() {
    return this.request<any>("/analytics/dashboard");
  }

  async getFormAnalytics(formId: string) {
    return this.request<any>(`/analytics/form/${formId}`);
  }

  async getMyReviewStats() {
    return this.get<{
      totalResponses: number;
      reviewed: number;
      accepted: number;
      rejected: number;
      rework: number;
      performanceScore: number;
    }>("/analytics/my-review-stats");
  }

  async getPerformanceTable(params?: { startDate?: string; endDate?: string }) {
    let url = "/analytics/performance-table";
    if (params) {
      const searchParams = new URLSearchParams();
      if (params.startDate) searchParams.append("startDate", params.startDate);
      if (params.endDate) searchParams.append("endDate", params.endDate);
      const query = searchParams.toString();
      if (query) url += `?${query}`;
    }
    return this.get<any[]>(url);
  }

  // ── Form Session Tracking ─────────────────────────────────────────────────
  // Start a time-tracking session when a user opens a form
  // NOTE: Uses raw fetch because the backend returns { success, sessionId, startedAt }
  // at the top level (not wrapped in data:{}), so the standard request() wrapper won't work.
  async startFormSession(
    formId: string,
    formTitle?: string,
  ): Promise<{ sessionId: string; startedAt: string }> {
    try {
      const headers: Record<string, string> = {
        "Content-Type": "application/json",
      };
      if (this.token) headers.Authorization = `Bearer ${this.token}`;

      const res = await fetch(`${this.baseUrl}/forms/${formId}/track/start`, {
        method: "POST",
        headers,
        body: JSON.stringify({ formTitle }),
      });
      const json = await res.json();

      if (json.sessionId) {
        return {
          sessionId: json.sessionId,
          startedAt: json.startedAt || new Date().toISOString(),
        };
      }
      // Also handle if backend ever wraps it: { success, data: { sessionId } }
      if (json.data?.sessionId) {
        return {
          sessionId: json.data.sessionId,
          startedAt: json.data.startedAt || new Date().toISOString(),
        };
      }
      throw new Error("No sessionId in response");
    } catch (err) {
      console.warn(
        "[apiClient.startFormSession] Failed, using client-side fallback:",
        err,
      );
      // Fallback: local session — startedAt still tracked correctly on the client
      return {
        sessionId: `local-${Date.now()}`,
        startedAt: new Date().toISOString(),
      };
    }
  }

  // Track time spent on an individual question
  async trackQuestionTime(
    formId: string,
    data: {
      sessionId: string;
      questionId: string;
      questionText?: string;
      questionType?: string;
      sectionId?: string;
      sectionTitle?: string;
      timeSpent: number;
      answer?: any;
    },
  ) {
    try {
      return await this.request<any>(`/forms/${formId}/track/question`, {
        method: "POST",
        body: JSON.stringify(data),
      });
    } catch (err) {
      console.warn("[apiClient.trackQuestionTime] Failed (non-critical):", err);
    }
  }

  // Track section completion progress
  async trackFormProgress(
    formId: string,
    data: {
      sessionId: string;
      sectionId?: string;
      sectionTitle?: string;
      timeSpent?: number;
      questionCount?: number;
    },
  ) {
    try {
      return await this.request<any>(`/forms/${formId}/track/progress`, {
        method: "POST",
        body: JSON.stringify(data),
      });
    } catch (err) {
      console.warn("[apiClient.trackFormProgress] Failed (non-critical):", err);
    }
  }

  // Mark the session as complete just before form submission
  async trackFormComplete(
    formId: string,
    data: {
      sessionId: string;
      answers?: any;
    },
  ) {
    try {
      return await this.request<any>(`/forms/${formId}/track/complete`, {
        method: "POST",
        body: JSON.stringify(data),
      });
    } catch (err) {
      console.warn("[apiClient.trackFormComplete] Failed (non-critical):", err);
    }
  }

  // Send a periodic heartbeat to show the session is still alive
  async sendHeartbeat(data: { url?: string; formSessionId?: string }) {
    try {
      return await this.request<any>(`/activity/heartbeat`, {
        method: "POST",
        body: JSON.stringify(data),
      });
    } catch (err) {
      // Heartbeat failures are expected (network issues, etc.) – silently ignore
    }
  }

  // Rank tracking
  async getResponseRank(
    formId: string,
    questionId: string,
    answer: any,
    tenantSlug?: string,
  ) {
    let endpoint = tenantSlug
      ? `/responses/${tenantSlug}/forms/${formId}/rank`
      : `/responses/rank`;

    // Handle array answers (checkboxes) by converting to string if needed
    // The backend expects a single value for matching
    const answerParam = Array.isArray(answer) ? JSON.stringify(answer) : answer;

    const queryParams = new URLSearchParams({
      formId,
      questionId,
      answer: String(answerParam),
    });

    return this.request<{ rank: number }>(
      `${endpoint}?${queryParams.toString()}`,
    );
  }

  async getGlobalFormStats(formId: string) {
    return this.request<{ stats: any[] }>(`/forms/${formId}/global-stats`);
  }

  async getUserAnalytics() {
    return this.request<any>("/analytics/users");
  }

  // OTP Verification
  async sendOtp(mobile: string) {
    return this.request<{ success: boolean; message: string; otp?: string }>(
      "/otp/send",
      {
        method: "POST",
        body: JSON.stringify({ mobile }),
      },
    );
  }

  async verifyOtp(mobile: string, otp: string) {
    return this.request<{ success: boolean; message: string }>("/otp/verify", {
      method: "POST",
      body: JSON.stringify({ mobile, otp }),
    });
  }

  async getSuggestedAnswers(
    formId: string,
    questionId: string,
    answer: any,
    tenantSlug?: string,
  ) {
    let endpoint = tenantSlug
      ? `/responses/${tenantSlug}/forms/${formId}/suggestions`
      : `/responses/suggestions`;

    const queryParams = new URLSearchParams({
      formId,
      questionId,
      answer: String(answer),
    });

    return this.request<{ suggestedAnswers: Record<string, any> }>(
      `${endpoint}?${queryParams.toString()}`,
    );
  }

  async getQuestionPreviousAnswers(
    formId: string,
    questionId: string,
    tenantSlug?: string,
  ) {
    let endpoint = tenantSlug
      ? `/responses/${tenantSlug}/forms/${formId}/previous-answers`
      : `/responses/previous-answers`;

    const queryParams = new URLSearchParams({
      formId,
      questionId,
    });

    return this.request<{ answers: string[] }>(
      `${endpoint}?${queryParams.toString()}`,
    );
  }

  async getAdminPerformance(
    adminId: string,
    params?: { startDate?: string; endDate?: string },
  ) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);

    const endpoint = `/analytics/admin/${adminId}/performance${query.toString() ? `?${query.toString()}` : ""}`;
    return this.request<any>(endpoint);
  }

  async getAdminActivity(
    adminId: string,
    params?: { startDate?: string; endDate?: string; limit?: number },
  ) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.limit !== undefined)
      query.set("limit", params.limit.toString());

    const endpoint = `/analytics/admin/${adminId}/activity${query.toString() ? `?${query.toString()}` : ""}`;
    return this.request<any>(endpoint);
  }
  // Roles

  async getAdminResponseDetails(
    adminId: string,
    params?: { startDate?: string; endDate?: string },
  ) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);

    const endpoint = `/analytics/admin/${adminId}/response-details${query.toString() ? `?${query.toString()}` : ""}`;
    return this.request<{
      totalResponses: number;
      statusBreakdown: { pending: number; verified: number; rejected: number };
      yesNoNA: { yes: number; no: number; na: number };
      formBreakdown: {
        formId: string;
        formTitle: string;
        yes: number;
        no: number;
        na: number;
        responseCount: number; // Updated name
        avgTimeSpent?: number;
        sessionCount?: number;
      }[];
      personalSubmissions?: {
        id: string;
        formTitle: string;
        submittedAt: string;
        status: string;
      }[];
    }>(endpoint);
  }

  async getTenantSubmissionStats(
    tenantId?: string,
    params?: { startDate?: string; endDate?: string },
  ) {
    const query = new URLSearchParams();
    if (tenantId) query.set("tenantId", tenantId);
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    return this.request<any>(
      `/analytics/tenant/stats${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }
  async getTenantForms(tenantId?: string) {
    // Re-uses the newly added tenant submission stats.
    return this.getTenantSubmissionStats(tenantId);
  }

  // Roles
  async getRoles() {
    return this.request<{ roles: any[] }>("/roles");
  }

  async createRole(roleData: any) {
    return this.request<{ role: any }>("/roles", {
      method: "POST",
      body: JSON.stringify(roleData),
    });
  }

  async updateRole(id: string, roleData: any) {
    return this.request<{ role: any }>(`/roles/${id}`, {
      method: "PUT",
      body: JSON.stringify(roleData),
    });
  }

  async deleteRole(id: string) {
    return this.request(`/roles/${id}`, {
      method: "DELETE",
    });
  }

  async getAvailablePermissions() {
    return this.request<{ permissions: string[] }>("/roles/permissions");
  }

  async assignRole(userId: string, roleId: string) {
    return this.request("/roles/assign", {
      method: "POST",
      body: JSON.stringify({ userId, roleId }),
    });
  }

  async getUsersByRole(roleId: string) {
    return this.request<{ users: any[]; role: any; pagination: any }>(
      `/roles/${roleId}/users`,
    );
  }

  // Files
  async uploadFile(
    file: File,
    category: string = "general",
    associatedId?: string,
    onProgress?: (progress: {
      percentage: number;
      loaded: number;
      total: number;
      timeRemaining?: number;
      speed?: number;
    }) => void,
  ) {
    // Validate file size (10MB limit)
    const maxSize = 10 * 1024 * 1024; // 10MB
    if (file.size > maxSize) {
      throw new ApiError(
        400,
        null,
        `File size (${(file.size / 1024 / 1024).toFixed(2)}MB) exceeds maximum limit of 10MB`,
      );
    }

    try {
      // STEP 1: Get presigned URL from your backend
      console.log("Requesting presigned URL for:", file.name);

      // Normalize category (fix for 'form' vs 'forms' issue)
      const normalizedCategory = category === "form" ? "forms" : category;

      console.log("Original category:", category);
      console.log("Normalized category:", normalizedCategory);

      // Use the baseUrl from ApiClient
      const presignedEndpoint = "/upload/presigned-url";

      // ✅ FIX: Define uploadApiUrl here where it's accessible
      const uploadApiUrl = `${this.baseUrl}${presignedEndpoint}`;
      console.log("Upload API URL:", uploadApiUrl);

      // Prepare headers
      const headers: Record<string, string> = {
        "Content-Type": "application/json",
      };

      if (this.token) {
        headers["Authorization"] = `Bearer ${this.token}`;
      }

      const presignedResponse = await fetch(uploadApiUrl, {
        method: "POST",
        headers,
        body: JSON.stringify({
          filename: file.name,
          fileType: file.type,
          category: normalizedCategory,
          ...(associatedId && { associatedId }),
        }),
      });

      if (!presignedResponse.ok) {
        const errorText = await presignedResponse.text();
        console.error(
          "Presigned URL request failed:",
          presignedResponse.status,
          errorText,
        );
        throw new ApiError(
          presignedResponse.status,
          null,
          "Failed to get upload URL",
        );
      }

      const presignedData = await presignedResponse.json();

      if (!presignedData.success) {
        throw new ApiError(
          500,
          presignedData,
          presignedData.error || "Invalid response from server",
        );
      }

      const { uploadUrl, key, publicUrl } = presignedData;
      console.log("Received presigned URL for S3 upload");

      // STEP 2: Upload directly to S3 using XMLHttpRequest
      return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        const startTime = Date.now();

        xhr.upload.addEventListener("progress", (event) => {
          if (event.lengthComputable && onProgress) {
            const percentage = Math.round((event.loaded / event.total) * 100);
            const elapsed = Date.now() - startTime;
            const speed = event.loaded / (elapsed / 1000);
            const remaining = (event.total - event.loaded) / speed;
            const timeRemaining = isFinite(remaining)
              ? Math.round(remaining)
              : undefined;

            onProgress({
              percentage,
              loaded: event.loaded,
              total: event.total,
              timeRemaining,
              speed,
            });
          }
        });

        xhr.addEventListener("load", () => {
          console.log("S3 upload response status:", xhr.status);

          if (xhr.status === 200) {
            // Success! Return the CloudFront URL
            resolve({
              url: publicUrl,
              file: {
                url: publicUrl,
                filename: key,
                originalName: file.name,
                size: file.size,
                s3Key: key,
                uploadedAt: new Date().toISOString(),
              },
            });
          } else {
            console.error("S3 upload failed:", xhr.status, xhr.responseText);
            reject(
              new ApiError(
                xhr.status,
                null,
                `S3 upload failed with status ${xhr.status}`,
              ),
            );
          }
        });

        xhr.addEventListener("error", () => {
          reject(new ApiError(0, null, "Network error during S3 upload"));
        });

        xhr.addEventListener("abort", () => {
          reject(new ApiError(0, null, "Upload was cancelled"));
        });

        // Upload directly to S3
        xhr.open("PUT", uploadUrl);
        xhr.setRequestHeader("Content-Type", file.type);
        xhr.send(file);
      });
    } catch (error) {
      if (error instanceof ApiError) {
        throw error;
      }
      throw new ApiError(
        500,
        null,
        `Upload failed: ${(error as Error).message}`,
      );
    }
  }

  resolveUploadedFileUrl(uploadResult: any) {
    console.log("[resolveUploadedFileUrl] Input:", uploadResult);

    if (!uploadResult) {
      console.warn("[resolveUploadedFileUrl] No upload result provided");
      return "";
    }

    if (typeof uploadResult === "string") {
      console.log(
        "[resolveUploadedFileUrl] Returning string URL:",
        uploadResult,
      );
      return this.ensureAbsoluteFileUrl(uploadResult);
    }

    if (uploadResult.publicUrl) {
      console.log(
        "[resolveUploadedFileUrl] Using publicUrl:",
        uploadResult.publicUrl,
      );
      return this.ensureAbsoluteFileUrl(uploadResult.publicUrl);
    }

    if (uploadResult.url) {
      console.log("[resolveUploadedFileUrl] Using url:", uploadResult.url);
      return this.ensureAbsoluteFileUrl(uploadResult.url);
    }

    if (uploadResult.secureUrl) {
      console.log(
        "[resolveUploadedFileUrl] Using secureUrl:",
        uploadResult.secureUrl,
      );
      return this.ensureAbsoluteFileUrl(uploadResult.secureUrl);
    }

    if (uploadResult.location) {
      console.log(
        "[resolveUploadedFileUrl] Using location:",
        uploadResult.location,
      );
      return this.ensureAbsoluteFileUrl(uploadResult.location);
    }

    if (uploadResult.file && uploadResult.file.url) {
      console.log(
        "[resolveUploadedFileUrl] Using file.url:",
        uploadResult.file.url,
      );
      return this.ensureAbsoluteFileUrl(uploadResult.file.url);
    }

    console.warn(
      "[resolveUploadedFileUrl] No URL found in result:",
      uploadResult,
    );
    return "";
  }

  async getUserFiles() {
    return this.request<{ files: any[] }>("/files");
  }

  async deleteFile(id: string) {
    return this.request(`/files/${id}`, {
      method: "DELETE",
    });
  }

  getFileUrl(filename: string) {
    return `${this.baseUrl}/files/${filename}`;
  }

  // Profile
  async getUserProfile() {
    return this.request<{ profile: any }>("/profile");
  }

  async updateUserProfile(profileData: any) {
    return this.request<{ profile: any; user?: any }>("/profile", {
      method: "PUT",
      body: JSON.stringify(profileData),
    });
  }

  async updateProfileSettings(settings: any) {
    return this.request<{ profile: any }>("/profile/settings", {
      method: "PATCH",
      body: JSON.stringify(settings),
    });
  }

  async uploadAvatar(file: File) {
    const formData = new FormData();
    formData.append("avatar", file);

    const headers: Record<string, string> = {};

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    const response = await fetch(`${this.baseUrl}/profile/avatar`, {
      method: "POST",
      headers,
      body: formData,
    });

    const data: ApiResponse<any> = await response.json();

    if (!response.ok || !data.success) {
      throw new ApiError(response.status, data, data.message);
    }

    return data.data;
  }

  // Tenants (SuperAdmin only)
  async getTenants(search: string = "", status: string = "all") {
    return this.request<{ tenants: any[]; total: number }>(
      `/tenants?search=${search}&status=${status}`,
    );
  }

  async getTenantsMinimal() {
    return this.request<{ tenants: any[] }>("/tenants/minimal");
  }

  async getTenant(tenantId: string) {
    return this.request<{ tenant: any }>(`/tenants/${tenantId}`);
  }

  async createTenant(tenantData: any) {
    return this.request<{ tenant: any; admin: any }>("/tenants", {
      method: "POST",
      body: JSON.stringify(tenantData),
    });
  }

  async updateTenant(tenantId: string, tenantData: any) {
    return this.request<{ tenant: any; admin?: any }>(`/tenants/${tenantId}`, {
      method: "PUT",
      body: JSON.stringify(tenantData),
    });
  }

  async deleteTenant(tenantId: string) {
    return this.request(`/tenants/${tenantId}`, {
      method: "DELETE",
    });
  }

  async getTenantStats(tenantId: string) {
    return this.request<{ stats: any }>(`/tenants/${tenantId}/stats`);
  }

  async getOfficeLocation() {
    return this.get<{ lat: number; lng: number; radius: number }>(
      "/tenants/office-location",
    );
  }

  async getTenantMessages() {
    return this.request<any[]>("/messages/tenant-messages");
  }

  async updateOfficeLocation(data: {
    lat: number;
    lng: number;
    radius: number;
  }) {
    console.log("API Client: Updating office location", data);
    const result = await this.request<{ success: boolean }>(
      "/tenants/office-location",
      {
        method: "PUT",
        body: JSON.stringify(data),
      },
    );
    console.log("API Client: Update office location result", result);
    return { success: true, data: result };
  }

  async toggleTenantStatus(tenantId: string) {
    return this.request<{ tenant: any }>(`/tenants/${tenantId}/toggle-status`, {
      method: "PATCH",
    });
  }

  async addAdminToTenant(
    tenantId: string,
    adminData: {
      firstName: string;
      lastName: string;
      email: string;
      password: string;
    },
  ) {
    return this.request<{
      admin: {
        _id: string;
        firstName: string;
        lastName: string;
        email: string;
        role: string;
        isActive: boolean;
      };
    }>(`/tenants/${tenantId}/add-admin`, {
      method: "POST",
      body: JSON.stringify(adminData),
    });
  }

  async removeAdminFromTenant(tenantId: string, adminId: string) {
    return this.request(`/tenants/${tenantId}/remove-admin/${adminId}`, {
      method: "DELETE",
    });
  }

  // Parameters
  async getParameters(params?: {
    type?: "main" | "followup";
    search?: string;
    formId?: string;
  }) {
    const query = new URLSearchParams();

    if (params?.type) {
      query.set("type", params.type);
    }

    if (params?.search) {
      query.set("search", params.search);
    }

    if (params?.formId) {
      query.set("formId", params.formId);
    }

    const endpoint = `/parameters${query.toString() ? `?${query.toString()}` : ""
      }`;

    return this.request<{ parameters: any[] }>(endpoint);
  }

  async createParameter(parameterData: {
    name: string;
    type: "main" | "followup";
    formId: string;
    tenantId?: string;
  }) {
    return this.request<{ parameter: any }>("/parameters", {
      method: "POST",
      body: JSON.stringify(parameterData),
    });
  }

  async updateParameter(
    id: string,
    parameterData: {
      name: string;
      type: "main" | "followup";
      formId: string;
    },
  ) {
    return this.request<{ parameter: any }>(`/parameters/${id}`, {
      method: "PUT",
      body: JSON.stringify(parameterData),
    });
  }

  async deleteParameter(id: string) {
    return this.request(`/parameters/${id}`, {
      method: "DELETE",
    });
  }
  async processBulkImages(answers: any, batchId?: string) {
    return this.request<any>("/responses/process-bulk-images", {
      method: "POST",

      body: JSON.stringify({ answers, batchId }),
    });
  }

  async generatePDF(options: {
    htmlContent: string;
    filename?: string;
    format?: "custom" | "a4" | "a4-portrait";
    compressed?: boolean;
  }): Promise<Blob> {
    const url = `${this.baseUrl}/pdf/generate`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    const requestBody = {
      htmlContent: options.htmlContent,
      filename: options.filename || "report.pdf",
      format: options.format || "custom",
      compressed: false,
    };

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 120000); // 2 minute timeout

      const response = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify(requestBody),
        signal: controller.signal,
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        console.error("API Error Response:", errorData);
        throw new Error(
          errorData.details ||
          errorData.error ||
          `PDF generation failed: ${response.statusText}`,
        );
      }

      return response.blob();
    } catch (error: any) {
      if (error.name === "AbortError") {
        throw new Error(
          "PDF generation timed out (120s). The document may be too large. Try with fewer sections or images.",
        );
      }
      throw error;
    }
  }


  async generateOPSPDF(data: {
    htmlContent: string;
    filename: string;
  }): Promise<Blob> {
    console.log('📄 Generating OPS PDF via server...');

    const url = `${this.baseUrl}/pdf/generate-ops`;

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    };

    if (this.token) {
      headers['Authorization'] = `Bearer ${this.token}`;
    }

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          htmlContent: data.htmlContent,
          filename: data.filename
        }),
      });

      if (!response.ok) {
        const error = await response.json().catch(() => ({}));
        throw new Error(error.details || error.error || `Failed to generate OPS PDF: ${response.status}`);
      }

      return response.blob();
    } catch (error: any) {
      console.error('OPS PDF generation error:', error);
      throw new Error(error.message || 'Failed to generate OPS PDF');
    }
  }

  // Add helper method for blob responses if not exists

  // Form Invite Management
  async uploadInvites(formId: string, formData: FormData) {
    const url = `${this.baseUrl}/forms/${formId}/invites/upload`;

    const headers: Record<string, string> = {};

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    // Don't set Content-Type for FormData - browser will set it automatically with boundary

    const response = await fetch(url, {
      method: "POST",
      headers,
      body: formData,
    });

    const data: ApiResponse<any> = await response.json();

    if (!response.ok || !data.success) {
      throw new ApiError(response.status, data, data.message);
    }

    return data;
  }

  async sendInvites(
    formId: string,
    data: { emails: Array<{ email: string; phone?: string }> },
  ) {
    const response = await this.request<any>(`/forms/${formId}/invites/send`, {
      method: "POST",
      body: JSON.stringify(data),
    });

    // Your backend returns {success: true, message: '...', data: {...}}
    // The this.request() method already extracts the 'data' field
    // So 'response' here IS the data object
    return {
      success: true, // Manually add this back
      message: "Invites processed successfully",
      data: response, // response contains {total, successful, failed, results, failures}
    };
  }

  async getInviteStats(formId: string) {
    return this.request<{
      form: any;
      invites: {
        total: number;
        sent: number;
        responded: number;
        expired: number;
        responseRate: number;
      };
      responses: {
        total: number;
        invited: number;
        public: number;
      };
    }>(`/forms/${formId}/invites/stats`);
  }

  async getFormInvites(
    formId: string,
    options?: {
      page?: number;
      limit?: number;
      search?: string;
      status?: string;
      dateFilter?: string;
      startDate?: string;
      endDate?: string;
      sortBy?: string;
      sortOrder?: string;
    },
  ) {
    const query = new URLSearchParams();

    if (options) {
      Object.entries(options).forEach(([key, value]) => {
        if (value !== undefined && value !== "") {
          query.append(key, value.toString());
        }
      });
    }

    const endpoint = `/forms/${formId}/invites${query.toString() ? `?${query.toString()}` : ""
      }`;
    const url = `${this.baseUrl}${endpoint}`;

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    const response = await fetch(url, {
      headers,
    });

    const data: ApiResponse<any> = await response.json();

    if (!response.ok || !data.success) {
      throw new ApiError(
        response.status,
        data,
        data.message || "Failed to fetch invites",
      );
    }

    return data.data;
  }

  // WhatsApp Invite Management
  async uploadWhatsAppInvites(formId: string, formData: FormData) {
    const url = `${this.baseUrl}/forms/${formId}/invites/whatsapp/upload`;

    const headers: Record<string, string> = {};

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    const response = await fetch(url, {
      method: "POST",
      headers,
      body: formData,
    });

    const data: ApiResponse<any> = await response.json();

    if (!response.ok || !data.success) {
      throw new ApiError(response.status, data, data.message);
    }

    return data;
  }

  async sendWhatsAppInvites(
    formId: string,
    data: { phones: Array<{ phone: string; email?: string }> },
  ) {
    const response = await this.request<any>(
      `/forms/${formId}/invites/whatsapp/send`,
      {
        method: "POST",
        body: JSON.stringify(data),
      },
    );

    return {
      success: true,
      message: "WhatsApp invites processed successfully",
      data: response,
    };
  }

  // SMS Invite Management
  async sendSMSInvites(
    formId: string,
    data: { phones: Array<{ phone: string; email?: string }> },
  ) {
    const response = await this.request<any>(
      `/forms/${formId}/invites/sms/send`,
      {
        method: "POST",
        body: JSON.stringify(data),
      },
    );

    return {
      success: true,
      message: "SMS invites processed successfully",
      data: response,
    };
  }

  async getUnassignedResponses(params?: {
    tenantId?: string;
    startDate?: string;
    endDate?: string;
    limit?: number;
  }) {
    const query = new URLSearchParams();
    if (params?.tenantId) query.set("tenantId", params.tenantId);
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.limit) query.set("limit", params.limit.toString());

    const endpoint = `/responses/unassigned${query.toString() ? `?${query.toString()}` : ""}`;
    return this.request<{ responses: any[]; total: number; hasMore: boolean }>(
      endpoint,
    );
  }

  async assignResponses(data: { responseIds: string[]; adminId: string }) {
    return this.request<{ modifiedCount: number }>("/responses/assign", {
      method: "POST",
      body: JSON.stringify(data),
    });
  }

  async autoAssignResponse(
    responseId: string,
    data: {
      tenantId: string;
    },
  ) {
    console.log(`[API] Calling autoAssignResponse with:`, { responseId, data });
    try {
      const result = await this.request<{
        responseId: string;
        assignedTo: string;
        adminLoad: Record<string, number>;
      }>(`/responses/${responseId}/auto-assign`, {
        method: "POST",
        body: JSON.stringify(data),
      });
      console.log(`[API] autoAssignResponse result:`, result);
      return result;
    } catch (error) {
      console.error(`[API] autoAssignResponse error:`, error);
      throw error;
    }
  }

  // Also add this helper method to get available admins for assignment
  async getAvailableAdmins(params?: { tenantId?: string; role?: string }) {
    const query = new URLSearchParams();
    if (params?.tenantId) query.set("tenantId", params.tenantId);
    if (params?.role) query.set("role", params.role);

    const endpoint = `/users/available${query.toString() ? `?${query.toString()}` : ""}`;
    return this.request<{ users: any[] }>(endpoint);
  }

  // Activity Analytics
  async getTenantActivitySummary(
    tenantId: string,
    params?: { startDate?: string; endDate?: string; role?: string },
  ) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.role) query.set("role", params.role);

    return this.request<any>(
      `/activity/tenant/${tenantId}/summary${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }

  async getUserTimeline(
    userId: string,
    params?: { limit?: number; startDate?: string; endDate?: string },
  ) {
    const query = new URLSearchParams();
    if (params?.limit) query.set("limit", params.limit.toString());
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);

    return this.request<any>(
      `/activity/user/${userId}/timeline${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }

  async getFormSessions(
    formId: string,
    params?: { startDate?: string; endDate?: string },
  ) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);

    return this.request<any>(
      `/activity/forms/${formId}/sessions${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }

  // Attendance
  async getAttendance(params?: {
    startDate?: string;
    endDate?: string;
    userId?: string;
    page?: number;
    limit?: number;
    tenantId?: string;
  }) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.userId) query.set("userId", params.userId);
    if (params?.page) query.set("page", params.page.toString());
    if (params?.limit) query.set("limit", params.limit.toString());
    if (params?.tenantId) query.set("tenantId", params.tenantId);

    return this.request<{ logs: any[]; pagination: any }>(
      `/attendance${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }

  async getAttendanceSummary(date?: string) {
    const query = date ? `?date=${date}` : "";
    return this.request<{
      date: string;
      totalUsers: number;
      present: number;
      absent: number;
      activeNow: number;
      users: any[];
    }>(`/attendance/summary${query}`);
  }

  async getMyAttendance(params?: {
    startDate?: string;
    endDate?: string;
    page?: number;
    limit?: number;
  }) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.page) query.set("page", params.page.toString());
    if (params?.limit) query.set("limit", params.limit.toString());

    return this.request<{ logs: any[]; summary: any; pagination: any }>(
      `/attendance/my${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }

  async getAttendanceUsers() {
    return this.request<{ users: any[] }>("/attendance/users");
  }

  async exportAttendance(params?: {
    startDate?: string;
    endDate?: string;
    userId?: string;
    tenantId?: string;
  }) {
    const query = new URLSearchParams();
    if (params?.startDate) query.set("startDate", params.startDate);
    if (params?.endDate) query.set("endDate", params.endDate);
    if (params?.userId) query.set("userId", params.userId);
    if (params?.tenantId) query.set("tenantId", params.tenantId);

    return this.request<any[]>(
      `/attendance/export${query.toString() ? `?${query.toString()}` : ""}`,
    );
  }

  async updateLastActive(sessionLogId: string) {
    return this.request("/attendance/heartbeat", {
      method: "POST",
      body: JSON.stringify({ sessionLogId }),
    });
  }

  async updateLoginLocation(sessionLogId: string, location: any) {
    return this.request("/attendance/login-location", {
      method: "PUT",
      body: JSON.stringify({ sessionLogId, location }),
    });
  }

  async createShift(shiftData: any) {
    return this.request("/hr/shifts", {
      method: "POST",
      body: JSON.stringify(shiftData),
    });
  }

  async getShifts() {
    return this.request<{ data: any[] }>("/hr/shifts");
  }

  async updateShift(id: string, shiftData: any) {
    return this.request(`/hr/shifts/${id}`, {
      method: "PUT",
      body: JSON.stringify(shiftData),
    });
  }

  async deleteShift(id: string) {
    return this.request(`/hr/shifts/${id}`, {
      method: "DELETE",
    });
  }

  async getAvailableInspectors() {
    return this.request<{ data: any[] }>("/hr/shifts/available-inspectors");
  }

  async assignInspectorsToShift(shiftId: string, inspectorIds: string[]) {
    return this.request(`/hr/shifts/${shiftId}/assign`, {
      method: "POST",
      body: JSON.stringify({ inspectorIds }),
    });
  }

  async removeInspectorsFromShift(shiftId: string, inspectorIds: string[]) {
    return this.request(`/hr/shifts/${shiftId}/remove`, {
      method: "DELETE",
      body: JSON.stringify({ inspectorIds }),
    });
  }

  async sendAttendanceOTP() {
    return this.request("/hr/attendance/send-otp", {
      method: "POST",
    });
  }

  async checkIn(location: {
    lat: number;
    lng: number;
    accuracy: number;
    otp: string;
  }) {
    return this.request<{ data: any }>("/hr/attendance/checkin", {
      method: "POST",
      body: JSON.stringify(location),
    });
  }

  async checkOut(location: { lat: number; lng: number; accuracy: number }) {
    return this.request<{ data: any }>("/hr/attendance/checkout", {
      method: "POST",
      body: JSON.stringify(location),
    });
  }

  async getMyHRAttendanceStatus() {
    return this.request<{ data: any }>("/hr/attendance/my-status");
  }

  async getMyHRAttendanceHistory(params?: { page?: number; limit?: number }) {
    const query = new URLSearchParams();
    if (params?.page) query.set("page", params.page.toString());
    if (params?.limit) query.set("limit", params.limit.toString());
    return this.request<{ data: any }>(
      "/hr/attendance/my-history" +
      (query.toString() ? `?${query.toString()}` : ""),
    );
  }

  async getMyHRShift() {
    return this.request<{ data: any }>("/hr/attendance/my-shift");
  }

  async getHRAttendanceReport(params: {
    startDate: string;
    endDate: string;
    inspectorId?: string;
    status?: string;
    shiftId?: string;
  }) {
    const query = new URLSearchParams();
    Object.entries(params).forEach(([key, value]) => {
      if (value) query.set(key, value);
    });
    console.log("getHRAttendanceReport - query:", query.toString());
    return this.request<{ data: any }>(
      "/hr/attendance/report?" + query.toString(),
    );
  }

  async exportHRAttendanceReport(params: {
    startDate: string;
    endDate: string;
    inspectorId?: string;
    status?: string;
    shiftId?: string;
  }) {
    const query = new URLSearchParams();
    Object.entries(params).forEach(([key, value]) => {
      if (value) query.set(key, value);
    });

    const headers: Record<string, string> = {};
    if (this.token) headers.Authorization = `Bearer ${this.token}`;

    const response = await fetch(
      `${this.baseUrl}/hr/attendance/export?${query.toString()}`,
      { headers },
    );
    if (!response.ok)
      throw new ApiError(response.status, null, "Export failed");
    return response.blob();
  }

  async getHRTenantSummary() {
    return this.request<{ data: any }>("/hr/attendance/summary");
  }

  // --- LEAVE MANAGEMENT ---
  async applyLeave(data: {
    leaveType: string;
    startDate: string;
    endDate: string;
    reason: string;
  }) {
    return this.request<{ success: boolean; data: any }>("/hr/leaves/apply", {
      method: "POST",
      body: JSON.stringify(data),
    });
  }

  async getMyLeaves(params?: { page?: number; limit?: number }) {
    const query = new URLSearchParams();
    if (params?.page) query.set("page", params.page.toString());
    if (params?.limit) query.set("limit", params.limit.toString());
    const url =
      "/hr/leaves/my" + (query.toString() ? `?${query.toString()}` : "");
    const res = await this.request<{
      success: boolean;
      data: any[];
      pagination?: any;
    }>(url);
    console.log("getMyLeaves raw response:", res);
    return res;
  }

  async getAllLeaves(
    params?:
      | { page?: number; limit?: number; getAll?: boolean; status?: string }
      | string,
  ) {
    let url = "/hr/leaves/all";
    if (typeof params === "string") {
      url += params ? `?status=${params}` : "";
    } else if (params) {
      const query = new URLSearchParams();
      if (params.page) query.set("page", params.page.toString());
      if (params.limit) query.set("limit", params.limit.toString());
      if (params.getAll) query.set("getAll", "true");
      if (params.status) query.set("status", params.status);
      if (query.toString()) url += `?${query.toString()}`;
    }
    return this.request<{ success: boolean; data: any[]; pagination?: any }>(
      url,
    );
  }

  async updateLeaveStatus(
    id: string,
    status: "approved" | "rejected",
    comments?: string,
  ) {
    return this.request<{ success: boolean; data: any }>(
      `/hr/leaves/${id}/status`,
      {
        method: "PUT",
        body: JSON.stringify({ status, comments }),
      },
    );
  }

  // --- PERMISSION MANAGEMENT ---
  async applyPermission(data: {
    permissionType: string;
    date: string;
    startTime: string;
    endTime: string;
    duration: number;
    reason: string;
  }) {
    return this.request<{ success: boolean; data: any }>(
      "/hr/permissions/apply",
      {
        method: "POST",
        body: JSON.stringify(data),
      },
    );
  }

  async getMyPermissions() {
    return this.request<{ success: boolean; data: any[] }>(
      "/hr/permissions/my",
    );
  }

  async getAllPermissions(status?: string) {
    const url = "/hr/permissions/all" + (status ? `?status=${status}` : "");
    return this.request<{ success: boolean; data: any[] }>(url);
  }

  async updatePermissionStatus(id: string, status: "approved" | "rejected") {
    return this.request<{ success: boolean; data: any }>(
      `/hr/permissions/${id}/status`,
      {
        method: "PUT",
        body: JSON.stringify({ status }),
      },
    );
  }

  // --- NOTIFICATIONS ---
  async getMyNotifications() {
    if (!this.token) throw new Error("No auth token");

    const url = `${this.baseUrl}/hr/notifications/my`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${this.token}`,
    };

    const response = await fetch(url, { headers });
    const data = await response.json();

    // Return both data and unreadCount
    return {
      data: data.data || [],
      unreadCount: data.unreadCount || 0,
    };
  }

  async markNotificationAsRead(id: string) {
    return this.request<{ success: boolean }>(`/hr/notifications/${id}/read`, {
      method: "PUT",
    });
  }

  async markAllNotificationsAsRead() {
    return this.request<{ success: boolean }>(
      "/hr/notifications/mark-all-read",
      {
        method: "PUT",
      },
    );
  }

  // --- ANALYTICS INVITES ---
  async uploadAnalyticsInvites(formId: string, file: File) {
    const formData = new FormData();
    formData.append("file", file);
    return this.request<{
      total: number;
      valid: number;
      invalid: number;
      preview: any[];
    }>(`/analytics-invites/${formId}/upload`, {
      method: "POST",
      body: formData,
      headers: {}, // Let browser set Content-Type for FormData
    });
  }

  async sendAnalyticsInvites(
    formId: string,
    invites: any[],
    channels: string[] = ["email"],
    customMessage?: string,
  ) {
    return this.request<{
      sent: number;
      failed: number;
      allSuccessful: boolean;
      details: any[];
    }>(`/analytics-invites/${formId}/send`, {
      method: "POST",
      body: JSON.stringify({ invites, channels, customMessage }),
    });
  }

  async requestAnalyticsOTP(
    formId: string,
    email: string,
    phone: string,
    channel: "email" | "sms" = "email",
  ) {
    return this.request<{
      success: boolean;
      message: string;
    }>("/analytics-invites/request-otp", {
      method: "POST",
      body: JSON.stringify({ formId, email, phone, channel }),
    });
  }

  async verifyAnalyticsOTP(formId: string, email: string | null, phone: string | null, otp: string) {
    return this.request<{
      token: string;
      email?: string;
      phone?: string;
      formId: string;
      expiresAt: string;
    }>("/analytics-invites/verify-otp", {
      method: "POST",
      body: JSON.stringify({ formId, email, phone, otp }),
    });
  }

  async deleteNotification(id: string) {
    return this.request<{ success: boolean }>(`/hr/notifications/${id}`, {
      method: "DELETE",
    });
  }

  async deleteAllNotifications() {
    return this.request<{ success: boolean }>("/hr/notifications", {
      method: "DELETE",
    });
  }

  async getPerformanceScores() {
    return this.request<Record<string, number>>("/users/performance-scores");
  }

  async submitReview(reviewData: {
    responseId: string;
    reviewerId: string;
    submitterId: string;
    reviewOption: string;
    tenantId: string;
  }) {
    console.log("Calling submitReview API with:", reviewData);

    // Use fetch directly to avoid the request() wrapper issues
    const url = `${this.baseUrl}/users/reviews`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      "X-App-Type": "website",
    };

    if (this.token) {
      headers.Authorization = `Bearer ${this.token}`;
    }

    const response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(reviewData),
    });

    const data = await response.json();
    console.log("submitReview raw response:", {
      status: response.status,
      data,
    });

    if (!response.ok) {
      throw new Error(data.message || "Failed to submit review");
    }

    // Return in the format your component expects
    return {
      success: data.success,
      message: data.message,
      data: data.data,
    };
  }

  async getReviewsForResponse(responseId: string) {
    // Backend returns { success, reviews } without a 'data' key so use raw fetch
    // Now a public endpoint, no authentication required
    const url = `${this.baseUrl}/responses/reviews/${responseId}`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      "X-App-Type": "website",
    };
    // Reviews are now public, so no Authorization header needed
    const res = await fetch(url, { headers });
    if (!res.ok) throw new Error(`Failed to fetch reviews: ${res.status}`);
    return res.json() as Promise<{
      success: boolean;
      reviews: Array<{
        id: string;
        reviewer: { id: string; name: string; email: string };
        option: string;
        scoreChange: number;
        createdAt: string;
      }>;
    }>;
  }

}


// Create and export a singleton instance
export const apiClient = new ApiClient();
export { ApiError };
