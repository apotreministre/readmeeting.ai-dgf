export interface MsalRuntimeConfig {
  clientId: string;
  tenantId: string;
  redirectUri: string;
  scopes: string[];
  isConfigured: boolean;
}

export interface DgfUserProfile {
  id?: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  officeLocation?: string;
  department?: string;
}

export interface DiagnosticCheckItem {
  id: string;
  title: string;
  description: string;
  status: 'success' | 'warning' | 'error' | 'pending';
  details?: string;
  recommendation?: string;
}

export interface DiagnosticReport {
  timestamp: string;
  currentOrigin: string;
  config: MsalRuntimeConfig;
  checks: DiagnosticCheckItem[];
  lastError?: {
    code: string;
    message: string;
    solution: string;
  };
}
