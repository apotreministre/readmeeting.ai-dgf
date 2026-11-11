/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Ministère des Finances et du Budget - République de Côte d'Ivoire
 *
 * Types et interfaces pour le workflow natif Microsoft Entra ID :
 * « Admin Consent Request Workflow »
 */

export type AdminConsentWorkflowStatus =
  | 'admin_consent_required'    // 1. Autorisation administrateur requise
  | 'request_submitted'         // 2. Demande d'approbation soumise à Microsoft
  | 'decision_pending'          // 3. En attente de décision administrative
  | 'consent_granted'           // 4. Autorisation accordée
  | 'request_declined'          // 5. Demande refusée
  | 'workflow_unavailable'      // 6. Workflow administrateur indisponible / désactivé dans le tenant
  | 'verification_needed';      // 7. Vérification technique nécessaire

export interface AdminConsentWorkflowState {
  status: AdminConsentWorkflowStatus;
  lastActionDate: string;
  requestedByUpn?: string;
  requestedByName?: string;
  requestedScopes: string[];
  justificationNotes?: string;
  correlationId?: string;
  lastVerificationDate?: string;
  lastVerificationResult?: {
    success: boolean;
    message: string;
    details?: string;
    hasOnlineMeetingTranscriptScope?: boolean;
  };
  lastErrorCode?: string;
  lastErrorMessage?: string;
  submittedToMicrosoftAt?: string;
}

export interface EntraAdminPrerequisiteConfig {
  tenantId: string;
  applicationName: string;
  targetApi: string;
  permissionRequested: string;
  permissionType: 'Delegated' | 'Application';
  existingPermissions: string[];
  adminConsentSettingsPath: string;
  requiredTenantSettings: {
    setting: string;
    recommendedValue: string;
    description: string;
  }[];
}

export const DGF_ENTRA_PREREQUISITES: EntraAdminPrerequisiteConfig = {
  tenantId: '9ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a',
  applicationName: 'ReadMeeting.ai | DGF',
  targetApi: 'Microsoft Graph',
  permissionRequested: 'OnlineMeetingTranscript.Read.All',
  permissionType: 'Delegated',
  existingPermissions: ['User.Read', 'Calendars.Read', 'OnlineMeetings.Read'],
  adminConsentSettingsPath: 'Microsoft Entra ID > Applications d\'entreprise > Consentement et autorisations > Paramètres de consentement administrateur',
  requiredTenantSettings: [
    {
      setting: 'Les utilisateurs peuvent demander le consentement d\'un administrateur aux applications auxquelles ils ne peuvent pas consentir',
      recommendedValue: 'Oui (Yes)',
      description: 'Active l\'affichage du formulaire natif de justification Microsoft lorsque l\'utilisateur clique sur Demander les autorisations.',
    },
    {
      setting: 'Qui peut examiner les demandes de consentement administrateur',
      recommendedValue: 'Administrateurs généraux ou Réviseurs désignés DGF',
      description: 'Définit les destinataires habilités à recevoir le courriel officiel Microsoft Security et valider la demande.',
    },
    {
      setting: 'Les utilisateurs sélectionnés recevront des notifications par e-mail pour les demandes',
      recommendedValue: 'Oui (Yes)',
      description: 'Déclenche l\'expédition automatique du courriel officiel Microsoft Security lors de chaque soumission.',
    },
    {
      setting: 'Les utilisateurs sélectionnés recevront des rappels d\'expiration des demandes',
      recommendedValue: 'Oui (Yes)',
      description: 'Rappels réguliers envoyés par Microsoft jusqu\'à traitement de la demande.',
    },
    {
      setting: 'Durée de validité de la demande de consentement (jours)',
      recommendedValue: '30 jours',
      description: 'Période durant laquelle la demande reste en file d\'attente d\'approbation sur le portail Entra.',
    },
  ],
};

export const ADMIN_CONSENT_WORKFLOW_META: Record<
  AdminConsentWorkflowStatus,
  {
    stepNumber: number;
    label: string;
    shortLabel: string;
    description: string;
    badgeClass: string;
    indicatorColor: string;
    actionAdvice: string;
  }
> = {
  admin_consent_required: {
    stepNumber: 1,
    label: 'Autorisation administrateur requise',
    shortLabel: 'Consentement requis',
    description: 'La récupération des transcriptions Teams nécessite la permission déléguée OnlineMeetingTranscript.Read.All avec approbation administrative du locataire DGF.',
    badgeClass: 'bg-amber-950/80 text-amber-300 border-amber-800/80',
    indicatorColor: 'text-amber-400',
    actionAdvice: 'Cliquez sur « Demander les autorisations administrateur » pour déclencher le formulaire officiel Microsoft.',
  },
  request_submitted: {
    stepNumber: 2,
    label: 'Demande d\'approbation soumise à Microsoft',
    shortLabel: 'Demande transmise',
    description: 'La demande a été formulée via l\'interface Microsoft Entra ID. Le système Microsoft a pris en compte la sollicitation.',
    badgeClass: 'bg-blue-950/80 text-blue-300 border-blue-800/80',
    indicatorColor: 'text-blue-400',
    actionAdvice: 'Le workflow Microsoft achemine automatiquement la notification aux réviseurs de la DGF.',
  },
  decision_pending: {
    stepNumber: 3,
    label: 'En attente de décision administrative',
    shortLabel: 'Décision en attente',
    description: 'La demande est enregistrée dans la file d\'attente d\'approbation Entra ID de la DGF. L\'administrateur doit statuer dans le portail.',
    badgeClass: 'bg-purple-950/80 text-purple-300 border-purple-800/80',
    indicatorColor: 'text-purple-400',
    actionAdvice: 'Dès que l\'administrateur valide le courriel Microsoft Security, cliquez sur « Vérifier l\'état du consentement ».',
  },
  consent_granted: {
    stepNumber: 4,
    label: 'Autorisation accordée',
    shortLabel: 'Consentement accordé',
    description: 'Le consentement administrateur a été validé sur Microsoft Entra ID. Le jeton d\'accès intègre la permission OnlineMeetingTranscript.Read.All.',
    badgeClass: 'bg-emerald-950/80 text-emerald-300 border-emerald-800/80',
    indicatorColor: 'text-emerald-400',
    actionAdvice: 'Les transcriptions Microsoft Teams peuvent désormais être récupérées et traitées.',
  },
  request_declined: {
    stepNumber: 5,
    label: 'Demande refusée',
    shortLabel: 'Consentement refusé',
    description: 'L\'administrateur Microsoft 365 a rejeté la demande d\'autorisation dans le portail Entra ID ou la politique du locataire s\'y oppose.',
    badgeClass: 'bg-rose-950/80 text-rose-300 border-rose-800/80',
    indicatorColor: 'text-rose-400',
    actionAdvice: 'Contactez la Direction des Systèmes d\'Information ou votre référent sécurité pour connaître les motifs du rejet.',
  },
  workflow_unavailable: {
    stepNumber: 6,
    label: 'Workflow administrateur indisponible',
    shortLabel: 'Workflow inactif',
    description: 'Le paramètre « Admin consent request workflow » n\'est pas activé dans le locataire Entra ID de la DGF. Microsoft bloque la soumission automatique.',
    badgeClass: 'bg-orange-950/80 text-orange-300 border-orange-800/80',
    indicatorColor: 'text-orange-400',
    actionAdvice: 'L\'administrateur doit activer l\'option dans : Applications d\'entreprise > Consentement et autorisations > Paramètres de consentement administrateur.',
  },
  verification_needed: {
    stepNumber: 7,
    label: 'Vérification technique nécessaire',
    shortLabel: 'Vérification requise',
    description: 'Un contrôle de validité du jeton d\'accès et de la portée réelle Microsoft Graph est nécessaire pour confirmer l\'état opérationnel.',
    badgeClass: 'bg-cyan-950/80 text-cyan-300 border-cyan-800/80',
    indicatorColor: 'text-cyan-400',
    actionAdvice: 'Cliquez sur « Vérifier l\'état du consentement » pour interroger silencieusement Microsoft Entra ID.',
  },
};
