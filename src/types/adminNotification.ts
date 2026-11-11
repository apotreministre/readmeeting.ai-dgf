/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Types et interfaces pour la notification institutionnelle d'autorisations Microsoft 365
 * et le suivi des configurations Entra ID / Microsoft Teams.
 */

export type AdminAuthStatus =
  | 'pending'                  // Autorisations en attente
  | 'entra_configured'         // Configuration Entra ID effectuée
  | 'teams_configured'         // Configuration Microsoft Teams effectuée
  | 'verification_pending'     // Vérification technique en attente
  | 'fully_configured';        // Configuration complète

export interface AdminNotificationConfig {
  status: AdminAuthStatus;
  lastUpdated: string;
  updatedByUpn?: string;
  notes?: string;
  isBannerCollapsed?: boolean;
}

export const ADMIN_AUTH_STATUS_CONFIG: Record<
  AdminAuthStatus,
  {
    label: string;
    description: string;
    badgeClass: string;
    iconColor: string;
  }
> = {
  pending: {
    label: 'Autorisations en attente',
    description: 'Aucune intervention administrative n a encore été enregistrée sur le locataire DGF.',
    badgeClass: 'bg-amber-950/80 text-amber-300 border-amber-800/80',
    iconColor: 'text-amber-400',
  },
  entra_configured: {
    label: 'Configuration Entra ID effectuée',
    description: 'Permission déléguée ajoutée et consentement administrateur accordé sur Microsoft Entra ID.',
    badgeClass: 'bg-blue-950/80 text-blue-300 border-blue-800/80',
    iconColor: 'text-blue-400',
  },
  teams_configured: {
    label: 'Configuration Microsoft Teams effectuée',
    description: 'Paramètres EnableGraphTranscriptAccess et EnableAttributedTranscripts validés dans Teams Admin Center.',
    badgeClass: 'bg-indigo-950/80 text-indigo-300 border-indigo-800/80',
    iconColor: 'text-indigo-400',
  },
  verification_pending: {
    label: 'Vérification technique en attente',
    description: 'Les déclarations administratives sont enregistrées ; un test de requête Graph est requis.',
    badgeClass: 'bg-purple-950/80 text-purple-300 border-purple-800/80',
    iconColor: 'text-purple-400',
  },
  fully_configured: {
    label: 'Configuration complète',
    description: 'Les autorisations Entra ID et les stratégies Teams sont confirmées opérationnelles.',
    badgeClass: 'bg-emerald-950/80 text-emerald-300 border-emerald-800/80',
    iconColor: 'text-emerald-400',
  },
};

export const OFFICIAL_ADMIN_REQUEST_TEMPLATE = `OBJET : Demande d'activation des autorisations Microsoft 365 — ReadMeeting.ai | Direction Générale des Financements (DGF)

Madame, Monsieur l'Administrateur Microsoft 365,

Dans le cadre du déploiement de la solution institutionnelle ReadMeeting.ai pour la Direction Générale des Financements (Ministère des Finances et du Budget de Côte d'Ivoire), nous sollicitons votre intervention afin de finaliser les autorisations techniques requises pour la récupération des transcriptions des réunions Microsoft Teams.

1. CONFIGURATION MICROSOFT ENTRA ID :
- Application enregistrée : ReadMeeting.ai | DGF
- API cible : Microsoft Graph
- Type d'autorisation : Déléguée
- Permission à ajouter : OnlineMeetingTranscript.Read.All
- Action attendue :
  a) Ajouter la permission déléguée OnlineMeetingTranscript.Read.All à l'inscription d'application.
  b) Cliquer impérativement sur « Accorder un consentement d'administrateur pour la DGF ».
- Portail d'administration : https://entra.microsoft.com/

2. CONFIGURATION MICROSOFT TEAMS ADMIN CENTER :
Dans le Centre d'administration Microsoft Teams (Meetings > Meeting settings > Transcript API access), ou via PowerShell (Set-CsTeamsMeetingConfiguration) :
- EnableGraphTranscriptAccess : Autoriser l'accès aux transcriptions via Microsoft Graph ($true).
- EnableAttributedTranscripts : Autoriser l'attribution des interventions aux locuteurs ($true).
- Portail d'administration : https://admin.teams.microsoft.com/

Ces configurations respectent les politiques de gouvernance et de sécurité des données du Ministère des Finances et du Budget.

Restant à votre disposition pour toute précision technique.

L'équipe projet ReadMeeting.ai | DGF`;
