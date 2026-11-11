import React, { useState, useEffect } from 'react';
import {
  Calendar,
  Clock,
  Users,
  Video,
  FileText,
  Filter,
  CheckCircle2,
  AlertCircle,
  Sparkles,
  ArrowRight,
  RefreshCw,
  Building2,
  ExternalLink,
  MapPin,
  Mail,
  Briefcase,
  Info,
  Check,
  CheckCheck,
  XCircle,
  HelpCircle,
  Lock,
  ShieldAlert,
  ChevronDown,
  ChevronUp,
  FileCheck,
  AlertTriangle,
  FileCode2,
  Layers,
  Sparkle
} from 'lucide-react';
import { MeetingItem, MeetingAttendee, UserMeetingRole, TranscriptAccessibilityStatus, MeetingProcessingStatus } from '../types/meeting';
import { DgfUserProfile } from '../types/auth';
import {
  getMeetingTemporalStatus,
  checkMeetingTranscriptStatus,
  detectUserMeetingRole,
  isOrganizerExternal
} from '../services/transcriptService';
import {
  TranscriptCheckResult,
  TranscriptStatus,
  TranscriptCue,
  TranscriptRetrievalAudit
} from '../types/transcript';
import { AdminPermissionNotificationBanner } from './AdminPermissionNotificationBanner';
import { AdminNotificationConfig } from '../types/adminNotification';
import { AdminConsentWorkflowState } from '../types/adminConsentWorkflow';
import {
  NativeConsentExecutionResult,
  ConsentVerificationResult,
} from '../services/adminConsentService';
import { TranscriptViewerModal } from './TranscriptViewerModal';

interface DashboardViewProps {
  user: DgfUserProfile | null;
  meetings: MeetingItem[];
  isLoadingMeetings: boolean;
  onRefreshMeetings: () => Promise<void>;
  calendarPeriod?: 'next_7_days' | 'past_30_days' | 'past_90_days';
  onChangeCalendarPeriod?: (period: 'next_7_days' | 'past_30_days' | 'past_90_days') => Promise<void>;
  onNavigateToAuth: () => void;
  onSelectMeetingForReport: (meeting: MeetingItem) => void;
  onCheckTranscript?: (meeting: MeetingItem, forceRefresh?: boolean) => Promise<TranscriptCheckResult>;
  adminConfig: AdminNotificationConfig;
  onUpdateAdminConfig: (newConfig: Partial<AdminNotificationConfig>) => void;
  onRunTechnicalCheck?: () => Promise<{ success: boolean; message: string; details?: string }>;
  isCheckingTechnical?: boolean;
  workflowState?: AdminConsentWorkflowState;
  onRequestConsent?: (justification?: string) => Promise<NativeConsentExecutionResult>;
  onVerifyConsent?: () => Promise<ConsentVerificationResult>;
  onUpdateWorkflowState?: (partial: Partial<AdminConsentWorkflowState>) => void;
  isRequestingConsent?: boolean;
  isVerifyingConsent?: boolean;
}

export const DashboardView: React.FC<DashboardViewProps> = ({
  user,
  meetings,
  isLoadingMeetings,
  onRefreshMeetings,
  calendarPeriod = 'next_7_days',
  onChangeCalendarPeriod,
  onNavigateToAuth,
  onSelectMeetingForReport,
  onCheckTranscript,
  adminConfig,
  onUpdateAdminConfig,
  onRunTechnicalCheck,
  isCheckingTechnical = false,
  workflowState,
  onRequestConsent,
  onVerifyConsent,
  onUpdateWorkflowState,
  isRequestingConsent = false,
  isVerifyingConsent = false,
}) => {
  const [filterType, setFilterType] = useState<'all' | 'teams' | 'in_person' | 'completed_teams'>('all');
  const [filterScenario, setFilterScenario] = useState<'all' | 'organized_by_me' | 'participated_by_me'>('all');
  const [selectedMeeting, setSelectedMeeting] = useState<MeetingItem | null>(null);
  const [transcriptCheck, setTranscriptCheck] = useState<TranscriptCheckResult | null>(null);
  const [isCheckingTranscript, setIsCheckingTranscript] = useState<boolean>(false);
  const [adminGuideOpen, setAdminGuideOpen] = useState<boolean>(false);
  const [pipelinePreviewOpen, setPipelinePreviewOpen] = useState<boolean>(false);
  const [isViewerModalOpen, setIsViewerModalOpen] = useState<boolean>(false);

  // Synchronisation de l'état de transcription dès l'ouverture d'une réunion
  useEffect(() => {
    if (selectedMeeting) {
      const temporal = getMeetingTemporalStatus(selectedMeeting.start, selectedMeeting.end);

      if (!selectedMeeting.isOnlineMeeting) {
        setTranscriptCheck({
          meetingId: selectedMeeting.id,
          meetingTitle: selectedMeeting.title,
          temporalStatus: temporal,
          status: 'transcript_unavailable',
          checkedAt: new Date().toISOString(),
          error: {
            code: 'NOT_A_TEAMS_MEETING',
            title: 'Séance en présentiel',
            message: "Cette réunion ne comporte aucun lien Microsoft Teams.",
            actionRequired: "Utilisez les modalités présentielles (compte rendu manuel ou audio de salle).",
          },
        });
      } else if (temporal === 'scheduled') {
        setTranscriptCheck({
          meetingId: selectedMeeting.id,
          meetingTitle: selectedMeeting.title,
          temporalStatus: 'scheduled',
          status: 'scheduled',
          checkedAt: new Date().toISOString(),
          error: {
            code: 'MEETING_SCHEDULED',
            title: 'Réunion programmée',
            message: "Cette séance Teams est planifiée dans le futur. La transcription ne peut être consultée qu'à l'issue de la tenue effective de la réunion.",
            actionRequired: "Attendez la tenue de la réunion pour vérifier l'enregistrement et la transcription.",
          },
        });
      } else if (temporal === 'in_progress') {
        setTranscriptCheck({
          meetingId: selectedMeeting.id,
          meetingTitle: selectedMeeting.title,
          temporalStatus: 'in_progress',
          status: 'in_progress',
          checkedAt: new Date().toISOString(),
          error: {
            code: 'MEETING_IN_PROGRESS',
            title: 'Réunion en cours de séance',
            message: "La séance est active. Si l'organisateur a lancé l'enregistrement, la transcription en direct est en cours de capture par Teams.",
            actionRequired: "Attendez la clôture de la réunion pour consolider le procès-verbal.",
          },
        });
      } else {
        // Completed : statut initial en attente de vérification Graph explicite
        setTranscriptCheck({
          meetingId: selectedMeeting.id,
          meetingTitle: selectedMeeting.title,
          temporalStatus: 'completed',
          status: 'completed',
          checkedAt: new Date().toISOString(),
        });
      }
    } else {
      setTranscriptCheck(null);
      setAdminGuideOpen(false);
      setPipelinePreviewOpen(false);
    }
  }, [selectedMeeting]);

  const handleTriggerCheck = async (force: boolean = false) => {
    if (!selectedMeeting || !onCheckTranscript) return;
    setIsCheckingTranscript(true);
    try {
      const res = await onCheckTranscript(selectedMeeting, force);
      setTranscriptCheck(res);
    } catch (err: any) {
      setTranscriptCheck({
        meetingId: selectedMeeting.id,
        meetingTitle: selectedMeeting.title,
        temporalStatus: getMeetingTemporalStatus(selectedMeeting.start, selectedMeeting.end),
        status: 'transient_error',
        checkedAt: new Date().toISOString(),
        error: {
          code: 'EXECUTION_ERROR',
          title: 'Erreur inattendue',
          message: err.message || String(err),
          actionRequired: 'Réessayez ultérieurement.',
        },
      });
    } finally {
      setIsCheckingTranscript(false);
    }
  };

  // Événements types DGF affichés en mode démonstration (illustrant les Scénarios A et B)
  const demoFallbackMeetings: MeetingItem[] = [
    // SCÉNARIO A : Réunion organisée par l'utilisateur connecté (Kouassi Jean-Marc / DGF)
    {
      id: 'dgf-demo-historical-teams',
      title: 'Comité de Clôture des Négociations Techniques — Prêt BAD & BOAD (Financements Énergétiques)',
      subject: 'Revue et arrêté définitif du procès-verbal des négociations',
      bodyPreview: 'Séance de clôture : Arbitrage des clauses financières, validation du calendrier de décaissement et arrêté officiel du procès-verbal avec les représentants des bailleurs.',
      organizer: {
        name: user?.displayName || 'Kouassi Jean-Marc',
        email: user?.mail || user?.userPrincipalName || 'jm.kouassi@finances.gouv.ci',
        isExternal: false,
      },
      start: new Date(Date.now() - 3600000 * 28).toISOString(),
      end: new Date(Date.now() - 3600000 * 26).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Microsoft Teams',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_historical_dgf_bad%40thread.v2/0?context=%7b%22Tid%22%3a%229ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a%22%2c%22Oid%22%3a%22user-dgf-organizer%22%7d',
      status: 'completed',
      captureMethod: 'teams_native',
      isDgfMeeting: true,
      departmentCode: 'DFEE (Financements Extérieurs)',
      userRole: 'organizer',
      isExternalOrganizer: false,
      transcriptAccessibility: 'accessible',
      processingStatus: 'ready_for_processing',
      attendees: [
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || 'jm.kouassi@finances.gouv.ci', type: 'required', role: 'organizer', responseStatus: 'accepted' },
        { name: 'Directeur Général des Financements', email: 'dg@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
        { name: 'Dr. Salimata Touré (Représentante BAD)', email: 's.toure@afdb.org', type: 'required', isExternal: true, responseStatus: 'accepted' },
        { name: 'Conseiller Juridique DGF', email: 'juridique@finances.gouv.ci', type: 'optional', responseStatus: 'accepted' },
      ],
      hasTranscript: true,
      hasOfficialReport: false,
    },
    // SCÉNARIO B.1 : Participant invité — Transcription autorisée et partagée avec tous les participants
    {
      id: 'dgf-demo-attendee-authorized',
      title: 'Revue Trimestrielle de Performance du Portefeuille des Projets DGF (Direction Générale)',
      subject: 'Examen des indicateurs de décaissement et arbitrages financiers',
      bodyPreview: 'Session plénière présidée par le Directeur Général. Bilan consolidé des décaissements trimestriels, arbitrage des points de blocage et validation des orientations opérationnelles. Transcription Teams ouverte à l ensemble des participants.',
      organizer: {
        name: 'Directeur Général des Financements',
        email: 'dg@finances.gouv.ci',
        department: 'Direction Générale',
        isExternal: false,
      },
      start: new Date(Date.now() - 3600000 * 52).toISOString(),
      end: new Date(Date.now() - 3600000 * 50).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Microsoft Teams',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_dgf_dg_performance%40thread.v2/0?context=%7b%22Tid%22%3a%229ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a%22%2c%22Oid%22%3a%22dg-dgf-organizer%22%7d',
      status: 'completed',
      captureMethod: 'teams_native',
      isDgfMeeting: true,
      departmentCode: 'Direction Générale (DGF)',
      userRole: 'attendee',
      isExternalOrganizer: false,
      transcriptAccessibility: 'accessible',
      processingStatus: 'ready_for_processing',
      attendees: [
        { name: 'Directeur Général des Financements', email: 'dg@finances.gouv.ci', type: 'required', role: 'organizer', responseStatus: 'accepted' },
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || user?.userPrincipalName || 'jm.kouassi@finances.gouv.ci', type: 'required', role: 'attendee', responseStatus: 'accepted' },
        { name: 'Touré Ibrahim (Directeur Dette)', email: 'i.toure@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
        { name: 'Bamba Mamadou (Directeur Adjoint)', email: 'm.bamba@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
      ],
      hasTranscript: true,
      hasOfficialReport: false,
    },
    // SCÉNARIO B.2 : Participant invité — Transcription restreinte par les paramètres de partage de l'organisateur (Organizers only)
    {
      id: 'dgf-demo-attendee-restricted',
      title: 'Comité d Arbitrage Budgétaire Interministériel Restreint — Financements d Urgence',
      subject: 'Arbitrage préalable et confidentiel des dotations de trésorerie',
      bodyPreview: 'Séance préparatoire restreinte. Paramètres de partage Teams verrouillés sur « Organisateurs et coorganisateurs ». L accès de transcription pour les participants invités requiert une levée de restriction par l organisateur.',
      organizer: {
        name: 'Touré Ibrahim',
        email: 'i.toure@finances.gouv.ci',
        department: 'DPE (Dette & Emprunts)',
        isExternal: false,
      },
      start: new Date(Date.now() - 3600000 * 76).toISOString(),
      end: new Date(Date.now() - 3600000 * 74).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Microsoft Teams',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_dgf_arbitrage_urgent%40thread.v2/0?context=%7b%22Tid%22%3a%229ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a%22%2c%22Oid%22%3a%22toure-dpe-organizer%22%7d',
      status: 'completed',
      captureMethod: 'teams_native',
      isDgfMeeting: true,
      departmentCode: 'DPE (Dette & Emprunts)',
      userRole: 'attendee',
      isExternalOrganizer: false,
      transcriptAccessibility: 'restricted_by_sharing',
      processingStatus: 'idle',
      attendees: [
        { name: 'Touré Ibrahim', email: 'i.toure@finances.gouv.ci', type: 'required', role: 'organizer', responseStatus: 'accepted' },
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || user?.userPrincipalName || 'jm.kouassi@finances.gouv.ci', type: 'required', role: 'attendee', responseStatus: 'accepted' },
        { name: 'Conseiller Spécial du Ministre', email: 'conseiller.special@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
      ],
      hasTranscript: true,
      hasOfficialReport: false,
    },
    // SCÉNARIO B.3 : Participant invité — Réunion sans transcription enregistrée (200 OK collection vide)
    {
      id: 'dgf-demo-attendee-notranscript',
      title: 'Point d Étape Technique — Intégration du Système Intégré de Gestion de la Dette (SIGD)',
      subject: 'Synchronisation des flux applicatifs et base de données centrale',
      bodyPreview: 'Session de synchronisation technique interne. Réunion tenue sur Teams sans activation de la commande « Démarrer la transcription ». Aucune transcription textuelle enregistrée.',
      organizer: {
        name: 'Bamba Mamadou',
        email: 'm.bamba@finances.gouv.ci',
        department: 'DSI / DGF',
        isExternal: false,
      },
      start: new Date(Date.now() - 3600000 * 100).toISOString(),
      end: new Date(Date.now() - 3600000 * 98).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Microsoft Teams',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_dgf_sigd_technique%40thread.v2/0?context=%7b%22Tid%22%3a%229ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a%22%2c%22Oid%22%3a%22bamba-dsi-organizer%22%7d',
      status: 'completed',
      captureMethod: 'teams_native',
      isDgfMeeting: true,
      departmentCode: 'DSI (Systèmes d Information)',
      userRole: 'attendee',
      isExternalOrganizer: false,
      transcriptAccessibility: 'no_transcript',
      processingStatus: 'idle',
      attendees: [
        { name: 'Bamba Mamadou', email: 'm.bamba@finances.gouv.ci', type: 'required', role: 'organizer', responseStatus: 'accepted' },
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || user?.userPrincipalName || 'jm.kouassi@finances.gouv.ci', type: 'required', role: 'attendee', responseStatus: 'accepted' },
        { name: 'Ingénieur Système SIGD', email: 'support.sigd@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
      ],
      hasTranscript: false,
      hasOfficialReport: false,
    },
    // SCÉNARIO B.4 : Participant invité — Réunion convoquée par un locataire externe (Bailleurs / BAD)
    {
      id: 'dgf-demo-attendee-external',
      title: 'Table Ronde Multilatérale des Bailleurs de Fonds (BAD, BM, AFD)',
      subject: 'Harmonisation des critères fiduciaires et des audits environnementaux',
      bodyPreview: 'Session multilatérale convoquée par la Représentation Régionale de la Banque Africaine de Développement (BAD). Réunion hébergée sur le tenant externe afdb.org.',
      organizer: {
        name: 'Dr. Salimata Touré',
        email: 's.toure@afdb.org',
        department: 'Division des Opérations Souveraines (BAD)',
        isExternal: true,
      },
      start: new Date(Date.now() - 3600000 * 124).toISOString(),
      end: new Date(Date.now() - 3600000 * 121).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Microsoft Teams (Tenant BAD / afdb.org)',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_afdb_multilateral%40thread.v2/0?context=%7b%22Tid%22%3a%22external-afdb-tenant-id%22%2c%22Oid%22%3a%22dr-toure-afdb%22%7d',
      status: 'completed',
      captureMethod: 'teams_native',
      isDgfMeeting: false,
      departmentCode: 'Coopération Multilatérale',
      userRole: 'attendee',
      isExternalOrganizer: true,
      organizerTenantId: 'external-afdb-tenant-id',
      transcriptAccessibility: 'external_tenant',
      processingStatus: 'idle',
      attendees: [
        { name: 'Dr. Salimata Touré', email: 's.toure@afdb.org', type: 'required', role: 'organizer', isExternal: true, responseStatus: 'accepted' },
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || user?.userPrincipalName || 'jm.kouassi@finances.gouv.ci', type: 'required', role: 'attendee', responseStatus: 'accepted' },
        { name: 'Chef de Projet AFD', email: 'projet.ci@afd.fr', type: 'required', isExternal: true, responseStatus: 'accepted' },
      ],
      hasTranscript: true,
      hasOfficialReport: false,
    },
    // Réunions futures d'agenda
    {
      id: 'dgf-demo-1',
      title: 'Comité de Cadrage des Financements Extérieurs (BAD & Banque Mondiale)',
      subject: 'Revue de portefeuille des projets d investissements structurants',
      bodyPreview: 'Ordre du jour : 1. Examen des conventions bilatérales. 2. État d avancement des décaissements trimestriels. 3. Divers.',
      organizer: {
        name: user?.displayName || 'Kouassi Jean-Marc',
        email: user?.mail || user?.userPrincipalName || 'jm.kouassi@finances.gouv.ci',
      },
      start: new Date(Date.now() + 3600000 * 2).toISOString(),
      end: new Date(Date.now() + 3600000 * 4).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Tour C, 14ème étage - Salle Auditorium',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/demo-meeting-1',
      status: 'scheduled',
      captureMethod: 'teams_native',
      isDgfMeeting: true,
      departmentCode: 'DFEE (Financements Extérieurs)',
      userRole: 'organizer',
      isExternalOrganizer: false,
      transcriptAccessibility: 'pending_check',
      processingStatus: 'idle',
      attendees: [
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || 'jm.kouassi@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
        { name: 'Directeur Général des Financements', email: 'dg@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
        { name: 'Bamba Salimata', email: 's.bamba@worldbank.org', type: 'optional', isExternal: true, responseStatus: 'tentativelyAccepted' },
      ],
      hasTranscript: false,
      hasOfficialReport: false,
    },
    {
      id: 'dgf-demo-2',
      title: 'Comité Hebdomadaire de Trésorerie et Arbitrage des Titres Publics',
      subject: 'Arbitrage sur les émissions obligataires du Trésor Public',
      bodyPreview: 'Revue des souscriptions primaires et prévisions de liquidité à 14 jours.',
      organizer: {
        name: 'Directeur Général des Financements',
        email: 'dg@finances.gouv.ci',
      },
      start: new Date(Date.now() + 3600000 * 26).toISOString(),
      end: new Date(Date.now() + 3600000 * 28).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Salle Conseil - 12ème étage, Plateau',
      isOnlineMeeting: false,
      status: 'scheduled',
      captureMethod: 'meeting_bot',
      isDgfMeeting: true,
      departmentCode: 'DPE (Dette & Emprunts)',
      userRole: 'attendee',
      isExternalOrganizer: false,
      transcriptAccessibility: 'no_transcript',
      processingStatus: 'idle',
      attendees: [
        { name: 'Directeur Général', email: 'dg@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || 'jm.kouassi@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
      ],
      hasTranscript: false,
      hasOfficialReport: false,
    },
    {
      id: 'dgf-demo-3',
      title: 'Séance de Négociation Technique — Financement Projets Solaires (AFD)',
      subject: 'Revue des clauses suspensives du prêt concessionnel',
      bodyPreview: 'Discussion technique sur les garanties souveraines et les jalons de réalisation.',
      organizer: {
        name: 'Touré Ibrahim',
        email: 'i.toure@finances.gouv.ci',
      },
      start: new Date(Date.now() + 3600000 * 50).toISOString(),
      end: new Date(Date.now() + 3600000 * 52).toISOString(),
      startTimeZone: 'UTC',
      endTimeZone: 'UTC',
      location: 'Microsoft Teams',
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/demo-meeting-3',
      status: 'scheduled',
      captureMethod: 'teams_native',
      isDgfMeeting: true,
      departmentCode: 'DFEE (Financements Extérieurs)',
      userRole: 'attendee',
      isExternalOrganizer: false,
      transcriptAccessibility: 'pending_check',
      processingStatus: 'idle',
      attendees: [
        { name: 'Touré Ibrahim', email: 'i.toure@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
        { name: user?.displayName || 'Kouassi Jean-Marc', email: user?.mail || 'jm.kouassi@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
      ],
      hasTranscript: false,
      hasOfficialReport: false,
    },
  ];

  // Si l'utilisateur est connecté, on affiche les réunions réelles Microsoft Graph
  const activeMeetings = user ? meetings : demoFallbackMeetings;

  // Calcul du rôle effectif de l'utilisateur (organisateur, participant ou coorganisateur)
  const getEffectiveRole = (m: MeetingItem): UserMeetingRole => {
    if (m.userRole) return m.userRole;
    return detectUserMeetingRole(m, user);
  };

  const filteredMeetings = activeMeetings.filter((m) => {
    if (filterType === 'teams' && !m.isOnlineMeeting) return false;
    if (filterType === 'in_person' && m.isOnlineMeeting) return false;
    if (filterType === 'completed_teams') {
      const temporal = getMeetingTemporalStatus(m.start, m.end);
      if (!m.isOnlineMeeting || temporal !== 'completed') return false;
    }

    const role = getEffectiveRole(m);
    if (filterScenario === 'organized_by_me') {
      if (role !== 'organizer') return false;
    } else if (filterScenario === 'participated_by_me') {
      if (role !== 'attendee' && role !== 'coorganizer') return false;
    }

    return true;
  });

  const countOrganized = activeMeetings.filter((m) => getEffectiveRole(m) === 'organizer').length;
  const countParticipated = activeMeetings.filter((m) => {
    const role = getEffectiveRole(m);
    return role === 'attendee' || role === 'coorganizer';
  }).length;

  const formatDate = (isoString: string) => {
    try {
      const d = new Date(isoString);
      return d.toLocaleDateString('fr-FR', {
        weekday: 'short',
        day: 'numeric',
        month: 'short',
        year: 'numeric',
      });
    } catch {
      return isoString;
    }
  };

  const formatTime = (isoString: string) => {
    try {
      const d = new Date(isoString);
      return d.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
    } catch {
      return '';
    }
  };

  const getResponseStatusBadge = (status?: string) => {
    switch (status) {
      case 'accepted':
        return (
          <span className="flex items-center gap-1 text-[10px] text-emerald-400 bg-emerald-950/60 px-1.5 py-0.5 rounded border border-emerald-800/50">
            <Check className="w-2.5 h-2.5" /> Accepté
          </span>
        );
      case 'declined':
        return (
          <span className="flex items-center gap-1 text-[10px] text-rose-400 bg-rose-950/60 px-1.5 py-0.5 rounded border border-rose-800/50">
            <XCircle className="w-2.5 h-2.5" /> Refusé
          </span>
        );
      case 'tentativelyAccepted':
        return (
          <span className="flex items-center gap-1 text-[10px] text-amber-400 bg-amber-950/60 px-1.5 py-0.5 rounded border border-amber-800/50">
            <Clock className="w-2.5 h-2.5" /> Provisoire
          </span>
        );
      default:
        return (
          <span className="flex items-center gap-1 text-[10px] text-slate-400 bg-slate-900 px-1.5 py-0.5 rounded border border-slate-800">
            <HelpCircle className="w-2.5 h-2.5" /> En attente
          </span>
        );
    }
  };

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-16">
      {/* 0. Notification institutionnelle d'autorisations Microsoft 365 */}
      <AdminPermissionNotificationBanner
        user={user}
        adminConfig={adminConfig}
        onUpdateAdminConfig={onUpdateAdminConfig}
        onRunTechnicalCheck={onRunTechnicalCheck}
        isCheckingTechnical={isCheckingTechnical}
        workflowState={workflowState}
        onRequestConsent={onRequestConsent}
        onVerifyConsent={onVerifyConsent}
        onUpdateWorkflowState={onUpdateWorkflowState}
        isRequestingConsent={isRequestingConsent}
        isVerifyingConsent={isVerifyingConsent}
      />

      {/* 1. Carte de profil utilisateur connecté (Microsoft Graph /me) */}
      {user ? (
        <div className="bg-gradient-to-r from-slate-900 via-blue-950/40 to-slate-900 border border-blue-900/40 rounded-xl p-5 shadow-lg">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-start gap-4">
              <div className="w-12 h-12 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-800 flex items-center justify-center text-white font-bold text-base shadow-md border border-blue-400/30 shrink-0">
                {user.displayName.slice(0, 2).toUpperCase()}
              </div>
              <div className="space-y-1">
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-base font-bold text-white tracking-tight">{user.displayName}</h2>
                  <span className="px-2 py-0.5 text-[10px] font-semibold bg-emerald-950 text-emerald-300 border border-emerald-800 rounded flex items-center gap-1">
                    <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                    Compte Microsoft 365 vérifié
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-4 text-xs text-slate-300">
                  <span className="flex items-center gap-1 text-slate-400">
                    <Mail className="w-3.5 h-3.5 text-blue-400" />
                    {user.mail || user.userPrincipalName}
                  </span>
                  {user.jobTitle && (
                    <span className="flex items-center gap-1 text-slate-300">
                      <Briefcase className="w-3.5 h-3.5 text-amber-400" />
                      {user.jobTitle}
                    </span>
                  )}
                  {user.officeLocation && (
                    <span className="flex items-center gap-1 text-slate-400">
                      <MapPin className="w-3.5 h-3.5 text-emerald-400" />
                      {user.officeLocation}
                    </span>
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2 shrink-0">
              <button
                onClick={onRefreshMeetings}
                disabled={isLoadingMeetings}
                className="px-3.5 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-sm border border-blue-400/30 flex items-center gap-2 transition-all disabled:opacity-50"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isLoadingMeetings ? 'animate-spin' : ''}`} />
                {isLoadingMeetings ? 'Interrogation Microsoft Graph...' : 'Actualiser le calendrier'}
              </button>
            </div>
          </div>
        </div>
      ) : (
        /* Bannière d'invitation à la connexion si non authentifié */
        <div className="bg-gradient-to-r from-emerald-950/60 to-slate-900 border border-emerald-800/40 rounded-xl p-5 shadow-sm flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-start gap-3">
            <div className="w-9 h-9 rounded-lg bg-emerald-900/60 border border-emerald-700/50 flex items-center justify-center shrink-0">
              <Building2 className="w-4 h-4 text-emerald-300" />
            </div>
            <div>
              <h2 className="text-sm font-semibold text-white">
                Synchronisation de l'Agenda Institutionnel DGF
              </h2>
              <p className="text-xs text-slate-300 mt-0.5">
                Connectez-vous à votre compte professionnel Microsoft 365 pour récupérer vos réunions Outlook et vos sessions Teams réelles.
              </p>
            </div>
          </div>

          <button
            onClick={onNavigateToAuth}
            className="px-4 py-2 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-md transition-all flex items-center gap-2 shrink-0 border border-blue-400/20"
          >
            Se connecter avec Microsoft Entra ID
            <ArrowRight className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* 2. Section « Mes réunions à traiter » et Barre d'outils institutionnelle */}
      <div className="bg-slate-900/90 p-5 rounded-2xl border border-slate-800 space-y-4 shadow-md">
        {/* En-tête de section */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800/80 pb-3.5">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-400"></span>
              <h2 className="text-base sm:text-lg font-bold text-white tracking-tight">
                Mes réunions à traiter
              </h2>
              <span className="px-2 py-0.5 text-[10px] font-bold bg-blue-950 text-blue-300 border border-blue-800/60 rounded">
                Scénarios A & B (Microsoft Graph v1.0)
              </span>
            </div>
            <p className="text-xs text-slate-400">
              Traitement des transcriptions des réunions organisées par vos soins (Scénario A) ou auxquelles vous avez été formellement invité en tant que participant (Scénario B).
            </p>
          </div>

          <div className="text-xs text-slate-400 font-mono self-start sm:self-center">
            {user ? `${filteredMeetings.length} sur ${activeMeetings.length} réunion(s)` : 'Mode démonstration multi-scénarios DGF'}
          </div>
        </div>

        {/* Onglets de filtrage par Scénario (Organisées par moi vs Réunions auxquelles j'ai participé) */}
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs font-semibold text-slate-300 flex items-center gap-1.5">
            <Users className="w-3.5 h-3.5 text-blue-400" />
            Périmètre de rôle :
          </span>

          <div className="flex flex-wrap items-center gap-1.5 bg-slate-950 p-1 rounded-xl border border-slate-800">
            <button
              onClick={() => setFilterScenario('all')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                filterScenario === 'all'
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-white hover:bg-slate-900'
              }`}
            >
              Toutes les réunions ({activeMeetings.length})
            </button>

            <button
              onClick={() => setFilterScenario('organized_by_me')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1.5 ${
                filterScenario === 'organized_by_me'
                  ? 'bg-emerald-600 text-white shadow-sm'
                  : 'text-slate-300 hover:text-white hover:bg-slate-900'
              }`}
            >
              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-300" />
              Organisées par moi — Scénario A ({countOrganized})
            </button>

            <button
              onClick={() => setFilterScenario('participated_by_me')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1.5 ${
                filterScenario === 'participated_by_me'
                  ? 'bg-purple-600 text-white shadow-sm'
                  : 'text-slate-300 hover:text-white hover:bg-slate-900'
              }`}
            >
              <Users className="w-3.5 h-3.5 text-purple-300" />
              Réunions auxquelles j'ai participé — Scénario B ({countParticipated})
            </button>
          </div>
        </div>

        {/* Filtres secondaires et Sélecteur de période */}
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4 pt-2 border-t border-slate-800/60">
          {/* Types de réunions */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs font-medium text-slate-400 flex items-center gap-1">
              <Filter className="w-3 h-3 text-slate-500" />
              Type :
            </span>
            <div className="flex flex-wrap items-center gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800 text-xs">
              <button
                onClick={() => setFilterType('all')}
                className={`px-2.5 py-1 rounded transition-colors ${
                  filterType === 'all'
                    ? 'bg-slate-800 text-white font-semibold'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                Tous types
              </button>
              <button
                onClick={() => setFilterType('teams')}
                className={`px-2.5 py-1 rounded transition-colors flex items-center gap-1 ${
                  filterType === 'teams'
                    ? 'bg-indigo-600 text-white font-semibold'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                <Video className="w-3 h-3" />
                Teams
              </button>
              <button
                onClick={() => setFilterType('completed_teams')}
                className={`px-2.5 py-1 rounded transition-colors flex items-center gap-1 ${
                  filterType === 'completed_teams'
                    ? 'bg-emerald-600 text-white font-semibold'
                    : 'text-emerald-400 hover:text-white'
                }`}
                title="Afficher uniquement les réunions Teams terminées"
              >
                <Clock className="w-3 h-3 text-emerald-300" />
                Historique terminé
              </button>
              <button
                onClick={() => setFilterType('in_person')}
                className={`px-2.5 py-1 rounded transition-colors flex items-center gap-1 ${
                  filterType === 'in_person'
                    ? 'bg-slate-800 text-white font-semibold'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                <MapPin className="w-3 h-3 text-amber-300" />
                Présentiel
              </button>
            </div>
          </div>

          {/* Sélecteur de période Graph */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs text-slate-400 flex items-center gap-1">
              <Calendar className="w-3 h-3 text-blue-400" />
              Période :
            </span>
            <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800 text-xs">
              <button
                onClick={() => onChangeCalendarPeriod && onChangeCalendarPeriod('next_7_days')}
                disabled={isLoadingMeetings}
                className={`px-2.5 py-1 rounded transition-colors ${
                  calendarPeriod === 'next_7_days'
                    ? 'bg-blue-600/80 text-white font-semibold'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                7 prochains jours
              </button>
              <button
                onClick={() => onChangeCalendarPeriod && onChangeCalendarPeriod('past_30_days')}
                disabled={isLoadingMeetings}
                className={`px-2.5 py-1 rounded transition-colors ${
                  calendarPeriod === 'past_30_days'
                    ? 'bg-indigo-600/80 text-white font-semibold'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                30 derniers jours (Historique)
              </button>
              <button
                onClick={() => onChangeCalendarPeriod && onChangeCalendarPeriod('past_90_days')}
                disabled={isLoadingMeetings}
                className={`px-2.5 py-1 rounded transition-colors ${
                  calendarPeriod === 'past_90_days'
                    ? 'bg-indigo-600/80 text-white font-semibold'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                90 derniers jours
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* 3. État de chargement */}
      {isLoadingMeetings && (
        <div className="p-12 text-center bg-slate-900/60 border border-slate-800 rounded-xl space-y-3">
          <RefreshCw className="w-6 h-6 text-blue-400 animate-spin mx-auto" />
          <div className="text-sm font-semibold text-white">Interrogation de Microsoft Graph API (/me/calendarView)...</div>
          <p className="text-xs text-slate-400">Récupération des événements du calendrier Outlook en cours.</p>
        </div>
      )}

      {/* 4. État calendrier vide (utilisateur connecté sans réunion) */}
      {!isLoadingMeetings && user && filteredMeetings.length === 0 && (
        <div className="p-12 text-center bg-slate-900/60 border border-slate-800 rounded-xl space-y-4">
          <div className="w-12 h-12 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center mx-auto text-slate-400">
            <Calendar className="w-6 h-6" />
          </div>
          <div className="space-y-1">
            <h3 className="text-sm font-bold text-white">Aucune réunion ne correspond aux filtres sélectionnés</h3>
            <p className="text-xs text-slate-400 max-w-md mx-auto">
              Modifiez le filtre de rôle ou la période temporelle pour afficher d'autres événements.
            </p>
          </div>
          <button
            onClick={() => {
              setFilterScenario('all');
              setFilterType('all');
            }}
            className="px-4 py-2 text-xs font-semibold text-white bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg transition-colors inline-flex items-center gap-2"
          >
            Réinitialiser les filtres
          </button>
        </div>
      )}

      {/* 5. Liste des réunions à traiter */}
      {!isLoadingMeetings && filteredMeetings.length > 0 && (
        <div className="space-y-3.5">
          {filteredMeetings.map((meeting) => {
            const role = getEffectiveRole(meeting);
            const temporal = getMeetingTemporalStatus(meeting.start, meeting.end);
            const isAccessible =
              meeting.transcriptAccessibility === 'accessible' ||
              (meeting.isOnlineMeeting && temporal === 'completed' && role === 'organizer' && meeting.hasTranscript);

            // Libellé et classe du badge Mon Rôle
            const roleBadgeConfig =
              role === 'organizer'
                ? { label: 'Organisateur (Scénario A)', className: 'bg-emerald-950 text-emerald-300 border-emerald-800' }
                : role === 'coorganizer'
                ? { label: 'Coorganisateur (Scénario B)', className: 'bg-blue-950 text-blue-300 border-blue-800' }
                : role === 'attendee'
                ? { label: 'Participant invité (Scénario B)', className: 'bg-purple-950 text-purple-300 border-purple-800' }
                : { label: 'Non convoqué', className: 'bg-slate-900 text-slate-400 border-slate-800' };

            // Statut de disponibilité de la transcription
            const availabilityBadge =
              !meeting.isOnlineMeeting
                ? { label: 'Présentiel', className: 'text-slate-400 bg-slate-950 border-slate-800' }
                : temporal === 'scheduled'
                ? { label: 'Séance future', className: 'text-blue-400 bg-blue-950/60 border-blue-800/50' }
                : temporal === 'in_progress'
                ? { label: 'En séance active', className: 'text-amber-400 bg-amber-950/60 border-amber-800/50 animate-pulse' }
                : meeting.hasTranscript
                ? { label: 'Transcription disponible', className: 'text-emerald-300 bg-emerald-950/70 border-emerald-800' }
                : { label: 'Non transcrite', className: 'text-slate-500 bg-slate-950 border-slate-800' };

            // Statut d'accessibilité de la transcription
            const accessibilityBadge =
              !meeting.isOnlineMeeting
                ? { label: 'N/A (Présentiel)', className: 'text-slate-500 border-slate-800 bg-slate-950' }
                : meeting.transcriptAccessibility === 'accessible' || (isAccessible && meeting.hasTranscript)
                ? { label: 'Accès autorisé', className: 'text-emerald-300 bg-emerald-950 border-emerald-700/80', icon: CheckCircle2 }
                : meeting.transcriptAccessibility === 'restricted_by_sharing'
                ? { label: 'Restreint par l\'organisateur', className: 'text-rose-300 bg-rose-950 border-rose-800/80', icon: Lock }
                : meeting.transcriptAccessibility === 'external_tenant'
                ? { label: 'Locataire externe (Restreint)', className: 'text-amber-300 bg-amber-950 border-amber-800/80', icon: ShieldAlert }
                : meeting.transcriptAccessibility === 'no_transcript' || (!meeting.hasTranscript && temporal === 'completed')
                ? { label: 'Aucune transcription', className: 'text-slate-400 bg-slate-950 border-slate-800', icon: XCircle }
                : meeting.transcriptAccessibility === 'permission_missing'
                ? { label: 'Permission requise', className: 'text-rose-300 bg-rose-950 border-rose-800/80', icon: Lock }
                : { label: 'Vérification requise', className: 'text-slate-400 bg-slate-950 border-slate-800', icon: HelpCircle };

            // Statut de traitement
            const processingBadge =
              meeting.processingStatus === 'processed'
                ? { label: 'Traité', className: 'text-blue-300 bg-blue-950 border-blue-800' }
                : isAccessible
                ? { label: 'Prêt à traiter', className: 'text-emerald-300 bg-emerald-950/80 border-emerald-800' }
                : { label: 'En attente', className: 'text-slate-400 bg-slate-950 border-slate-800' };

            const AccessibilityIcon = accessibilityBadge.icon;

            return (
              <div
                key={meeting.id}
                className="bg-slate-900/90 border border-slate-800 hover:border-slate-700 rounded-xl p-5 transition-all shadow-sm flex flex-col xl:flex-row xl:items-center justify-between gap-5 group"
              >
                {/* Colonne gauche : Rôle, Titre, Métadonnées et Badges de statut */}
                <div className="space-y-2.5 flex-1 min-w-0">
                  {/* Ligne 1 : Badges de cadrage institutionnel */}
                  <div className="flex flex-wrap items-center gap-2">
                    {/* Badge Mon Rôle (Scénario A vs Scénario B) */}
                    <span className={`px-2.5 py-0.5 text-[10px] font-bold rounded border flex items-center gap-1 shadow-sm ${roleBadgeConfig.className}`}>
                      <Users className="w-3 h-3" />
                      {roleBadgeConfig.label}
                    </span>

                    {/* Badge Type Visioconférence Teams vs Présentiel */}
                    {meeting.isOnlineMeeting ? (
                      <span className="px-2 py-0.5 text-[10px] font-bold bg-indigo-950 text-indigo-300 border border-indigo-700/60 rounded flex items-center gap-1 shadow-sm">
                        <Video className="w-3 h-3 text-indigo-400" />
                        Teams
                      </span>
                    ) : (
                      <span className="px-2 py-0.5 text-[10px] font-medium bg-slate-800 text-slate-300 border border-slate-700 rounded flex items-center gap-1">
                        <MapPin className="w-3 h-3 text-amber-400" />
                        Présentiel
                      </span>
                    )}

                    {meeting.isDgfMeeting && (
                      <span className="px-2 py-0.5 text-[10px] font-semibold bg-emerald-950 text-emerald-300 border border-emerald-800 rounded">
                        DGF
                      </span>
                    )}

                    {/* Organisateur et indicateur de domaine */}
                    <span className="text-slate-600">•</span>
                    <span className="text-[11px] text-slate-400 font-mono flex items-center gap-1 truncate">
                      Organisé par : <strong className="text-slate-200">{meeting.organizer.name}</strong>
                      {meeting.isExternalOrganizer ? (
                        <span className="ml-1 px-1.5 py-0.2 text-[9px] font-semibold bg-amber-950 text-amber-300 border border-amber-800 rounded">
                          Locataire externe
                        </span>
                      ) : (
                        <span className="text-slate-500 text-[10px]">({meeting.organizer.email})</span>
                      )}
                    </span>
                  </div>

                  {/* Ligne 2 : Titre de la réunion */}
                  <h3 className="text-sm sm:text-base font-semibold text-white group-hover:text-blue-300 transition-colors leading-snug">
                    {meeting.title}
                  </h3>

                  {/* Ligne 3 : Horaires et participants */}
                  <div className="flex flex-wrap items-center gap-4 text-xs text-slate-400">
                    <span className="flex items-center gap-1.5 font-mono text-slate-300">
                      <Calendar className="w-3.5 h-3.5 text-blue-400" />
                      {formatDate(meeting.start)}
                    </span>
                    <span className="flex items-center gap-1.5 font-mono text-slate-300">
                      <Clock className="w-3.5 h-3.5 text-emerald-400" />
                      {formatTime(meeting.start)} - {formatTime(meeting.end)}
                      <span className="text-[10px] text-slate-500 font-sans">({meeting.startTimeZone || 'UTC'})</span>
                    </span>
                    <span className="flex items-center gap-1.5">
                      <Users className="w-3.5 h-3.5 text-amber-400" />
                      {meeting.attendees.length} participant(s)
                    </span>
                    {meeting.departmentCode && (
                      <span className="text-[11px] text-slate-400 bg-slate-950 px-2 py-0.5 rounded border border-slate-800">
                        {meeting.departmentCode}
                      </span>
                    )}
                  </div>

                  {/* Ligne 4 : Taxonomie des 3 statuts stricts (Disponibilité, Accessibilité, Traitement) */}
                  <div className="flex flex-wrap items-center gap-2 pt-1 border-t border-slate-800/60 text-[11px]">
                    <span className="text-slate-500 text-[10px] font-medium">Statuts :</span>

                    {/* Disponibilité */}
                    <span className={`px-2 py-0.5 rounded border text-[10px] font-medium flex items-center gap-1 ${availabilityBadge.className}`}>
                      <FileText className="w-2.5 h-2.5" />
                      Disponibilité : {availabilityBadge.label}
                    </span>

                    {/* Accessibilité */}
                    <span className={`px-2 py-0.5 rounded border text-[10px] font-bold flex items-center gap-1 ${accessibilityBadge.className}`}>
                      {AccessibilityIcon && <AccessibilityIcon className="w-2.5 h-2.5" />}
                      Accessibilité : {accessibilityBadge.label}
                    </span>

                    {/* Traitement */}
                    <span className={`px-2 py-0.5 rounded border text-[10px] font-semibold flex items-center gap-1 ${processingBadge.className}`}>
                      Traitement : {processingBadge.label}
                    </span>
                  </div>
                </div>

                {/* Colonne droite : Actions (Bouton d'action principal conditionné) */}
                <div className="flex flex-wrap xl:flex-col items-stretch xl:items-end gap-2 shrink-0 border-t xl:border-t-0 pt-3 xl:pt-0 border-slate-800">
                  {/* Bouton d'action principal : « Récupérer et traiter la transcription » */}
                  {meeting.isOnlineMeeting && (
                    <button
                      onClick={() => {
                        setSelectedMeeting(meeting);
                        if (isAccessible) {
                          handleTriggerCheck(true);
                        }
                      }}
                      disabled={!isAccessible}
                      title={
                        isAccessible
                          ? "Lancer la récupération du flux WebVTT et préparer le projet de compte rendu institutionnel"
                          : meeting.transcriptAccessibility === 'restricted_by_sharing'
                          ? "Cette transcription n'est pas accessible avec les droits de votre compte Microsoft 365. Veuillez demander à l'organisateur de vous accorder l'accès."
                          : meeting.transcriptAccessibility === 'external_tenant'
                          ? "Non accessible : La réunion a été organisée sur un locataire externe sans accès inter-locataires."
                          : meeting.transcriptAccessibility === 'no_transcript'
                          ? "Non accessible : Aucune transcription n'a été enregistrée pour cette séance Teams."
                          : "Non accessible : Droits d'accès ou transcription non confirmés."
                      }
                      className={`px-3.5 py-2 text-xs font-bold rounded-lg flex items-center justify-center gap-2 transition-all shadow-sm ${
                        isAccessible
                          ? 'text-white bg-emerald-600 hover:bg-emerald-500 border border-emerald-400/40 cursor-pointer shadow-emerald-950/50'
                          : 'text-slate-500 bg-slate-950/70 border border-slate-800/80 cursor-not-allowed opacity-60'
                      }`}
                    >
                      <Sparkles className={`w-3.5 h-3.5 ${isAccessible ? 'text-emerald-200' : 'text-slate-600'}`} />
                      Récupérer et traiter la transcription
                    </button>
                  )}

                  <div className="flex items-center gap-2 w-full xl:w-auto">
                    {/* Bouton d'inspection des droits et diagnostic */}
                    <button
                      onClick={() => setSelectedMeeting(meeting)}
                      className="px-3 py-1.5 text-xs font-medium text-slate-200 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg flex items-center gap-1.5 transition-colors flex-1 xl:flex-none justify-center cursor-pointer"
                    >
                      <Info className="w-3.5 h-3.5 text-blue-400" />
                      Détails & Droits
                    </button>

                    {meeting.joinWebUrl && (
                      <a
                        href={meeting.joinWebUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="px-3 py-1.5 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 border border-indigo-400/30 rounded-lg transition-all flex items-center gap-1.5 shadow-sm"
                      >
                        <Video className="w-3.5 h-3.5" />
                        Teams
                      </a>
                    )}

                    {meeting.webLink && (
                      <a
                        href={meeting.webLink}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="p-1.5 text-slate-400 hover:text-white bg-slate-950 hover:bg-slate-800 border border-slate-800 rounded-lg transition-colors"
                        title="Ouvrir dans Outlook Web"
                      >
                        <ExternalLink className="w-4 h-4" />
                      </a>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* 6. Modal de consultation des détails d'une réunion */}
      {selectedMeeting && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-2xl w-full p-6 space-y-5 shadow-2xl max-h-[90vh] overflow-y-auto">
            {/* Modal Header */}
            <div className="flex items-start justify-between border-b border-slate-800 pb-3">
              <div className="space-y-1">
                <div className="flex items-center gap-2">
                  <span className={`px-2 py-0.5 text-[10px] font-bold rounded border ${
                    selectedMeeting.isOnlineMeeting
                      ? 'bg-indigo-950 text-indigo-300 border-indigo-800'
                      : 'bg-slate-800 text-slate-300 border-slate-700'
                  }`}>
                    {selectedMeeting.isOnlineMeeting ? 'Session Microsoft Teams' : 'Rendez-vous Calendrier'}
                  </span>
                  {selectedMeeting.isDgfMeeting && (
                    <span className="px-2 py-0.5 text-[10px] font-semibold bg-emerald-950 text-emerald-300 border border-emerald-800 rounded">
                      DGF
                    </span>
                  )}
                </div>
                <h3 className="text-base font-bold text-white">{selectedMeeting.title}</h3>
              </div>
              <button
                onClick={() => setSelectedMeeting(null)}
                className="text-slate-400 hover:text-white text-sm p-1 rounded hover:bg-slate-800 transition-colors"
              >
                ✕
              </button>
            </div>

            {/* Informations Générales */}
            <div className="space-y-2.5 text-xs text-slate-300 bg-slate-950 p-4 rounded-xl border border-slate-800">
              <div className="flex flex-col sm:flex-row sm:justify-between py-1 border-b border-slate-800/80 gap-1">
                <span className="text-slate-500 font-medium">Organisateur :</span>
                <span className="font-semibold text-slate-200">
                  {selectedMeeting.organizer.name} ({selectedMeeting.organizer.email || 'Adresse non spécifiée'})
                </span>
              </div>
              <div className="flex flex-col sm:flex-row sm:justify-between py-1 border-b border-slate-800/80 gap-1">
                <span className="text-slate-500 font-medium">Date & Horaire :</span>
                <span className="font-mono text-slate-200">
                  {formatDate(selectedMeeting.start)} • {formatTime(selectedMeeting.start)} - {formatTime(selectedMeeting.end)} ({selectedMeeting.startTimeZone || 'UTC'})
                </span>
              </div>
              <div className="flex flex-col sm:flex-row sm:justify-between py-1 border-b border-slate-800/80 gap-1">
                <span className="text-slate-500 font-medium">Lieu / Emplacement :</span>
                <span className="text-slate-200 font-medium">{selectedMeeting.location || 'Non spécifié'}</span>
              </div>
              <div className="flex flex-col sm:flex-row sm:justify-between py-1 gap-1">
                <span className="text-slate-500 font-medium">Type de réunion :</span>
                <span className={selectedMeeting.isOnlineMeeting ? 'text-indigo-400 font-medium' : 'text-slate-300'}>
                  {selectedMeeting.isOnlineMeeting ? 'Visioconférence Teams (En ligne)' : 'Présentiel / Réunion classique'}
                </span>
              </div>
            </div>

            {/* Aperçu du corps du message Outlook (bodyPreview) */}
            {selectedMeeting.bodyPreview && (
              <div className="space-y-1.5">
                <h4 className="text-xs font-semibold text-slate-300">Description / Ordre du jour (Outlook) :</h4>
                <div className="p-3 bg-slate-950/80 border border-slate-800 rounded-lg text-xs text-slate-300 leading-relaxed max-h-32 overflow-y-auto">
                  {selectedMeeting.bodyPreview}
                </div>
              </div>
            )}

            {/* Liste des participants convoqués */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <h4 className="text-xs font-semibold text-slate-300">
                  Participants ({selectedMeeting.attendees.length}) :
                </h4>
                <span className="text-[11px] text-slate-500">Statut de réponse Outlook</span>
              </div>
              <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1">
                {selectedMeeting.attendees.length === 0 ? (
                  <div className="text-xs text-slate-500 italic p-2 bg-slate-950 rounded">
                    Aucun participant externe renseigné dans l'invitation.
                  </div>
                ) : (
                  selectedMeeting.attendees.map((att: MeetingAttendee, i: number) => (
                    <div
                      key={i}
                      className="flex items-center justify-between text-xs py-1.5 px-3 rounded-lg bg-slate-950 border border-slate-800/70"
                    >
                      <div className="truncate mr-2">
                        <span className="text-slate-200 font-medium">{att.name}</span>
                        {att.email && (
                          <span className="text-[11px] text-slate-500 block truncate">{att.email}</span>
                        )}
                      </div>
                      <div className="flex items-center gap-2 shrink-0">
                        {att.type && (
                          <span className="text-[10px] text-slate-500 uppercase font-mono">
                            {att.type === 'required' ? 'Requis' : 'Optionnel'}
                          </span>
                        )}
                        {getResponseStatusBadge(att.responseStatus)}
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>

            {/* Zone « Transcription Microsoft Teams et Préparation du Compte Rendu » */}
            <div className="space-y-3.5 border-t border-slate-800 pt-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <FileText className="w-4 h-4 text-emerald-400" />
                  <h4 className="text-xs font-bold text-white uppercase tracking-wider">
                    Transcription Microsoft Teams & Préparation du Compte Rendu
                  </h4>
                </div>
                {transcriptCheck && (
                  <span className="text-[10px] text-slate-500 font-mono">
                    {transcriptCheck.cached ? 'État en cache local' : 'Évaluation en temps réel'}
                  </span>
                )}
              </div>

              {/* Rendu des 8 états stricts de transcription */}
              {transcriptCheck && (
                <div className="bg-slate-950 p-4 rounded-xl border border-slate-800 space-y-3">
                  {/* En-tête de statut avec badge */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800/80 pb-3">
                    <span className="text-xs text-slate-400">Statut de la transcription :</span>

                    {/* État 1 : Réunion programmée */}
                    {transcriptCheck.status === 'scheduled' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-blue-950 text-blue-300 border border-blue-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <Calendar className="w-3.5 h-3.5 text-blue-400" />
                        Réunion programmée
                      </span>
                    )}

                    {/* État 2 : Réunion en cours (affichage purement temporel, aucune prétention de capture) */}
                    {transcriptCheck.status === 'in_progress' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-amber-950 text-amber-300 border border-amber-800 rounded-lg flex items-center gap-1.5 shadow-sm animate-pulse">
                        <Clock className="w-3.5 h-3.5 text-amber-400" />
                        Réunion en cours de séance
                      </span>
                    )}

                    {/* État 3 : Réunion terminée (prête pour vérification) */}
                    {transcriptCheck.status === 'completed' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-indigo-950 text-indigo-300 border border-indigo-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <CheckCircle2 className="w-3.5 h-3.5 text-indigo-400" />
                        Réunion terminée (Vérification requise)
                      </span>
                    )}

                    {/* État 4 : Transcription en attente */}
                    {transcriptCheck.status === 'transcript_pending' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-amber-950 text-amber-300 border border-amber-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <RefreshCw className="w-3.5 h-3.5 text-amber-400 animate-spin" />
                        Transcription en cours de compilation Teams
                      </span>
                    )}

                    {/* État 5 : Transcription disponible (strictement confirmé par Microsoft Graph) */}
                    {transcriptCheck.status === 'transcript_available' && (
                      <span className="px-2.5 py-1 text-xs font-bold bg-emerald-950 text-emerald-300 border border-emerald-700 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <CheckCheck className="w-3.5 h-3.5 text-emerald-400" />
                        Transcription confirmée & disponible
                      </span>
                    )}

                    {/* État 6 : Transcription indisponible / collection vide */}
                    {transcriptCheck.status === 'transcript_unavailable' && (
                      <span className="px-2.5 py-1 text-xs font-medium bg-slate-900 text-slate-400 border border-slate-800 rounded-lg flex items-center gap-1.5">
                        <XCircle className="w-3.5 h-3.5 text-slate-500" />
                        Aucune transcription générée
                      </span>
                    )}

                    {/* État 6.B : Restreint par les paramètres de partage Teams de l'organisateur (Organizers only) */}
                    {transcriptCheck.status === 'sharing_restricted_by_organizer' && (
                      <span className="px-2.5 py-1 text-xs font-bold bg-rose-950 text-rose-300 border border-rose-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <Lock className="w-3.5 h-3.5 text-rose-400" />
                        Partage restreint par l'organisateur
                      </span>
                    )}

                    {/* État 6.C : Organisateur externe (Accès cross-tenant restreint) */}
                    {transcriptCheck.status === 'external_tenant_restricted' && (
                      <span className="px-2.5 py-1 text-xs font-bold bg-amber-950 text-amber-300 border border-amber-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <ShieldAlert className="w-3.5 h-3.5 text-amber-400" />
                        Locataire externe (Restreint)
                      </span>
                    )}

                    {/* État 7.A : Permission déléguée Entra ID non accordée */}
                    {transcriptCheck.status === 'permission_missing' && (
                      <span className="px-2.5 py-1 text-xs font-bold bg-amber-950 text-amber-300 border border-amber-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <Lock className="w-3.5 h-3.5 text-amber-400" />
                        Autorisation administrateur en attente
                      </span>
                    )}

                    {/* État 7.B : Accès Graph désactivé sur le tenant Teams */}
                    {transcriptCheck.status === 'tenant_access_disabled' && (
                      <span className="px-2.5 py-1 text-xs font-bold bg-purple-950 text-purple-300 border border-purple-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <ShieldAlert className="w-3.5 h-3.5 text-purple-400" />
                        Accès API désactivé (Tenant Teams)
                      </span>
                    )}

                    {/* État 7.C : Réunion Teams non résolue */}
                    {transcriptCheck.status === 'meeting_unresolved' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-amber-950 text-amber-300 border border-amber-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <HelpCircle className="w-3.5 h-3.5 text-amber-400" />
                        Réunion Teams non résolue
                      </span>
                    )}

                    {/* État 7.D : Réunion introuvable ou expirée (HTTP 404) */}
                    {transcriptCheck.status === 'meeting_not_found' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-rose-950 text-rose-300 border border-rose-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <AlertCircle className="w-3.5 h-3.5 text-rose-400" />
                        Réunion introuvable ou expirée (404)
                      </span>
                    )}

                    {/* État 8 : Erreur temporaire de récupération */}
                    {transcriptCheck.status === 'transient_error' && (
                      <span className="px-2.5 py-1 text-xs font-semibold bg-orange-950 text-orange-300 border border-orange-800 rounded-lg flex items-center gap-1.5 shadow-sm">
                        <AlertTriangle className="w-3.5 h-3.5 text-orange-400" />
                        Erreur temporaire (HTTP 429 ou réseau)
                      </span>
                    )}
                  </div>

                  {/* Message descriptif institutionnel */}
                  <div className="text-xs text-slate-300 space-y-1.5 leading-relaxed">
                    {transcriptCheck.status === 'scheduled' && (
                      <p>
                        Cette séance Microsoft Teams est planifiée dans le futur. Conformément au fonctionnement de Microsoft 365, la transcription ne pourra être consultée qu'après la tenue effective de la réunion si elle a été activée.
                      </p>
                    )}

                    {transcriptCheck.status === 'in_progress' && (
                      <p>
                        La réunion est actuellement en cours. Microsoft Teams génère la transcription en direct si l'organisateur ou un participant autorisé l'a activée dans Teams. ReadMeeting.ai interrogera Microsoft Graph une fois la séance achevée.
                      </p>
                    )}

                    {transcriptCheck.status === 'completed' && (
                      <p>
                        La réunion s'est achevée. Vous pouvez interroger Microsoft Graph pour vérifier si la transcription textuelle est disponible dans Microsoft Teams.
                      </p>
                    )}

                    {transcriptCheck.status === 'transcript_pending' && (
                      <p>
                        Le fichier de transcription est en cours de compilation et d'indexation linguistique par les serveurs Microsoft Teams. Cette opération prend habituellement quelques minutes après la fin de l'appel.
                      </p>
                    )}

                    {transcriptCheck.status === 'transcript_available' && (
                      <div className="space-y-3">
                        <div className="p-3 bg-emerald-950/50 border border-emerald-800/80 rounded-xl space-y-2">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-emerald-900/60 pb-2">
                            <span className="text-xs font-bold text-emerald-300 flex items-center gap-1.5">
                              <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                              Statut : Transcription Teams confirmée & récupérée
                            </span>
                            <button
                              onClick={() => setIsViewerModalOpen(true)}
                              className="px-3 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-500 rounded-lg flex items-center gap-1.5 shadow-sm transition-all cursor-pointer self-start sm:self-auto"
                            >
                              <FileText className="w-3.5 h-3.5" />
                              Consulter la transcription
                            </button>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] text-slate-300 font-mono">
                            <div>
                              <span className="text-slate-500 block">Date de la transcription :</span>
                              <span className="text-white font-medium">
                                {transcriptCheck.metadata?.createdDateTime
                                  ? new Date(transcriptCheck.metadata.createdDateTime).toLocaleDateString('fr-FR', {
                                      day: 'numeric',
                                      month: 'long',
                                      year: 'numeric',
                                      hour: '2-digit',
                                      minute: '2-digit',
                                    })
                                  : 'Horodatage non renseigné'}
                              </span>
                            </div>
                            <div>
                              <span className="text-slate-500 block">Identifiant Teams :</span>
                              <span className="text-emerald-400 truncate block" title={transcriptCheck.metadata?.id}>
                                {transcriptCheck.metadata?.id || 'ID non disponible'}
                              </span>
                            </div>
                          </div>
                        </div>

                        {transcriptCheck.content && (
                          <div className="space-y-2.5">
                            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 p-2.5 bg-slate-900 rounded-lg border border-slate-800 text-[11px]">
                              <div>
                                <span className="text-slate-500 block">Durée totale :</span>
                                <span className="text-white font-bold">{Math.round(transcriptCheck.content.durationSeconds / 60)} min ({Math.round(transcriptCheck.content.durationSeconds)}s)</span>
                              </div>
                              <div>
                                <span className="text-slate-500 block">Interventions :</span>
                                <span className="text-white font-bold">{transcriptCheck.content.cues.length} segments</span>
                              </div>
                              <div>
                                <span className="text-slate-500 block">Mots transcrits :</span>
                                <span className="text-white font-bold">{transcriptCheck.content.totalWords.toLocaleString('fr-FR')}</span>
                              </div>
                              <div>
                                <span className="text-slate-500 block">Locuteurs :</span>
                                <span className="text-emerald-400 font-bold">{transcriptCheck.content.speakersCount} identifiés</span>
                              </div>
                            </div>

                            {transcriptCheck.content.speakers && transcriptCheck.content.speakers.length > 0 && (
                              <div className="p-2.5 bg-slate-900/60 rounded-lg border border-slate-800/80 space-y-1.5">
                                <span className="text-[11px] font-semibold text-slate-300 block">
                                  Locuteurs identifiés (Attribution nominale Teams) :
                                </span>
                                <div className="flex flex-wrap items-center gap-1.5">
                                  {transcriptCheck.content.speakers.map((spk) => (
                                    <span
                                      key={spk.name}
                                      className="px-2 py-0.5 rounded-md bg-slate-950 border border-slate-800 text-[10px] text-slate-300 flex items-center gap-1.5"
                                    >
                                      <span className={`w-1.5 h-1.5 rounded-full bg-gradient-to-r ${spk.avatarColor || 'from-blue-500 to-indigo-700'}`} />
                                      <span className="font-medium text-white">{spk.name}</span>
                                      <span className="text-slate-500 font-mono">({spk.interventionsCount} int.)</span>
                                    </span>
                                  ))}
                                </div>
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    )}

                    {transcriptCheck.status === 'transcript_unavailable' && (
                      <div className="space-y-1">
                        <p className="text-slate-400">
                          {transcriptCheck.error?.message || "La réunion Teams s'est achevée sans que la transcription en direct n'ait été activée dans Teams. Aucune transcription n'a été enregistrée pour cette séance."}
                        </p>
                        {transcriptCheck.error?.teamsPolicyNote && (
                          <p className="text-[11px] text-slate-500 italic">
                            Note : {transcriptCheck.error.teamsPolicyNote}
                          </p>
                        )}
                      </div>
                    )}

                    {transcriptCheck.status === 'tenant_access_disabled' && (
                      <div className="space-y-2">
                        <p className="text-purple-200">
                          {transcriptCheck.error?.message}
                        </p>
                        <div className="p-2.5 bg-purple-950/40 border border-purple-900/60 rounded-lg text-[11px] text-purple-300 space-y-1">
                          <div className="flex items-center gap-1.5 font-bold">
                            <ShieldAlert className="w-3.5 h-3.5 text-purple-400" />
                            Configuration requise dans le Teams Admin Center :
                          </div>
                          <div>Paramètre : <code className="font-mono bg-purple-950 px-1 py-0.5 rounded text-white">GraphAccessToTranscriptsDisabled = $false</code></div>
                          <div>Action : Activer l'accès API Microsoft Graph aux transcriptions dans les paramètres de réunion Teams.</div>
                        </div>
                      </div>
                    )}

                    {transcriptCheck.status === 'sharing_restricted_by_organizer' && (
                      <div className="p-3.5 bg-rose-950/40 border border-rose-800/80 rounded-xl space-y-2.5">
                        <div className="flex items-center gap-2 text-rose-300 font-bold text-xs">
                          <Lock className="w-4 h-4 text-rose-400" />
                          Accès à la transcription restreint par l'organisateur
                        </div>
                        <p className="text-xs text-white font-medium bg-rose-950/90 p-3 rounded-lg border border-rose-900/80 leading-relaxed shadow-sm">
                          « Cette transcription n'est pas accessible avec les droits de votre compte Microsoft 365. Veuillez demander à l'organisateur de vous accorder l'accès. »
                        </p>
                        <div className="text-[11px] text-rose-200/90 space-y-1 bg-slate-900/60 p-2.5 rounded-lg border border-slate-800">
                          <p><strong>Cause identifiée :</strong> L'organisateur de la réunion a configuré les options de réunion Teams (<em>« Qui peut afficher la transcription »</em>) sur <em>« Organisateurs et coorganisateurs »</em>.</p>
                          <p><strong>Action attendue :</strong> Contactez l'organisateur <strong>{selectedMeeting.organizer.name}</strong> ({selectedMeeting.organizer.email}) pour lui demander de modifier les options de la réunion Teams sur <em>« Tous les participants »</em> (Everyone).</p>
                        </div>
                      </div>
                    )}

                    {transcriptCheck.status === 'external_tenant_restricted' && (
                      <div className="p-3.5 bg-amber-950/40 border border-amber-800/80 rounded-xl space-y-2.5">
                        <div className="flex items-center gap-2 text-amber-300 font-bold text-xs">
                          <ShieldAlert className="w-4 h-4 text-amber-400" />
                          Réunion convoquée par un locataire externe (Bailleurs / Partenaires)
                        </div>
                        <p className="text-xs text-slate-200 leading-relaxed">
                          L'organisateur de cette réunion appartient au domaine externe <strong>{selectedMeeting.organizer.email.split('@')[1] || 'externe'}</strong>. Conformément aux politiques de sécurité Entra ID et de souveraineté DGF, les droits délégués ne peuvent traverser les limites du locataire sans fédération inter-organisations (B2B).
                        </p>
                        <div className="text-[11px] text-amber-200/90 bg-amber-950/80 p-2.5 rounded-lg border border-amber-900/60">
                          <strong>Solution recommandée :</strong> Sollicitez l'organisateur ({selectedMeeting.organizer.name}) pour obtenir l'exportation du fichier WebVTT officiel ou la transmission du compte rendu institutionnel.
                        </div>
                      </div>
                    )}

                    {/* Piste d'audit et de traçabilité institutionnelle DGF */}
                    {transcriptCheck.auditTrail && (
                      <div className="p-3 bg-slate-900/90 border border-slate-800 rounded-xl space-y-2 text-xs">
                        <div className="flex items-center justify-between text-[11px] font-bold text-slate-300 border-b border-slate-800 pb-1.5">
                          <span className="flex items-center gap-1.5 text-blue-300">
                            <ShieldAlert className="w-3.5 h-3.5 text-blue-400" />
                            Traçabilité & Piste d'audit institutionnelle (DGF)
                          </span>
                          <span className="text-[10px] text-emerald-400 bg-emerald-950 px-1.5 py-0.5 rounded border border-emerald-800 font-mono">
                            Gouvernance DGF
                          </span>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] text-slate-300 font-mono">
                          <div>Demandeur : <strong className="text-white">{transcriptCheck.auditTrail.requestedByName}</strong> ({transcriptCheck.auditTrail.requestedByUpn})</div>
                          <div>Rôle vérifié : <strong className="text-emerald-400 uppercase">{transcriptCheck.auditTrail.userRole}</strong></div>
                          <div>Organisateur : <strong className="text-white">{transcriptCheck.auditTrail.organizerName}</strong> ({transcriptCheck.auditTrail.organizerUpn})</div>
                          <div>Horodatage : <strong className="text-slate-200">{new Date(transcriptCheck.auditTrail.retrievedAt).toLocaleString('fr-FR')}</strong></div>
                          <div className="sm:col-span-2 truncate">Session Teams : <span className="text-slate-400">{transcriptCheck.auditTrail.onlineMeetingId}</span></div>
                        </div>
                      </div>
                    )}

                    {transcriptCheck.status === 'permission_missing' && (
                      <div className="space-y-2">
                        <p className="text-rose-200">
                          {transcriptCheck.error?.message || "L'autorisation déléguée 'OnlineMeetingTranscript.Read.All' n'est pas encore consentie sur le locataire Entra ID de la DGF."}
                        </p>
                        <div className="p-2.5 bg-rose-950/40 border border-rose-900/60 rounded-lg text-[11px] text-rose-300 space-y-1">
                          <div className="flex items-center gap-1.5 font-bold">
                            <ShieldAlert className="w-3.5 h-3.5 text-rose-400" />
                            Blocage administratif identifié :
                          </div>
                          <div>Permission requise : <code className="font-mono bg-rose-950 px-1 py-0.5 rounded text-white">OnlineMeetingTranscript.Read.All</code></div>
                          <div>Statut : <span className="underline font-semibold">Consentement administrateur requis dans Microsoft Entra ID</span></div>
                        </div>
                      </div>
                    )}

                    {transcriptCheck.status === 'meeting_unresolved' && (
                      <div className="space-y-1 text-slate-400">
                        <p>{transcriptCheck.error?.message}</p>
                        <p className="text-[11px] text-slate-500 italic">
                          Action : {transcriptCheck.error?.actionRequired}
                        </p>
                      </div>
                    )}

                    {transcriptCheck.status === 'meeting_not_found' && (
                      <p className="text-rose-300">
                        {transcriptCheck.error?.message || "La ressource de réunion demandée est introuvable ou a expiré dans Microsoft Teams."}
                      </p>
                    )}

                    {transcriptCheck.status === 'transient_error' && (
                      <p className="text-orange-300">
                        {transcriptCheck.error?.message || "Une erreur temporaire ou un dépassement de quota Microsoft Graph (HTTP 429) est survenu. Veuillez patienter avant de réinterroger."}
                      </p>
                    )}

                    {/* Rappel institutionnel : Transcription textuelle vs Enregistrement vidéo */}
                    <div className="mt-2 pt-2 border-t border-slate-900 grid grid-cols-1 sm:grid-cols-2 gap-2 text-[10px] text-slate-400">
                      <div className="flex items-start gap-1.5">
                        <FileText className="w-3 h-3 text-emerald-400 shrink-0 mt-0.5" />
                        <span><strong>Transcription Teams (.vtt) :</strong> flux textuel structuré avec orateurs et minutage, exploité par ReadMeeting.ai via Graph.</span>
                      </div>
                      <div className="flex items-start gap-1.5">
                        <Video className="w-3 h-3 text-blue-400 shrink-0 mt-0.5" />
                        <span><strong>Enregistrement vidéo (.mp4) :</strong> fichier audiovisuel OneDrive/SharePoint, indépendant de l'API de transcription.</span>
                      </div>
                    </div>
                  </div>

                  {/* Boutons d'action pour la vérification */}
                  {selectedMeeting.isOnlineMeeting && (
                    <div className="pt-2 flex flex-wrap items-center gap-2">
                      <button
                        onClick={() => handleTriggerCheck(true)}
                        disabled={isCheckingTranscript || !user}
                        className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 disabled:opacity-50 rounded-lg flex items-center gap-1.5 transition-all shadow-sm"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${isCheckingTranscript ? 'animate-spin' : ''}`} />
                        {isCheckingTranscript
                          ? 'Interrogation Microsoft Graph en cours...'
                          : transcriptCheck.status === 'completed'
                          ? 'Vérifier la disponibilité de la transcription'
                          : 'Réinterroger Microsoft Graph'}
                      </button>

                      {transcriptCheck.status === 'transcript_available' && transcriptCheck.content && (
                        <button
                          onClick={() => setIsViewerModalOpen(true)}
                          className="px-3.5 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-500 rounded-lg flex items-center gap-1.5 transition-all shadow-sm cursor-pointer"
                        >
                          <FileText className="w-3.5 h-3.5" />
                          Consulter la transcription
                        </button>
                      )}

                      {transcriptCheck.status === 'permission_missing' && (
                        <button
                          onClick={() => setAdminGuideOpen(!adminGuideOpen)}
                          className="px-3 py-1.5 text-xs font-medium text-rose-300 hover:text-white bg-rose-950/60 hover:bg-rose-900/60 border border-rose-800 rounded-lg flex items-center gap-1.5 transition-colors"
                        >
                          <Info className="w-3.5 h-3.5" />
                          Instructions Administrateur DGF
                          {adminGuideOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                        </button>
                      )}

                      <button
                        onClick={() => setPipelinePreviewOpen(!pipelinePreviewOpen)}
                        className="px-3 py-1.5 text-xs font-medium text-slate-300 hover:text-white bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg flex items-center gap-1.5 transition-colors"
                      >
                        <Layers className="w-3.5 h-3.5 text-emerald-400" />
                        Pipeline documentaire DGF
                        {pipelinePreviewOpen ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
                      </button>
                    </div>
                  )}

                  {/* Volet déroulant : Instructions Administrateur DGF */}
                  {adminGuideOpen && (
                    <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-2.5 text-xs text-slate-300">
                      <div className="flex items-center gap-2 font-bold text-white">
                        <ShieldAlert className="w-4 h-4 text-rose-400" />
                        Procédure d'activation pour l'administrateur Microsoft 365 DGF
                      </div>
                      <ol className="list-decimal list-inside space-y-1.5 text-slate-300 leading-relaxed pl-1 text-[11px]">
                        <li>Connectez-vous au portail <strong>Microsoft Entra ID</strong> (<code className="text-blue-300">https://entra.microsoft.com</code>) avec un compte Administrateur Général.</li>
                        <li>Accédez à <strong>Identité &gt; Applications &gt; Inscriptions d'applications</strong>.</li>
                        <li>Sélectionnez l'application <strong>ReadMeeting.ai | DGF</strong>.</li>
                        <li>Rendez-vous dans la rubrique <strong>Autorisations d'API</strong>.</li>
                        <li>Cliquez sur <strong>Ajouter une autorisation &gt; Microsoft Graph &gt; Autorisations déléguées</strong>.</li>
                        <li>Recherchez et cochez : <code className="text-emerald-400 font-mono font-bold">OnlineMeetingTranscript.Read.All</code>.</li>
                        <li>Cliquez sur <strong>Accorder un consentement d'administrateur pour la DGF</strong>.</li>
                      </ol>
                    </div>
                  )}

                  {/* Volet déroulant : Pipeline documentaire institutionnel */}
                  {pipelinePreviewOpen && (
                    <div className="p-3.5 bg-slate-900 border border-slate-800 rounded-xl space-y-3 text-xs text-slate-300">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-white flex items-center gap-2">
                          <Layers className="w-4 h-4 text-emerald-400" />
                          Architecture du Pipeline de Traitement Documentaire DGF
                        </span>
                        <span className="text-[10px] text-emerald-400 font-mono bg-emerald-950 px-2 py-0.5 rounded border border-emerald-800">
                          Validation Humaine Stricte
                        </span>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px]">
                        <div className="p-2 bg-slate-950 rounded border border-slate-800">
                          <span className="font-semibold text-slate-200 block">1. Récupération Teams</span>
                          Flux VTT via <code className="text-slate-400 font-mono text-[10px]">/transcripts/{'{id}'}/content</code>
                        </div>
                        <div className="p-2 bg-slate-950 rounded border border-slate-800">
                          <span className="font-semibold text-slate-200 block">2. Nettoyage & Normalisation</span>
                          Élimination du bruit, synchronisation temporelle
                        </div>
                        <div className="p-2 bg-slate-950 rounded border border-slate-800">
                          <span className="font-semibold text-slate-200 block">3. Analyse Thématique</span>
                          Segmentation par ordre du jour institutionnel
                        </div>
                        <div className="p-2 bg-slate-950 rounded border border-slate-800">
                          <span className="font-semibold text-slate-200 block">4. Extraction des Décisions</span>
                          Arbitrages budgétaires, consensus, votes
                        </div>
                        <div className="p-2 bg-slate-950 rounded border border-slate-800">
                          <span className="font-semibold text-slate-200 block">5. Matrice d'Actions</span>
                          Assignation formelle des responsables et échéances
                        </div>
                        <div className="p-2 bg-slate-950 rounded border border-slate-800">
                          <span className="font-semibold text-slate-200 block">6. Projet de Compte Rendu</span>
                          Rédaction du brouillon selon le modèle Word DGF
                        </div>
                      </div>

                      {/* Point d'arrêt et de conformité obligatoire */}
                      <div className="p-2.5 bg-amber-950/40 border border-amber-800/60 rounded-lg text-[11px] text-amber-200 flex items-start gap-2">
                        <Lock className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" />
                        <div>
                          <strong className="block text-white">Garde-fou institutionnel obligatoire (Point d'arrêt) :</strong>
                          Le compte rendu ne sera <strong>JAMAIS</strong> automatiquement considéré comme officiel ni diffusé sans la validation humaine formelle, la relecture et la signature numérique du rapporteur ou du président de séance.
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Actions de la modale */}
            <div className="pt-2 flex items-center justify-between border-t border-slate-800">
              <div className="flex items-center gap-2">
                {selectedMeeting.joinWebUrl && (
                  <a
                    href={selectedMeeting.joinWebUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="px-3.5 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg flex items-center gap-1.5 transition-all shadow-sm"
                  >
                    <Video className="w-3.5 h-3.5" />
                    Rejoindre Teams
                  </a>
                )}
                {selectedMeeting.webLink && (
                  <a
                    href={selectedMeeting.webLink}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="px-3 py-2 text-xs font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 rounded-lg flex items-center gap-1.5 transition-colors border border-slate-700"
                  >
                    <ExternalLink className="w-3.5 h-3.5" />
                    Ouvrir dans Outlook
                  </a>
                )}
              </div>

              <button
                onClick={() => setSelectedMeeting(null)}
                className="px-4 py-2 text-xs font-medium text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition-colors"
              >
                Fermer
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 4. Modal de consultation de la transcription réelle Teams */}
      {selectedMeeting && (
        <TranscriptViewerModal
          isOpen={isViewerModalOpen}
          onClose={() => setIsViewerModalOpen(false)}
          meetingTitle={selectedMeeting.title}
          metadata={transcriptCheck?.metadata}
          content={transcriptCheck?.content}
          onlineMeetingId={transcriptCheck?.onlineMeetingId}
        />
      )}
    </div>
  );
};
