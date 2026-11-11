export type MeetingCaptureStatus =
  | 'scheduled'
  | 'in_progress'
  | 'completed'
  | 'transcribing'
  | 'transcribed'
  | 'report_generated'
  | 'archived';

export type UserMeetingRole = 'organizer' | 'coorganizer' | 'attendee' | 'undetermined';

export type TranscriptAccessibilityStatus =
  | 'accessible'                  // Transcription confirmée et accessible avec les droits du compte connecté
  | 'restricted_by_sharing'       // Restreint par les paramètres de partage Teams de l'organisateur (Organizers only)
  | 'external_tenant'             // Organisateur issu d'un tenant externe (accès délégué inter-organisations restreint)
  | 'no_transcript'               // Aucune transcription enregistrée pour cette réunion
  | 'unresolved'                  // Identifiant onlineMeetingId non résolu
  | 'permission_missing'          // Permission Graph manquante
  | 'pending_check';              // En attente d'interrogation Graph

export type MeetingProcessingStatus =
  | 'idle'                        // Non traité
  | 'ready_for_processing'        // Transcription récupérée, prêt pour le traitement documentaire
  | 'processing_pending'          // Traitement documentaire initié
  | 'processed';                  // Compte rendu généré / archivé

export interface MeetingAttendee {
  name: string;
  email: string;
  type?: 'required' | 'optional' | 'resource' | string;
  role?: string;
  department?: string;
  isExternal?: boolean;
  responseStatus?: string;
}

export interface MeetingItem {
  id: string;
  graphId?: string;
  title: string;
  subject?: string;
  bodyPreview?: string;
  organizer: {
    name: string;
    email: string;
    department?: string;
    isExternal?: boolean;
  };
  start: string;
  end: string;
  startTimeZone?: string;
  endTimeZone?: string;
  location?: string;
  isOnlineMeeting?: boolean;
  onlineMeetingProvider?: string;
  joinWebUrl?: string;
  webLink?: string;
  status: MeetingCaptureStatus;
  captureMethod: 'teams_native' | 'meeting_bot';
  isDgfMeeting: boolean;
  departmentCode?: string;
  attendees: MeetingAttendee[];
  hasTranscript: boolean;
  hasOfficialReport: boolean;
  reportReference?: string;
  userRole?: UserMeetingRole;
  isExternalOrganizer?: boolean;
  organizerTenantId?: string;
  transcriptAccessibility?: TranscriptAccessibilityStatus;
  processingStatus?: MeetingProcessingStatus;
}

export interface OfficialReportItem {
  id: string;
  referenceNumber: string;
  meetingTitle: string;
  meetingDate: string;
  departmentCode: string;
  chairperson: string;
  rapporteur: string;
  status: 'draft' | 'under_review' | 'approved' | 'published';
  sharePointDocxUrl?: string;
  executiveSummary: string;
  actionItemsCount: number;
}
