/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Types et interfaces pour le service de transcription Microsoft Teams
 * et le pipeline de traitement documentaire institutionnel.
 */

import { UserMeetingRole, TranscriptAccessibilityStatus } from './meeting';

export type MeetingTemporalStatus = 'scheduled' | 'in_progress' | 'completed';

export type TranscriptStatus =
  | 'scheduled'                            // Réunion programmée (date de début/fin future)
  | 'in_progress'                          // Réunion en cours (actuellement en séance)
  | 'completed'                            // Réunion terminée (en attente de consultation Graph)
  | 'transcript_pending'                  // Réunion terminée, transcription en cours de compilation par Microsoft Teams
  | 'transcript_available'                // Transcription confirmée et récupérée avec succès via Microsoft Graph
  | 'transcript_unavailable'              // Collection de transcriptions vide (réunion non transcrite dans Teams)
  | 'sharing_restricted_by_organizer'     // Restreint par les paramètres de partage Teams de l'organisateur (Organizers and co-organizers only)
  | 'external_tenant_restricted'          // Organisateur issu d'un locataire externe sans accès inter-locataires fédéré
  | 'permission_missing'                  // Permission Entra ID OnlineMeetingTranscript.Read.All non accordée
  | 'tenant_access_disabled'              // Accès Graph aux transcriptions désactivé au niveau du locataire (GraphAccessToTranscriptsDisabled)
  | 'meeting_unresolved'                  // Impossible de résoudre l'ID Teams depuis le joinWebUrl
  | 'meeting_not_found'                   // Réunion introuvable ou expirée dans Microsoft Graph (HTTP 404)
  | 'transient_error';                    // Erreur temporaire ou limitation de débit (HTTP 429 Rate Limit, 503, réseau)

export interface TranscriptCue {
  id: string;
  speakerName: string;
  speakerEmail?: string;
  speakerId?: string;
  startTime: string; // e.g. "00:01:23.450"
  endTime: string;   // e.g. "00:01:35.120"
  startSeconds: number;
  endSeconds: number;
  text: string;
}

export interface TranscriptMetadata {
  id: string;
  meetingId: string;
  createdDateTime: string;
  meetingOrganizerId?: string;
  transcriptContentUrl: string;
  callId?: string;
  language?: string;
}

export interface IdentifiedSpeaker {
  name: string;
  interventionsCount: number;
  totalWords: number;
  avatarColor?: string;
}

export interface TranscriptContent {
  format: 'text/vtt' | 'text/plain';
  rawContent: string;
  cues: TranscriptCue[];
  totalWords: number;
  durationSeconds: number;
  speakersCount: number;
  speakers: IdentifiedSpeaker[];
}

export interface TranscriptErrorDetail {
  code: string;
  title: string;
  message: string;
  httpStatus?: number;
  actionRequired: string;
  retryAfterSeconds?: number;
  missingPermission?: string;
  teamsPolicyNote?: string;
}

export interface TranscriptRetrievalAudit {
  meetingId: string;
  meetingTitle: string;
  onlineMeetingId: string;
  transcriptId: string;
  requestedByUpn: string;
  requestedByName: string;
  userRole: UserMeetingRole;
  organizerUpn: string;
  organizerName: string;
  isExternalOrganizer: boolean;
  retrievedAt: string;
  cuesCount: number;
  totalWords: number;
}

export interface TranscriptCheckResult {
  meetingId: string;
  meetingTitle: string;
  temporalStatus: MeetingTemporalStatus;
  status: TranscriptStatus;
  onlineMeetingId?: string;
  userRole?: UserMeetingRole;
  isExternalOrganizer?: boolean;
  accessibility?: TranscriptAccessibilityStatus;
  metadata?: TranscriptMetadata;
  content?: TranscriptContent;
  error?: TranscriptErrorDetail;
  auditTrail?: TranscriptRetrievalAudit;
  checkedAt: string;
  cached?: boolean;
}

/**
 * Pipeline de traitement documentaire institutionnel (DGF)
 * Strictement conditionné à une validation humaine obligatoire.
 */
export type PipelineStage =
  | '1_transcript_retrieved'
  | '2_content_cleaning'
  | '3_topic_segmentation'
  | '4_decisions_extraction'
  | '5_actions_matrix'
  | '6_draft_generation'
  | '7_human_review_required'  // POINT D'ARRÊT OBLIGATOIRE : Validation humaine requise
  | '8_approved_institutional' // Validé par le rapporteur / président de séance
  | '9_published_sharepoint';  // Publié sur SharePoint Online DGF

export interface ActionItemProposal {
  id: string;
  description: string;
  assigneeName: string;
  assigneeEmail?: string;
  deadlineDate?: string;
  priority: 'low' | 'medium' | 'high' | 'urgent';
  contextSnippet?: string;
}

export interface DecisionItemProposal {
  id: string;
  topic: string;
  decisionText: string;
  votedOrConsensus: 'consensus' | 'formal_vote' | 'presidential_arbitration';
  budgetImpact?: string;
}

export interface DocumentDraftProposal {
  meetingTitle: string;
  meetingDate: string;
  departmentCode: string;
  chairperson: string;
  rapporteur: string;
  executiveSummary: string;
  topicsDiscussed: Array<{
    title: string;
    summary: string;
    keyPoints: string[];
  }>;
  decisions: DecisionItemProposal[];
  actions: ActionItemProposal[];
}

export interface HumanValidationAudit {
  isApproved: boolean;
  validationStatus: 'draft' | 'under_review' | 'changes_requested' | 'approved';
  reviewedByUpn?: string;
  reviewedByName?: string;
  reviewedAt?: string;
  reviewerComments?: string;
  digitalSignatureRef?: string;
}
