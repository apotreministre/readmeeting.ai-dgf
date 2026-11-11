/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Service de récupération, de diagnostic et de gestion du cache des transcriptions Microsoft Teams.
 * Conforme à Microsoft Graph API v1.0, aux stratégies Teams et aux exigences de gouvernance DGF.
 */

import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import { MeetingItem, UserMeetingRole, TranscriptAccessibilityStatus } from '../types/meeting';
import { DgfUserProfile } from '../types/auth';
import {
  MeetingTemporalStatus,
  TranscriptCheckResult,
  TranscriptContent,
  TranscriptCue,
  TranscriptErrorDetail,
  TranscriptMetadata,
  TranscriptStatus,
  HumanValidationAudit,
  IdentifiedSpeaker,
  TranscriptRetrievalAudit
} from '../types/transcript';

// Cache en mémoire pour éviter les requêtes Graph répétées et le dépassement de quota (HTTP 429)
interface CacheEntry {
  result: TranscriptCheckResult;
  expiresAt: number;
}

const memoryCache = new Map<string, CacheEntry>();
const CACHE_TTL_SUCCESS_MS = 10 * 60 * 1000; // 10 minutes pour une transcription confirmée
const CACHE_TTL_ERROR_MS = 2 * 60 * 1000;     // 2 minutes pour une absence ou erreur temporaire

/**
 * 1. Détermine le statut temporel d'une réunion (programmée, en cours, terminée)
 * Basé strictement sur l'horloge et les heures de début et fin de convocation.
 */
export function getMeetingTemporalStatus(
  startIso: string,
  endIso: string,
  now: Date = new Date()
): MeetingTemporalStatus {
  const startDate = new Date(startIso);
  const endDate = new Date(endIso);

  if (now < startDate) {
    return 'scheduled';
  }
  if (now >= startDate && now <= endDate) {
    return 'in_progress';
  }
  return 'completed';
}

/**
  * Détermine le rôle de l'utilisateur connecté dans la réunion :
  * - 'organizer' : L'utilisateur est l'organisateur convoquant
  * - 'coorganizer' : L'utilisateur a été désigné coorganisateur
  * - 'attendee' : L'utilisateur a été convoqué comme participant invité
  * - 'undetermined' : Simple détenteur de lien, absent de l'invitation
  */
export function detectUserMeetingRole(
  meeting: MeetingItem,
  user: DgfUserProfile | null
): UserMeetingRole {
  if (!user) return 'undetermined';

  const userMail = (user.mail || '').toLowerCase().trim();
  const userUpn = (user.userPrincipalName || '').toLowerCase().trim();
  const userName = (user.displayName || '').toLowerCase().trim();

  const orgEmail = (meeting.organizer.email || '').toLowerCase().trim();
  const orgName = (meeting.organizer.name || '').toLowerCase().trim();

  // 1. Contrôle du rôle Organisateur
  if (
    (userMail && orgEmail === userMail) ||
    (userUpn && orgEmail === userUpn) ||
    (userName && orgName && (orgName.includes(userName) || userName.includes(orgName)))
  ) {
    return 'organizer';
  }

  // 2. Contrôle de la convocation dans la liste des participants
  if (Array.isArray(meeting.attendees)) {
    const matchingAttendee = meeting.attendees.find((att) => {
      const attEmail = (att.email || '').toLowerCase().trim();
      const attName = (att.name || '').toLowerCase().trim();
      return (
        (userMail && attEmail === userMail) ||
        (userUpn && attEmail === userUpn) ||
        (userName && attName && (attName.includes(userName) || userName.includes(attName)))
      );
    });

    if (matchingAttendee) {
      if (
        matchingAttendee.role === 'coorganizer' ||
        matchingAttendee.type === 'coorganizer'
      ) {
        return 'coorganizer';
      }
      return 'attendee';
    }
  }

  // 3. Simple détention du lien sans convocation formelle
  return 'undetermined';
}

/**
 * Détecte si l'organisateur de la réunion appartient à une organisation / locataire externe.
 */
export function isOrganizerExternal(
  organizerEmail: string,
  userUpnOrMail?: string
): boolean {
  if (!organizerEmail) return false;
  const orgLower = organizerEmail.toLowerCase().trim();

  // Domaine institutionnel du Ministère des Finances et du Budget DGF
  if (orgLower.endsWith('@finances.gouv.ci') || orgLower.endsWith('.gouv.ci')) {
    return false;
  }

  if (userUpnOrMail && userUpnOrMail.includes('@')) {
    const userDomain = userUpnOrMail.split('@')[1].toLowerCase().trim();
    if (orgLower.endsWith(`@${userDomain}`)) {
      return false;
    }
  }

  return true;
}

/**
 * Convertit un timestamp WebVTT (HH:MM:SS.mmm ou MM:SS.mmm) en secondes
 */
export function timeStringToSeconds(timeStr: string): number {
  const parts = timeStr.trim().split(':');
  if (parts.length === 3) {
    const hours = parseFloat(parts[0]);
    const minutes = parseFloat(parts[1]);
    const seconds = parseFloat(parts[2]);
    return hours * 3600 + minutes * 60 + seconds;
  }
  if (parts.length === 2) {
    const minutes = parseFloat(parts[0]);
    const seconds = parseFloat(parts[1]);
    return minutes * 60 + seconds;
  }
  return 0;
}

/**
 * Analyse et structure le contenu brut WebVTT d'une transcription Teams
 */
export function parseVttTranscript(rawVtt: string): TranscriptContent {
  const lines = rawVtt.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const cues: TranscriptCue[] = [];
  const speakersSet = new Set<string>();

  const timeRegex = /((\d{2}:)?\d{2}:\d{2}\.\d{3})\s*-->\s*((\d{2}:)?\d{2}:\d{2}\.\d{3})/;
  const voiceTagRegex = /<v\s+([^>]+)>(.*?)<\/v>/i;
  const legacyVoiceTagRegex = /<v\s+([^>]+)>(.*)/i;

  let currentStartTime = '';
  let currentEndTime = '';
  let currentSpeaker = 'Intervenant';
  let currentTextBuffer: string[] = [];

  const finalizeCue = () => {
    if (currentStartTime && currentEndTime && currentTextBuffer.length > 0) {
      const fullText = currentTextBuffer.join(' ').trim();
      if (fullText) {
        cues.push({
          id: `cue-${cues.length + 1}`,
          speakerName: currentSpeaker,
          startTime: currentStartTime,
          endTime: currentEndTime,
          startSeconds: timeStringToSeconds(currentStartTime),
          endSeconds: timeStringToSeconds(currentEndTime),
          text: fullText,
        });
        speakersSet.add(currentSpeaker);
      }
    }
    currentTextBuffer = [];
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (!line || line.startsWith('WEBVTT') || line.startsWith('NOTE')) {
      continue;
    }

    const timeMatch = line.match(timeRegex);
    if (timeMatch) {
      finalizeCue();
      currentStartTime = timeMatch[1];
      currentEndTime = timeMatch[3];
      currentSpeaker = 'Intervenant';
      continue;
    }

    // Détection de balise voix Teams <v Nom Prénom>Texte</v>
    const voiceMatch = line.match(voiceTagRegex) || line.match(legacyVoiceTagRegex);
    if (voiceMatch) {
      currentSpeaker = voiceMatch[1].trim();
      const spokenText = voiceMatch[2].replace(/<\/v>/g, '').trim();
      if (spokenText) {
        currentTextBuffer.push(spokenText);
      }
      continue;
    }

    // Ligne de texte ordinaire
    if (currentStartTime) {
      currentTextBuffer.push(line);
    }
  }

  finalizeCue();

  const totalWords = cues.reduce((acc, c) => acc + c.text.split(/\s+/).filter(Boolean).length, 0);
  const durationSeconds = cues.length > 0 ? cues[cues.length - 1].endSeconds : 0;

  // Calcul des statistiques par intervenant pour l'attribution des locuteurs
  const speakerStatsMap = new Map<string, { interventionsCount: number; totalWords: number }>();
  for (const cue of cues) {
    const words = cue.text.split(/\s+/).filter(Boolean).length;
    const existing = speakerStatsMap.get(cue.speakerName) || { interventionsCount: 0, totalWords: 0 };
    speakerStatsMap.set(cue.speakerName, {
      interventionsCount: existing.interventionsCount + 1,
      totalWords: existing.totalWords + words,
    });
  }

  const AVATAR_GRADIENTS = [
    'from-emerald-600 to-teal-800',
    'from-blue-600 to-indigo-800',
    'from-purple-600 to-violet-800',
    'from-amber-600 to-orange-800',
    'from-rose-600 to-pink-800',
    'from-cyan-600 to-blue-800',
  ];

  const speakers: IdentifiedSpeaker[] = Array.from(speakerStatsMap.entries()).map(([name, stats], idx) => ({
    name,
    interventionsCount: stats.interventionsCount,
    totalWords: stats.totalWords,
    avatarColor: AVATAR_GRADIENTS[idx % AVATAR_GRADIENTS.length],
  }));

  return {
    format: 'text/vtt',
    rawContent: rawVtt,
    cues,
    totalWords,
    durationSeconds,
    speakersCount: speakerStatsMap.size,
    speakers,
  };
}

/**
 * Traduit et qualifie rigoureusement les erreurs Microsoft Graph sur les transcriptions.
 * Distingue explicitement :
 * - Permission absente / consentement non accordé
 * - GraphAccessToTranscriptsDisabled (désactivation au niveau du tenant Teams)
 * - Réunion introuvable ou expirée (HTTP 404)
 * - Limitation de débit (HTTP 429)
 * - Erreur temporaire serveur (HTTP 5xx)
 * Ne confond jamais un 403 avec une simple permission, ni un 404 avec une absence d'enregistrement.
 */
export interface ErrorContext {
  userRole?: UserMeetingRole;
  isExternalOrganizer?: boolean;
  hasTranscriptPermission?: boolean;
  organizerName?: string;
  organizerEmail?: string;
}

/**
 * Traduit et qualifie rigoureusement les erreurs Microsoft Graph sur les transcriptions.
 * Distingue explicitement :
 * - Permission absente / consentement non accordé
 * - GraphAccessToTranscriptsDisabled (désactivation au niveau du tenant Teams)
 * - Refus de partage Teams par l'organisateur (Organizers only)
 * - Réunion issue d'un tenant externe (accès cross-tenant restreint)
 * - Réunion introuvable ou expirée (HTTP 404)
 * - Limitation de débit (HTTP 429)
 * - Erreur temporaire serveur (HTTP 5xx)
 */
export function explainTranscriptError(error: any, context?: ErrorContext): TranscriptErrorDetail {
  const status = error?.statusCode || error?.status || error?.response?.status;
  const message = error?.message || String(error);
  const code = error?.code || error?.error?.code || '';

  // 1. Accès désactivé au niveau du tenant Microsoft Teams (GraphAccessToTranscriptsDisabled)
  if (
    code === 'GraphAccessToTranscriptsDisabled' ||
    message.includes('GraphAccessToTranscriptsDisabled') ||
    message.includes('transcripts is disabled') ||
    message.includes('Access to transcripts is disabled')
  ) {
    return {
      code: 'GRAPH_ACCESS_DISABLED',
      title: 'Accès API aux transcriptions désactivé sur le locataire DGF',
      message: "L'accès de Microsoft Graph aux transcriptions Teams est désactivé au niveau du locataire Microsoft 365 (stratégie de sécurité GraphAccessToTranscriptsDisabled).",
      httpStatus: 403,
      actionRequired: "Un administrateur Teams DGF doit activer l'accès API dans le Teams Admin Center (Meetings > Meeting settings > Transcript API access > Microsoft Graph access = On) ou via PowerShell: Set-CsTeamsMeetingConfiguration -Identity Global -EnableGraphTranscriptAccess $true.",
      teamsPolicyNote: "Par défaut chez Microsoft, cette option de sécurité est désactivée pour tous les locataires d'entreprise.",
    };
  }

  // 2. Attribution des locuteurs désactivée dans Teams (EnableAttributedTranscripts)
  if (
    code === 'AttributedTranscriptDisabled' ||
    message.includes('AttributedTranscriptDisabled') ||
    message.includes('EnableAttributedTranscripts') ||
    message.includes('attribution is disabled')
  ) {
    return {
      code: 'SPEAKER_ATTRIBUTION_DISABLED',
      title: 'Attribution des locuteurs non autorisée',
      message: "L'attribution nominale des interventions aux locuteurs convoqués est désactivée par stratégie dans le locataire Teams.",
      httpStatus: 403,
      actionRequired: "Un administrateur Teams DGF doit activer le paramètre EnableAttributedTranscripts ($true) via PowerShell ou Teams Admin Center.",
      teamsPolicyNote: "La transcription reste lisible mais les balises d'intervenants ne sont pas attribuées nominativement.",
    };
  }

  // 3. Réunion organisée par un locataire externe (Scénario cross-tenant restreint)
  if (
    context?.isExternalOrganizer ||
    code === 'CrossTenantAccessRestricted' ||
    message.includes('CrossTenantAccessRestricted') ||
    message.includes('external tenant')
  ) {
    return {
      code: 'EXTERNAL_TENANT_RESTRICTED',
      title: 'Réunion organisée par un locataire externe',
      message: "L'organisateur de cette réunion appartient à un locataire Microsoft 365 externe. L'accès délégué inter-organisations aux transcriptions Teams est restreint par les politiques de sécurité Entra ID.",
      httpStatus: 403,
      actionRequired: "Pour les réunions organisées par des partenaires ou bailleurs externes (ex: BAD, AFD), demandez à l'organisateur d'exporter le fichier WebVTT ou d'activer une collaboration inter-locataires autorisée (B2B).",
      teamsPolicyNote: "Conformité de gouvernance DGF : Les flux de transcription des locataires externes ne sont pas accessibles via les droits délégués du tenant DGF.",
    };
  }

  // 4. Refus de partage par l'organisateur (Organizers and co-organizers only)
  // Détecté lorsque l'utilisateur est participant invité et que Graph renvoie 403 Forbidden sur /transcripts alors que la permission OnlineMeetingTranscript.Read.All est accordée
  if (
    code === 'SharingRestriction' ||
    code === 'TranscriptSharingRestricted' ||
    code === 'OrganizerOnly' ||
    message.includes('SharingRestriction') ||
    message.includes('OrganizerOnly') ||
    message.includes('Who can view transcripts') ||
    (status === 403 &&
      context?.userRole === 'attendee' &&
      context?.hasTranscriptPermission !== false &&
      (code === 'AccessDenied' || message.includes('Forbidden') || message.includes('AccessDenied')))
  ) {
    return {
      code: 'SHARING_RESTRICTED_BY_ORGANIZER',
      title: "Accès à la transcription restreint par l'organisateur",
      message: "Cette transcription n'est pas accessible avec les droits de votre compte Microsoft 365. Veuillez demander à l'organisateur de vous accorder l'accès.",
      httpStatus: 403,
      actionRequired: `L'organisateur de la réunion (${context?.organizerName || 'Organisateur'}) a configuré les options Teams pour restreindre la consultation aux organisateurs et coorganisateurs. Demandez-lui d'autoriser le partage sur « Tous les participants » (Everyone) dans les options de la réunion Teams.`,
      teamsPolicyNote: "Paramètres de réunion Teams : « Qui peut afficher la transcription » configuré sur « Organisateurs et coorganisateurs ».",
    };
  }

  // 5. Permission déléguée absente ou consentement non accordé
  if (
    context?.hasTranscriptPermission === false ||
    message.includes('OnlineMeetingTranscript.Read.All') ||
    message.includes('AADSTS65001') ||
    message.includes('consent') ||
    (status === 403 && (message.includes('Forbidden') || message.includes('AccessDenied') || code === 'AccessDenied'))
  ) {
    return {
      code: 'PERMISSION_MISSING',
      title: 'Autorisation administrateur en attente',
      message: "L'autorisation déléguée 'OnlineMeetingTranscript.Read.All' requiert le consentement préalable de l'administrateur du locataire DGF. Aucune tentative de contournement n'est autorisée.",
      httpStatus: 403,
      actionRequired: "Un administrateur du locataire Microsoft 365 DGF doit accorder le consentement d'administrateur pour la permission déléguée 'OnlineMeetingTranscript.Read.All'.",
      missingPermission: 'OnlineMeetingTranscript.Read.All',
    };
  }

  // 6. Réunion introuvable ou expirée dans Microsoft Graph (HTTP 404)
  if (status === 404 || code === 'ItemNotFound' || code === 'MeetingNotFound' || message.includes('ItemNotFound')) {
    return {
      code: 'MEETING_NOT_FOUND',
      title: 'Réunion Teams introuvable ou expirée',
      message: "L'identifiant de la réunion est introuvable ou la session a expiré dans le service Microsoft Teams (404 Not Found).",
      httpStatus: 404,
      actionRequired: "Vérifiez que le lien de réunion est valide et que l'événement n'a pas été supprimé du calendrier.",
    };
  }

  // 7. Limitation de débit / Quota Microsoft Graph (HTTP 429)
  if (status === 429 || message.includes('TooManyRequests') || code === 'TooManyRequests') {
    const retryAfter = error?.headers?.get?.('Retry-After') || 10;
    return {
      code: 'RATE_LIMIT_EXCEEDED',
      title: 'Quota Microsoft Graph temporairement dépassé (HTTP 429)',
      message: `Trop de requêtes adressées à l'API Microsoft Graph. Veuillez patienter ${retryAfter} secondes.`,
      httpStatus: 429,
      actionRequired: `Attendez la fin de la période d'attente imposée par Microsoft (${retryAfter}s) avant de réinterroger.`,
      retryAfterSeconds: parseInt(String(retryAfter), 10) || 10,
    };
  }

  // 8. Session expirée ou non autorisée (HTTP 401)
  if (status === 401 || message.includes('Unauthorized') || code === 'InvalidAuthenticationToken') {
    return {
      code: 'UNAUTHORIZED',
      title: 'Session Microsoft 365 expirée',
      message: 'Le jeton de session Entra ID est expiré ou invalide.',
      httpStatus: 401,
      actionRequired: 'Veuillez vous reconnecter via la mire institutionnelle Microsoft.',
    };
  }

  // 9. Erreur temporaire de service (HTTP 500, 502, 503, 504)
  if (status && status >= 500) {
    return {
      code: 'TRANSIENT_SERVICE_ERROR',
      title: 'Service Microsoft Graph temporairement indisponible',
      message: `Les serveurs Microsoft ont renvoyé une erreur temporaire (HTTP ${status}).`,
      httpStatus: status,
      actionRequired: 'Veuillez renouveler votre demande dans quelques instants.',
    };
  }

  return {
    code: 'GRAPH_TRANSCRIPT_ERROR',
    title: 'Erreur technique Microsoft Graph',
    message,
    httpStatus: status,
    actionRequired: 'Consultez le journal de diagnostic ou renouvelez votre tentative ultérieurement.',
  };
}

export interface MeetingResolutionOptions {
  organizerEmail?: string;
  userRole?: UserMeetingRole;
  isExternalOrganizer?: boolean;
  organizerName?: string;
  hasTranscriptPermission?: boolean;
}

/**
 * 2. Tente de résoudre l'ID Teams (onlineMeetingId) à partir de l'URL JoinWebUrl de la réunion Outlook.
 * Prend en charge :
 * - SCÉNARIO A : Utilisateur organisateur (/me/onlineMeetings)
 * - SCÉNARIO B : Utilisateur participant invité (/me/onlineMeetings puis /users/{organizerEmail}/onlineMeetings si même tenant)
 * - Détection et refus propre pour les réunions de locataires externes (cross-tenant)
 */
export async function resolveOnlineMeetingId(
  client: Client,
  joinWebUrl: string,
  options?: MeetingResolutionOptions
): Promise<{ onlineMeetingId: string | null; error?: TranscriptErrorDetail }> {
  try {
    // Si l'organisateur est formellement identifié comme externe au locataire DGF
    if (options?.isExternalOrganizer) {
      return {
        onlineMeetingId: null,
        error: explainTranscriptError(new Error('CrossTenantAccessRestricted'), {
          isExternalOrganizer: true,
          organizerName: options?.organizerName,
          userRole: options?.userRole,
        }),
      };
    }

    const cleanUrl = joinWebUrl.split('?')[0];

    // Tentative 1 : /me/onlineMeetings (Fonctionne si organisateur ou réunion découverte dans la session utilisateur)
    const meResponse = await client
      .api('/me/onlineMeetings')
      .filter(`JoinWebUrl eq '${joinWebUrl}'`)
      .select('id,subject,joinWebUrl')
      .get()
      .catch(async () => {
        // Repli avec l'URL nettoyée si le premier filtre strict échoue
        return client
          .api('/me/onlineMeetings')
          .filter(`JoinWebUrl eq '${cleanUrl}'`)
          .select('id,subject,joinWebUrl')
          .get()
          .catch(() => null);
      });

    if (meResponse && Array.isArray(meResponse.value) && meResponse.value.length > 0) {
      return { onlineMeetingId: meResponse.value[0].id };
    }

    // Tentative 2 (SCÉNARIO B — Participant invité) :
    // Si l'utilisateur est participant et que l'organisateur appartient au même locataire DGF
    if (options?.organizerEmail && options.userRole === 'attendee') {
      try {
        const userMeetingsResponse = await client
          .api(`/users/${options.organizerEmail}/onlineMeetings`)
          .filter(`JoinWebUrl eq '${joinWebUrl}'`)
          .select('id,subject,joinWebUrl')
          .get()
          .catch(async () => {
            return client
              .api(`/users/${options.organizerEmail}/onlineMeetings`)
              .filter(`JoinWebUrl eq '${cleanUrl}'`)
              .select('id,subject,joinWebUrl')
              .get()
              .catch(() => null);
          });

        if (userMeetingsResponse && Array.isArray(userMeetingsResponse.value) && userMeetingsResponse.value.length > 0) {
          return { onlineMeetingId: userMeetingsResponse.value[0].id };
        }
      } catch (userRouteErr) {
        console.warn('[Attendee onlineMeetings resolution notice]', userRouteErr);
      }
    }

    // Comportement explicite lorsque la réunion ne peut pas être résolue
    return {
      onlineMeetingId: null,
      error: {
        code: 'MEETING_UNRESOLVED',
        title: 'Réunion Teams non résolue via Microsoft Graph',
        message: "L'identifiant de la session Teams n'a pas pu être résolu à partir de l'URL de connexion. Si vous êtes invité à une réunion organisée par un tiers ou un compte externe, la découverte dépend des stratégies d'accès et de partage du locataire.",
        actionRequired: "Assurez-vous que la réunion a été organisée par un compte du tenant DGF ou que vous disposez des droits d'accès à la session Teams.",
      },
    };
  } catch (err: any) {
    return {
      onlineMeetingId: null,
      error: explainTranscriptError(err, {
        userRole: options?.userRole,
        isExternalOrganizer: options?.isExternalOrganizer,
        organizerName: options?.organizerName,
      }),
    };
  }
}

/**
 * 3. Récupère la liste des métadonnées de transcription pour une réunion Teams.
 * Interroge /me/onlineMeetings/{id}/transcripts ou /users/{organizerEmail}/onlineMeetings/{id}/transcripts selon le rôle.
 */
export async function fetchMeetingTranscripts(
  client: Client,
  onlineMeetingId: string,
  options?: MeetingResolutionOptions
): Promise<{ transcripts: TranscriptMetadata[]; error?: TranscriptErrorDetail }> {
  try {
    // Tentative 1 : Route standard /me
    try {
      const response = await client
        .api(`/me/onlineMeetings/${onlineMeetingId}/transcripts`)
        .select('id,createdDateTime,meetingId,transcriptContentUrl,meetingOrganizerId')
        .get();

      if (response && Array.isArray(response.value)) {
        const list: TranscriptMetadata[] = response.value.map((item: any) => ({
          id: item.id,
          meetingId: item.meetingId || onlineMeetingId,
          createdDateTime: item.createdDateTime || new Date().toISOString(),
          meetingOrganizerId: item.meetingOrganizerId,
          transcriptContentUrl: item.transcriptContentUrl || `/me/onlineMeetings/${onlineMeetingId}/transcripts/${item.id}/content`,
        }));
        return { transcripts: list };
      }
    } catch (meErr: any) {
      const meStatus = meErr?.statusCode || meErr?.status;

      // Si 403 Forbidden sur /me alors que la permission est accordée et l'utilisateur est participant :
      // il s'agit d'un refus de partage par l'organisateur Teams (Organizers only)
      if (meStatus === 403 && options?.userRole === 'attendee' && options?.hasTranscriptPermission !== false) {
        return {
          transcripts: [],
          error: explainTranscriptError(meErr, {
            userRole: 'attendee',
            hasTranscriptPermission: true,
            organizerName: options?.organizerName,
            organizerEmail: options?.organizerEmail,
          }),
        };
      }

      // Si 404 (non trouvé sous /me) et organisateur DGF tiers identifié
      if ((meStatus === 404 || meErr?.code === 'ItemNotFound') && options?.organizerEmail && options.userRole === 'attendee') {
        const userResponse = await client
          .api(`/users/${options.organizerEmail}/onlineMeetings/${onlineMeetingId}/transcripts`)
          .select('id,createdDateTime,meetingId,transcriptContentUrl,meetingOrganizerId')
          .get();

        if (userResponse && Array.isArray(userResponse.value)) {
          const list: TranscriptMetadata[] = userResponse.value.map((item: any) => ({
            id: item.id,
            meetingId: item.meetingId || onlineMeetingId,
            createdDateTime: item.createdDateTime || new Date().toISOString(),
            meetingOrganizerId: item.meetingOrganizerId,
            transcriptContentUrl: item.transcriptContentUrl || `/users/${options.organizerEmail}/onlineMeetings/${onlineMeetingId}/transcripts/${item.id}/content`,
          }));
          return { transcripts: list };
        }
      } else {
        throw meErr;
      }
    }

    return { transcripts: [] };
  } catch (err: any) {
    return {
      transcripts: [],
      error: explainTranscriptError(err, {
        userRole: options?.userRole,
        hasTranscriptPermission: options?.hasTranscriptPermission,
        organizerName: options?.organizerName,
        organizerEmail: options?.organizerEmail,
      }),
    };
  }
}

/**
 * 4. Récupère le contenu textuel / VTT d'une transcription Teams
 * Endpoint : GET /me/onlineMeetings/{onlineMeetingId}/transcripts/{transcriptId}/content?$format=text/vtt
 * ou /users/{organizerEmail}/... pour les réunions de tiers autorisées.
 */
export async function fetchTranscriptContent(
  client: Client,
  onlineMeetingId: string,
  transcriptId: string,
  options?: MeetingResolutionOptions
): Promise<{ content: string; error?: TranscriptErrorDetail }> {
  try {
    let rawResponse: any;
    try {
      rawResponse = await client
        .api(`/me/onlineMeetings/${onlineMeetingId}/transcripts/${transcriptId}/content`)
        .query({ $format: 'text/vtt' })
        .responseType(ResponseType.TEXT)
        .get();
    } catch (meErr: any) {
      const meStatus = meErr?.statusCode || meErr?.status;
      if ((meStatus === 404 || meErr?.code === 'ItemNotFound') && options?.organizerEmail && options.userRole === 'attendee') {
        rawResponse = await client
          .api(`/users/${options.organizerEmail}/onlineMeetings/${onlineMeetingId}/transcripts/${transcriptId}/content`)
          .query({ $format: 'text/vtt' })
          .responseType(ResponseType.TEXT)
          .get();
      } else {
        throw meErr;
      }
    }

    const textContent = typeof rawResponse === 'string' ? rawResponse : String(rawResponse);
    return { content: textContent };
  } catch (err: any) {
    return {
      content: '',
      error: explainTranscriptError(err, {
        userRole: options?.userRole,
        hasTranscriptPermission: options?.hasTranscriptPermission,
        organizerName: options?.organizerName,
        organizerEmail: options?.organizerEmail,
      }),
    };
  }
}

/**
 * 5. Fonction principale : Orchestre la vérification complète d'une réunion avec gestion du cache
 * Strictement conforme aux limites :
 * - Ne prétend jamais capturer d'audio
 * - Ne déclare "disponible" qu'en cas de confirmation réelle par Microsoft Graph
 * - Distingue explicitement les erreurs 403 (droit vs tenant), les 404, et les collections vides 200 OK
 */
export async function checkMeetingTranscriptStatus(
  client: Client | null,
  meeting: MeetingItem,
  forceRefresh: boolean = false,
  currentUser?: DgfUserProfile | null,
  hasTranscriptPermission: boolean = true
): Promise<TranscriptCheckResult> {
  const temporalStatus = getMeetingTemporalStatus(meeting.start, meeting.end);
  const cacheKey = `${meeting.id}_${currentUser?.userPrincipalName || 'anon'}`;

  const userRole: UserMeetingRole = meeting.userRole || detectUserMeetingRole(meeting, currentUser || null);
  const isExternal = meeting.isExternalOrganizer ?? isOrganizerExternal(meeting.organizer.email, currentUser?.userPrincipalName || currentUser?.mail);

  // 1. Vérification du cache si rafraîchissement non forcé
  if (!forceRefresh && memoryCache.has(cacheKey)) {
    const entry = memoryCache.get(cacheKey)!;
    if (Date.now() < entry.expiresAt) {
      return {
        ...entry.result,
        temporalStatus,
        userRole,
        isExternalOrganizer: isExternal,
        cached: true,
      };
    } else {
      memoryCache.delete(cacheKey);
    }
  }

  // 2. Vérification s'il s'agit d'une réunion Teams
  if (!meeting.isOnlineMeeting || !meeting.joinWebUrl) {
    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus,
      status: 'transcript_unavailable',
      userRole,
      isExternalOrganizer: isExternal,
      accessibility: 'no_transcript',
      checkedAt: new Date().toISOString(),
      error: {
        code: 'NOT_A_TEAMS_MEETING',
        title: 'Séance non éligible à la transcription Teams native',
        message: "Cet événement d'agenda est présentiel ou ne dispose pas d'un lien Teams officiel.",
        actionRequired: "Pour les séances présentielles, utilisez la saisie manuelle ou le futur module d'enregistrement audio de salle.",
      },
    };
    return result;
  }

  // 3. Traitement selon la temporalité de la réunion (sans prétention de capture par ReadMeeting.ai)
  if (temporalStatus === 'scheduled') {
    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus: 'scheduled',
      status: 'scheduled',
      userRole,
      isExternalOrganizer: isExternal,
      accessibility: 'pending_check',
      checkedAt: new Date().toISOString(),
      error: {
        code: 'MEETING_SCHEDULED',
        title: 'Réunion programmée',
        message: 'Cette séance Teams est planifiée dans le futur. La transcription Teams ne peut être générée que pendant ou après la tenue effective de la réunion.',
        actionRequired: "Attendez la fin de la séance pour interroger la disponibilité de la transcription.",
      },
    };
    return result;
  }

  if (temporalStatus === 'in_progress') {
    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus: 'in_progress',
      status: 'in_progress',
      userRole,
      isExternalOrganizer: isExternal,
      accessibility: 'pending_check',
      checkedAt: new Date().toISOString(),
      error: {
        code: 'MEETING_IN_PROGRESS',
        title: 'Réunion actuellement en cours',
        message: "La séance Teams est en cours de déroulement. Microsoft Teams génère la transcription en direct si l'organisateur ou un participant autorisé l'a activée dans Teams. ReadMeeting.ai interrogera Microsoft Graph une fois la séance achevée.",
        actionRequired: "Attendez la fin de la réunion pour récupérer le flux de transcription complet.",
      },
    };
    return result;
  }

  // 4. Réunion terminée : vérification auprès de Microsoft Graph
  if (!client) {
    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus: 'completed',
      status: 'transient_error',
      userRole,
      isExternalOrganizer: isExternal,
      accessibility: 'pending_check',
      checkedAt: new Date().toISOString(),
      error: {
        code: 'CLIENT_NOT_INITIALIZED',
        title: 'Client Microsoft Graph non connecté',
        message: 'Une session Microsoft Entra ID active est requise pour interroger les services Teams.',
        actionRequired: 'Connectez-vous avec votre compte professionnel Microsoft 365.',
      },
    };
    return result;
  }

  const resolutionOptions: MeetingResolutionOptions = {
    organizerEmail: meeting.organizer.email,
    organizerName: meeting.organizer.name,
    userRole,
    isExternalOrganizer: isExternal,
    hasTranscriptPermission,
  };

  // 4.a Résolution de l'ID Teams onlineMeeting (Prise en charge Organisateur ET Participant invité)
  const resolveResult = await resolveOnlineMeetingId(client, meeting.joinWebUrl, resolutionOptions);
  if (!resolveResult.onlineMeetingId) {
    const errorCode = resolveResult.error?.code;
    const status: TranscriptStatus =
      errorCode === 'EXTERNAL_TENANT_RESTRICTED'
        ? 'external_tenant_restricted'
        : errorCode === 'PERMISSION_MISSING'
        ? 'permission_missing'
        : errorCode === 'GRAPH_ACCESS_DISABLED'
        ? 'tenant_access_disabled'
        : 'meeting_unresolved';

    const accessibility: TranscriptAccessibilityStatus =
      status === 'external_tenant_restricted'
        ? 'external_tenant'
        : status === 'permission_missing'
        ? 'permission_missing'
        : 'unresolved';

    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus: 'completed',
      status,
      userRole,
      isExternalOrganizer: isExternal,
      accessibility,
      checkedAt: new Date().toISOString(),
      error: resolveResult.error || {
        code: 'MEETING_UNRESOLVED',
        title: 'Identifiant Teams introuvable',
        message: "Impossible de relier l'événement Outlook à la session Microsoft Teams correspondante.",
        actionRequired: "Vérifiez que le lien Teams est valide et que vous disposez des droits d'accès à la session.",
      },
    };

    memoryCache.set(cacheKey, {
      result,
      expiresAt: Date.now() + CACHE_TTL_ERROR_MS,
    });
    return result;
  }

  const onlineMeetingId = resolveResult.onlineMeetingId;

  // 4.b Interrogation de l'endpoint /transcripts (avec prise en compte du contexte d'invité)
  const transcriptsResult = await fetchMeetingTranscripts(client, onlineMeetingId, resolutionOptions);

  if (transcriptsResult.error) {
    const errCode = transcriptsResult.error.code;
    let status: TranscriptStatus = 'transient_error';
    let accessibility: TranscriptAccessibilityStatus = 'unresolved';

    if (errCode === 'SHARING_RESTRICTED_BY_ORGANIZER') {
      status = 'sharing_restricted_by_organizer';
      accessibility = 'restricted_by_sharing';
    } else if (errCode === 'EXTERNAL_TENANT_RESTRICTED') {
      status = 'external_tenant_restricted';
      accessibility = 'external_tenant';
    } else if (errCode === 'GRAPH_ACCESS_DISABLED') {
      status = 'tenant_access_disabled';
      accessibility = 'unresolved';
    } else if (errCode === 'PERMISSION_MISSING') {
      status = 'permission_missing';
      accessibility = 'permission_missing';
    } else if (errCode === 'MEETING_NOT_FOUND') {
      status = 'meeting_not_found';
      accessibility = 'unresolved';
    } else if (errCode === 'RATE_LIMIT_EXCEEDED') {
      status = 'transient_error';
      accessibility = 'pending_check';
    }

    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus: 'completed',
      status,
      onlineMeetingId,
      userRole,
      isExternalOrganizer: isExternal,
      accessibility,
      checkedAt: new Date().toISOString(),
      error: transcriptsResult.error,
    };

    const ttl =
      errCode === 'RATE_LIMIT_EXCEEDED'
        ? (transcriptsResult.error.retryAfterSeconds || 10) * 1000
        : CACHE_TTL_ERROR_MS;

    memoryCache.set(cacheKey, {
      result,
      expiresAt: Date.now() + ttl,
    });
    return result;
  }

  // 4.c Collection de transcriptions vide (Réponse HTTP 200 OK avec tableau vide [])
  if (transcriptsResult.transcripts.length === 0) {
    const result: TranscriptCheckResult = {
      meetingId: meeting.id,
      meetingTitle: meeting.title,
      temporalStatus: 'completed',
      status: 'transcript_unavailable',
      onlineMeetingId,
      userRole,
      isExternalOrganizer: isExternal,
      accessibility: 'no_transcript',
      checkedAt: new Date().toISOString(),
      error: {
        code: 'EMPTY_TRANSCRIPTS',
        title: 'Aucune transcription disponible pour cette réunion',
        message: "La réunion Teams s'est achevée sans que la transcription en direct n'ait été activée dans Teams. Aucune transcription n'a été enregistrée pour cette séance.",
        actionRequired: "La transcription Teams n'est pas automatique. Pour vos prochaines réunions, activez « Démarrer la transcription » dès le début de séance dans Teams.",
        teamsPolicyNote: "La disponibilité dépend des stratégies Teams (AllowTranscription) attribuées aux organisateurs et participants.",
      },
    };

    memoryCache.set(cacheKey, {
      result,
      expiresAt: Date.now() + CACHE_TTL_ERROR_MS,
    });
    return result;
  }

  // 4.d Transcription existante confirmée par Microsoft Graph
  const primaryMetadata = transcriptsResult.transcripts[0];
  let parsedContent: TranscriptContent | undefined = undefined;

  // Tentative de récupération du contenu textuel WebVTT
  const contentResult = await fetchTranscriptContent(client, onlineMeetingId, primaryMetadata.id, resolutionOptions);
  if (contentResult.content) {
    parsedContent = parseVttTranscript(contentResult.content);
  }

  // Construction de la piste d'audit institutionnelle DGF
  const auditTrail: TranscriptRetrievalAudit = {
    meetingId: meeting.id,
    meetingTitle: meeting.title,
    onlineMeetingId,
    transcriptId: primaryMetadata.id,
    requestedByUpn: currentUser?.userPrincipalName || currentUser?.mail || 'collaborateur@finances.gouv.ci',
    requestedByName: currentUser?.displayName || 'Collaborateur DGF',
    userRole,
    organizerUpn: meeting.organizer.email,
    organizerName: meeting.organizer.name,
    isExternalOrganizer: isExternal,
    retrievedAt: new Date().toISOString(),
    cuesCount: parsedContent?.cues.length || 0,
    totalWords: parsedContent?.totalWords || 0,
  };

  const result: TranscriptCheckResult = {
    meetingId: meeting.id,
    meetingTitle: meeting.title,
    temporalStatus: 'completed',
    status: 'transcript_available',
    onlineMeetingId,
    userRole,
    isExternalOrganizer: isExternal,
    accessibility: 'accessible',
    metadata: primaryMetadata,
    content: parsedContent,
    auditTrail,
    checkedAt: new Date().toISOString(),
  };

  memoryCache.set(cacheKey, {
    result,
    expiresAt: Date.now() + CACHE_TTL_SUCCESS_MS,
  });

  return result;
}

/**
 * 6. Garde-fou institutionnel de validation humaine pour le futur traitement documentaire
 * Un compte rendu officiel ne doit JAMAIS être validé automatiquement sans signature humaine.
 */
export function enforceHumanValidationGate(validation: HumanValidationAudit): {
  canPublish: boolean;
  blockReason?: string;
} {
  if (!validation.isApproved) {
    return {
      canPublish: false,
      blockReason: 'Le compte rendu est encore au stade de projet ou de relecture et n a pas été approuvé.',
    };
  }

  if (validation.validationStatus !== 'approved') {
    return {
      canPublish: false,
      blockReason: `Statut non conforme (${validation.validationStatus}). L approbation institutionnelle formelle est requise.`,
    };
  }

  if (!validation.reviewedByUpn || !validation.reviewedByName) {
    return {
      canPublish: false,
      blockReason: "L'identité institutionnelle (UPN et Nom) du rapporteur ou président signataire est obligatoire.",
    };
  }

  return { canPublish: true };
}

/**
 * Purge manuelle du cache des transcriptions
 */
export function purgeTranscriptCache(meetingId?: string): void {
  if (meetingId) {
    memoryCache.delete(meetingId);
  } else {
    memoryCache.clear();
  }
}
