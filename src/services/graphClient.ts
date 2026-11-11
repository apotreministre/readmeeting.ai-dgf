import { AccountInfo, InteractionRequiredAuthError, IPublicClientApplication } from '@azure/msal-browser';
import { Client } from '@microsoft/microsoft-graph-client';
import { loginRequest } from '../config/msalConfig';
import { DgfUserProfile } from '../types/auth';
import { MeetingItem } from '../types/meeting';

export interface GraphErrorExplanation {
  code: string;
  title: string;
  description: string;
  actionRequired: string;
}

export function explainEntraIdError(error: unknown): GraphErrorExplanation {
  const errStr = String(error);
  const errMsg = (error as { message?: string })?.message || errStr;

  if (errMsg.includes('AADSTS50011') || errMsg.includes('reply URL')) {
    return {
      code: 'AADSTS50011',
      title: 'URI de redirection non reconnue',
      description: "L'URL de l'application active dans votre navigateur n'est pas configurée dans les Redirect URIs de l'application Entra ID.",
      actionRequired: "Ouvrez le portail Entra ID > Inscriptions d'applications > Authentification > Plateforme SPA > Ajoutez l'URI exacte indiquée dans la console de diagnostic.",
    };
  }

  if (errMsg.includes('AADSTS65001') || errMsg.includes('consent')) {
    return {
      code: 'AADSTS65001',
      title: 'Consentement administrateur requis',
      description: "Les autorisations demandées (notamment pour les transcriptions de réunions Teams) nécessitent l'accord préalable d'un administrateur du tenant.",
      actionRequired: "Un administrateur M365 DGF doit se rendre sur l'application dans Entra ID et cliquer sur « Accorder un consentement d'administrateur ».",
    };
  }

  if (errMsg.includes('AADSTS700016') || errMsg.includes('was not found in the directory')) {
    return {
      code: 'AADSTS700016',
      title: "Application non trouvée dans l'annuaire",
      description: "Le Client ID ou le Tenant ID renseigné n'existe pas ou ne correspond à aucune inscription active dans ce locataire.",
      actionRequired: "Vérifiez que le Client ID et le Tenant ID correspondent fidèlement aux GUIDs générés dans le portail Microsoft Entra.",
    };
  }

  if (errMsg.includes('redirect_bridge_timeout') || errMsg.includes('timedOut') || errMsg.includes('timeout')) {
    return {
      code: 'POPUP_TIMEOUT',
      title: 'Délai d attente dépassé pour la mire Microsoft',
      description: "La fenêtre de connexion Microsoft n'a pas répondu à temps ou a été fermée avant la validation des identifiants (délai dépassé). Le mode Legacy Polling et le mode Redirection sont désormais activés pour fiabiliser la session.",
      actionRequired: "Désactivez les bloqueurs de fenêtres pop-up, ou utilisez l'option « Connexion par Redirection » pour vous authentifier directement en pleine page.",
    };
  }

  if (errMsg.includes('user_cancelled')) {
    return {
      code: 'USER_CANCELLED',
      title: 'Authentification interrompue',
      description: "La fenêtre pop-up d'authentification a été fermée avant la fin de la validation des identifiants.",
      actionRequired: "Cliquez à nouveau sur « Se connecter » pour finaliser la saisie de vos identifiants Microsoft 365.",
    };
  }

  if (errMsg.includes('interaction_in_progress')) {
    return {
      code: 'INTERACTION_IN_PROGRESS',
      title: 'Opération d authentification déjà en cours',
      description: "Une demande de connexion est déjà en attente dans une autre fenêtre ou un autre onglet.",
      actionRequired: "Fermez les fenêtres de connexion ouvertes ou attendez quelques secondes avant de relancer.",
    };
  }

  if (errMsg.includes('interaction_required') || errMsg.includes('consent_required')) {
    return {
      code: 'INTERACTION_REQUIRED',
      title: 'Action requise de l utilisateur',
      description: 'Une validation MFA (authentification multifacteur) ou un consentement utilisateur est nécessaire.',
      actionRequired: 'Cliquez à nouveau sur « Se connecter avec Microsoft Entra ID » pour valider vos informations.',
    };
  }

  return {
    code: 'GRAPH_API_ERROR',
    title: 'Erreur Microsoft Graph / Entra ID',
    description: errMsg,
    actionRequired: "Consultez les détails de l'erreur dans la console de diagnostic ou vérifiez vos autorisations dans Microsoft Entra ID.",
  };
}

export const TRANSCRIPT_REQUIRED_SCOPES = [
  'User.Read',
  'Calendars.Read',
  'OnlineMeetings.Read',
  'OnlineMeetingTranscript.Read.All',
];

export interface TranscriptTokenResult {
  token: string | null;
  hasPermission: boolean;
  status: 'granted' | 'pending_admin_consent' | 'interaction_required' | 'error';
  grantedScopes?: string[];
  error?: string;
}

/**
 * ÉTAPE 1 & 2 : Vérifie les autorisations et acquiert le jeton Graph pour les transcriptions
 * Privilégie strictement acquireTokenSilent, gère InteractionRequiredAuthError
 * et ne divulgue jamais la valeur secrète du jeton dans les journaux ou l'interface.
 */
export async function getTranscriptAccessToken(
  msal: IPublicClientApplication,
  account: AccountInfo
): Promise<TranscriptTokenResult> {
  try {
    const silentResponse = await msal.acquireTokenSilent({
      account,
      scopes: TRANSCRIPT_REQUIRED_SCOPES,
    });

    const granted = (silentResponse.scopes || []).map((s) => s.toLowerCase());
    const hasTranscriptScope = granted.includes('onlinemeetingtranscript.read.all');

    if (hasTranscriptScope) {
      return {
        token: silentResponse.accessToken,
        hasPermission: true,
        status: 'granted',
        grantedScopes: silentResponse.scopes,
      };
    }

    return {
      token: null,
      hasPermission: false,
      status: 'pending_admin_consent',
      error: "Autorisation administrateur en attente : 'OnlineMeetingTranscript.Read.All' absente du jeton.",
      grantedScopes: silentResponse.scopes,
    };
  } catch (err: any) {
    if (
      err instanceof InteractionRequiredAuthError ||
      err?.errorCode === 'consent_required' ||
      err?.errorCode === 'interaction_required' ||
      err?.message?.includes('AADSTS65001') ||
      err?.message?.includes('AADSTS90094')
    ) {
      return {
        token: null,
        hasPermission: false,
        status: 'pending_admin_consent',
        error: "Autorisation administrateur en attente : le consentement du locataire DGF n'est pas encore accordé.",
      };
    }

    return {
      token: null,
      hasPermission: false,
      status: 'error',
      error: err?.message || String(err),
    };
  }
}

/**
 * Crée un client Microsoft Graph spécifiquement configuré pour les transcriptions Teams.
 * Émet une erreur typée avec code PERMISSION_MISSING si le consentement administrateur n'est pas actif.
 */
export function createTranscriptGraphClient(
  msal: IPublicClientApplication,
  account: AccountInfo
): Client {
  return Client.init({
    authProvider: async (done) => {
      const tokenRes = await getTranscriptAccessToken(msal, account);
      if (tokenRes.token && tokenRes.hasPermission) {
        done(null, tokenRes.token);
      } else {
        const error = new Error(
          tokenRes.error || "Autorisation administrateur en attente pour 'OnlineMeetingTranscript.Read.All'."
        );
        (error as any).code = 'PERMISSION_MISSING';
        (error as any).statusCode = 403;
        done(error, null);
      }
    },
  });
}

export function createGraphClient(
  msal: IPublicClientApplication,
  account: AccountInfo,
  customScopes?: string[]
): Client {
  const targetScopes = customScopes && customScopes.length > 0 ? customScopes : loginRequest.scopes;
  return Client.init({
    authProvider: async (done) => {
      try {
        const response = await msal.acquireTokenSilent({
          scopes: targetScopes,
          account,
        });
        done(null, response.accessToken);
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
          try {
            const popupResponse = await msal.acquireTokenPopup({
              scopes: targetScopes,
              account,
            });
            done(null, popupResponse.accessToken);
          } catch (popupErr) {
            done(popupErr as Error, null);
          }
        } else {
          done(error as Error, null);
        }
      }
    },
  });
}

export async function fetchGraphUserProfile(
  msal: IPublicClientApplication,
  account: AccountInfo
): Promise<DgfUserProfile> {
  const client = createGraphClient(msal, account);
  const data = await client
    .api('/me')
    .select('id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,officeLocation,department')
    .get();

  return {
    id: data.id || undefined,
    displayName: data.displayName || account.name || 'Collaborateur DGF',
    givenName: data.givenName || undefined,
    surname: data.surname || undefined,
    mail: data.mail || undefined,
    userPrincipalName: data.userPrincipalName || account.username,
    jobTitle: data.jobTitle || undefined,
    officeLocation: data.officeLocation || undefined,
    department: data.department || undefined,
  };
}

export interface OutlookCalendarEvent {
  id: string;
  subject?: string;
  bodyPreview?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  organizer?: { emailAddress?: { name?: string; address?: string } };
  location?: { displayName?: string };
  isOnlineMeeting?: boolean;
  onlineMeeting?: { joinUrl?: string };
  onlineMeetingProvider?: string;
  webLink?: string;
  attendees?: Array<{
    type?: string;
    status?: { response?: string; time?: string };
    emailAddress?: { name?: string; address?: string };
  }>;
}

export async function fetchRealCalendarMeetings(
  msal: IPublicClientApplication,
  account: AccountInfo,
  startIso: string,
  endIso: string
): Promise<MeetingItem[]> {
  const client = createGraphClient(msal, account);

  let rawEvents: OutlookCalendarEvent[] = [];

  // Première page d'événements du calendrier avec header UTC normalisé
  const firstPageResponse = await client
    .api('/me/calendarView')
    .header('Prefer', 'outlook.timezone="UTC"')
    .query({
      startDateTime: startIso,
      endDateTime: endIso,
      $top: 25,
      $orderby: 'start/dateTime asc',
      $select: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,onlineMeetingProvider,webLink',
    })
    .get();

  if (firstPageResponse && Array.isArray(firstPageResponse.value)) {
    rawEvents = rawEvents.concat(firstPageResponse.value);
  }

  // Traitement séquentiel des pages suivantes via @odata.nextLink (limité à 4 pages / 100 événements max)
  let nextLink: string | undefined = firstPageResponse?.['@odata.nextLink'];
  let pageCount = 1;
  const MAX_PAGES = 4;

  while (nextLink && pageCount < MAX_PAGES) {
    pageCount++;
    try {
      const nextPageResponse = await client.api(nextLink).header('Prefer', 'outlook.timezone="UTC"').get();
      if (nextPageResponse && Array.isArray(nextPageResponse.value)) {
        rawEvents = rawEvents.concat(nextPageResponse.value);
      }
      nextLink = nextPageResponse?.['@odata.nextLink'];
    } catch (pageErr) {
      console.warn('[Microsoft Graph Pagination Warning]', pageErr);
      break;
    }
  }

  const now = new Date();

  return rawEvents.map((evt) => {
    const subject = evt.subject || 'Réunion sans objet';
    const isTeams = Boolean(
      evt.isOnlineMeeting ||
      evt.onlineMeetingProvider === 'teamsForBusiness' ||
      evt.onlineMeeting?.joinUrl?.includes('teams.microsoft.com') ||
      evt.webLink?.includes('teams.microsoft.com')
    );

    const isDgf =
      subject.toUpperCase().includes('DGF') ||
      subject.toUpperCase().includes('FINANCEMENT') ||
      subject.toUpperCase().includes('PROJET') ||
      subject.toUpperCase().includes('BUDGET') ||
      Boolean(evt.organizer?.emailAddress?.address?.includes('finances.gouv.ci'));

    const attendees = (evt.attendees || []).map((att) => ({
      name: att.emailAddress?.name || 'Participant',
      email: att.emailAddress?.address || '',
      type: att.type || 'required',
      responseStatus: att.status?.response,
      isExternal: !att.emailAddress?.address?.includes('gouv.ci'),
    }));

    // Formatage UTC normalisé ISO
    const startDate = evt.start?.dateTime
      ? evt.start.dateTime.endsWith('Z')
        ? evt.start.dateTime
        : `${evt.start.dateTime}Z`
      : new Date().toISOString();

    const endDate = evt.end?.dateTime
      ? evt.end.dateTime.endsWith('Z')
        ? evt.end.dateTime
        : `${evt.end.dateTime}Z`
      : new Date().toISOString();

    const isPast = new Date(endDate) < now;

    const userUpn = (account.username || '').toLowerCase();
    const userName = (account.name || '').toLowerCase();
    const orgEmail = (evt.organizer?.emailAddress?.address || '').toLowerCase();
    const orgName = (evt.organizer?.emailAddress?.name || '').toLowerCase();

    let userRole: 'organizer' | 'coorganizer' | 'attendee' | 'undetermined' = 'undetermined';
    if (orgEmail === userUpn || (orgName && userName && (orgName.includes(userName) || userName.includes(orgName)))) {
      userRole = 'organizer';
    } else {
      const matchAttendee = attendees.find((a) => a.email.toLowerCase() === userUpn || (userName && a.name.toLowerCase().includes(userName)));
      if (matchAttendee) {
        userRole = 'attendee';
      }
    }

    const isExternalOrg = Boolean(orgEmail && !orgEmail.includes('finances.gouv.ci') && !orgEmail.includes('.gouv.ci'));

    return {
      id: evt.id,
      graphId: evt.id,
      title: subject,
      subject,
      bodyPreview: evt.bodyPreview || undefined,
      organizer: {
        name: evt.organizer?.emailAddress?.name || 'Organisateur inconnu',
        email: evt.organizer?.emailAddress?.address || '',
        isExternal: isExternalOrg,
      },
      start: startDate,
      end: endDate,
      startTimeZone: evt.start?.timeZone || 'UTC',
      endTimeZone: evt.end?.timeZone || 'UTC',
      location: evt.location?.displayName || (isTeams ? 'Microsoft Teams (En ligne)' : 'Salle non spécifiée'),
      isOnlineMeeting: isTeams,
      onlineMeetingProvider: evt.onlineMeetingProvider || (isTeams ? 'teamsForBusiness' : undefined),
      joinWebUrl: evt.onlineMeeting?.joinUrl || undefined,
      webLink: evt.webLink || undefined,
      status: isPast ? 'completed' : 'scheduled',
      captureMethod: isTeams ? 'teams_native' : 'meeting_bot',
      isDgfMeeting: isDgf,
      departmentCode: isDgf ? 'DGF' : undefined,
      attendees,
      hasTranscript: false,
      hasOfficialReport: false,
      userRole,
      isExternalOrganizer: isExternalOrg,
      transcriptAccessibility: isTeams ? (isPast ? 'pending_check' : 'pending_check') : 'no_transcript',
      processingStatus: 'idle',
    };
  });
}
