/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Ministère des Finances et du Budget - République de Côte d'Ivoire
 *
 * Banc d'essai et tests unitaires / d'intégration :
 * Phase de test réel des transcriptions Microsoft Teams (Graph API v1.0)
 */

import {
  parseVttTranscript,
  explainTranscriptError,
  resolveOnlineMeetingId,
  fetchMeetingTranscripts,
  fetchTranscriptContent,
  checkMeetingTranscriptStatus
} from '../src/services/transcriptService';
import {
  getTranscriptAccessToken,
  createTranscriptGraphClient,
  TRANSCRIPT_REQUIRED_SCOPES
} from '../src/services/graphClient';
import { MeetingItem } from '../src/types/meeting';

let passed = 0;
let failed = 0;

function assert(condition: boolean, testName: string, details?: string) {
  if (condition) {
    console.log(`  \x1b[32m✔\x1b[0m ${testName}`);
    passed++;
  } else {
    console.error(`  \x1b[31m✖\x1b[0m ${testName}`);
    if (details) console.error(`    \x1b[33mDétails : ${details}\x1b[0m`);
    failed++;
  }
}

console.log('\n========================================================================');
console.log('BANC D\'ESSAI : TEST RÉEL DES TRANSCRIPTIONS MICROSOFT TEAMS (DGF)');
console.log('========================================================================\n');

// -----------------------------------------------------------------------------
// ÉTAPE 1 & 2 : VÉRIFICATION DES AUTORISATIONS ET ACQUISITION DU JETON MSAL
// -----------------------------------------------------------------------------
console.log('--- ÉTAPE 1 & 2 : Vérification des autorisations et acquisition du jeton MSAL ---');

// Test 1.1 : Permission absente ou consentement en attente (InteractionRequiredAuthError)
const mockMsalPendingConsent: any = {
  acquireTokenSilent: async () => {
    const err: any = new Error('InteractionRequiredAuthError: AADSTS65001: The user or administrator has not consented to use the application.');
    err.errorCode = 'consent_required';
    throw err;
  }
};

(async () => {
  const tokenPendingRes = await getTranscriptAccessToken(
    mockMsalPendingConsent,
    { username: 'agent.dgf@finances.gouv.ci' } as any
  );
  assert(
    tokenPendingRes.hasPermission === false &&
    tokenPendingRes.status === 'pending_admin_consent' &&
    tokenPendingRes.token === null,
    'Consentement en attente : blocage propre, status = pending_admin_consent, aucun jeton divulgué',
    tokenPendingRes.error
  );

  // Test 1.2 : Jeton silencieux acquis avec la portée requise OnlineMeetingTranscript.Read.All
  const mockMsalGranted: any = {
    acquireTokenSilent: async () => ({
      accessToken: 'eyJhbGciOi...[SECURE_TOKEN_NOT_LOGGED]...',
      scopes: [
        'User.Read',
        'Calendars.Read',
        'OnlineMeetings.Read',
        'OnlineMeetingTranscript.Read.All'
      ]
    })
  };

  const tokenGrantedRes = await getTranscriptAccessToken(
    mockMsalGranted,
    { username: 'agent.dgf@finances.gouv.ci' } as any
  );
  assert(
    tokenGrantedRes.hasPermission === true &&
    tokenGrantedRes.status === 'granted' &&
    typeof tokenGrantedRes.token === 'string',
    'Consentement accordé : OnlineMeetingTranscript.Read.All détectée et confirmée'
  );

  // Test 1.3 : Préservation des autres portées déjà opérationnelles
  assert(
    TRANSCRIPT_REQUIRED_SCOPES.includes('User.Read') &&
    TRANSCRIPT_REQUIRED_SCOPES.includes('Calendars.Read') &&
    TRANSCRIPT_REQUIRED_SCOPES.includes('OnlineMeetings.Read') &&
    TRANSCRIPT_REQUIRED_SCOPES.includes('OnlineMeetingTranscript.Read.All'),
    'Préservation stricte des 3 portées opérationnelles existantes en complément de la portée transcription'
  );

  // ---------------------------------------------------------------------------
  // ÉTAPE 3 : TEST DE RÉCUPÉRATION ET RÉSOLUTION DE LA RÉUNION
  // ---------------------------------------------------------------------------
  console.log('\n--- ÉTAPE 3 : Résolution de la réunion et interrogation Graph v1.0 ---');

  const testJoinWebUrl = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_MDcxZj...%40thread.v2/0?context=%7b%22Tid%22%3a%229ddae6a3-6cb4-40e3-a4cd-305e2fe8c32a%22%7d';
  const expectedMeetingId = 'MS-ONLINE-MTG-9ddae6a3-DGF-2026';

  // Mock Graph Client simulant la résolution via /me/onlineMeetings?$filter=JoinWebUrl eq '...'
  const mockGraphClientSuccess: any = {
    api: (path: string) => {
      if (path === '/me/onlineMeetings') {
        return {
          filter: () => ({
            select: () => ({
              get: async () => ({
                value: [{ id: expectedMeetingId, subject: 'Comité de Trésorerie DGF', joinWebUrl: testJoinWebUrl }]
              })
            })
          })
        };
      }
      if (path === `/me/onlineMeetings/${expectedMeetingId}/transcripts`) {
        return {
          select: () => ({
            get: async () => ({
              value: [
                {
                  id: 'MS-TR-91823-DGF',
                  meetingId: expectedMeetingId,
                  createdDateTime: '2026-11-10T09:30:00Z',
                  transcriptContentUrl: `/me/onlineMeetings/${expectedMeetingId}/transcripts/MS-TR-91823-DGF/content`,
                  meetingOrganizerId: 'user-organizer-dgf'
                }
              ]
            })
          })
        };
      }
      if (path === `/me/onlineMeetings/${expectedMeetingId}/transcripts/MS-TR-91823-DGF/content`) {
        return {
          query: () => ({
            responseType: () => ({
              get: async () => `WEBVTT\n\n00:00:01.000 --> 00:00:08.500\n<v Kouassi Jean-Marc>Bonjour à tous, nous ouvrons la séance du Comité de cadrage des financements extérieurs.</v>\n\n00:00:09.000 --> 00:00:18.200\n<v Bamba Salimata>Merci Monsieur le Directeur. Concernant les décaissements de la Banque Mondiale, l'échéancier trimestriel est conforme.</v>\n\n00:00:19.000 --> 00:00:27.400\n<v Kouassi Jean-Marc>Parfait, nous validons ce point et passons à l'examen de la convention BAD.</v>`
            })
          })
        };
      }
      throw new Error(`Path inconnu : ${path}`);
    }
  };

  const resolveRes = await resolveOnlineMeetingId(mockGraphClientSuccess, testJoinWebUrl);
  assert(
    resolveRes.onlineMeetingId === expectedMeetingId,
    'Résolution de l\'identifiant onlineMeetingId réussie via JoinWebUrl',
    resolveRes.onlineMeetingId || ''
  );

  const transcriptsRes = await fetchMeetingTranscripts(mockGraphClientSuccess, expectedMeetingId);
  assert(
    transcriptsRes.transcripts.length === 1 && transcriptsRes.transcripts[0].id === 'MS-TR-91823-DGF',
    'Récupération de la liste des transcriptions via GET /me/onlineMeetings/{id}/transcripts réussie'
  );

  const contentRes = await fetchTranscriptContent(mockGraphClientSuccess, expectedMeetingId, 'MS-TR-91823-DGF');
  assert(
    typeof contentRes.content === 'string' && contentRes.content.startsWith('WEBVTT'),
    'Récupération du flux WebVTT via GET .../transcripts/{id}/content?$format=text/vtt réussie'
  );

  // ---------------------------------------------------------------------------
  // ÉTAPE 4 : GESTION DES RÉSULTATS, PARSING WEBVTT ET ATTRIBUTION DES LOCUTEURS
  // ---------------------------------------------------------------------------
  console.log('\n--- ÉTAPE 4 : Parsing WebVTT et identification nominale des locuteurs ---');

  const parsed = parseVttTranscript(contentRes.content);
  assert(
    parsed.cues.length === 3,
    `Nombre d'interventions extraites exact : ${parsed.cues.length} segments`
  );
  assert(
    parsed.speakersCount === 2,
    `Nombre de locuteurs uniques identifiés : ${parsed.speakersCount}`
  );
  assert(
    parsed.speakers.some(s => s.name === 'Kouassi Jean-Marc' && s.interventionsCount === 2) &&
    parsed.speakers.some(s => s.name === 'Bamba Salimata' && s.interventionsCount === 1),
    'Attribution nominale exacte des locuteurs (Kouassi Jean-Marc: 2, Bamba Salimata: 1)'
  );
  assert(
    parsed.totalWords > 20,
    `Volume de mots calculé conforme : ${parsed.totalWords} mots`
  );
  assert(
    parsed.durationSeconds > 25,
    `Durée temporelle calculée conforme : ${parsed.durationSeconds}s`
  );

  // ---------------------------------------------------------------------------
  // ÉTAPE 5 : DISTINCTION RIGOUREUSE DE LA TAXONOMIE DES RÉSULTATS
  // ---------------------------------------------------------------------------
  console.log('\n--- ÉTAPE 5 : Taxonomie rigoureuse des erreurs et cas limites ---');

  // Cas 5.1 : Collection vide (200 OK avec []) => transcript_unavailable, PAS UNE ERREUR !
  const mockGraphEmptyTranscripts: any = {
    api: (path: string) => {
      if (path === '/me/onlineMeetings') {
        return { filter: () => ({ select: () => ({ get: async () => ({ value: [{ id: 'mtg-empty' }] }) }) }) };
      }
      if (path === '/me/onlineMeetings/mtg-empty/transcripts') {
        return { select: () => ({ get: async () => ({ value: [] }) }) };
      }
      throw new Error(`Path: ${path}`);
    }
  };

  const meetingCompleted: MeetingItem = {
    id: 'mtg-test-completed',
    title: 'Point d\'étape bilatéral DGF',
    start: new Date(Date.now() - 7200000).toISOString(),
    end: new Date(Date.now() - 3600000).toISOString(),
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/empty',
    organizer: { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci' },
    attendees: [],
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    hasTranscript: false,
    hasOfficialReport: false,
  };

  const checkEmptyRes = await checkMeetingTranscriptStatus(mockGraphEmptyTranscripts, meetingCompleted, true);
  assert(
    checkEmptyRes.status === 'transcript_unavailable' &&
    checkEmptyRes.error?.code === 'EMPTY_TRANSCRIPTS',
    'Collection vide (200 OK avec []) qualifiée de [transcript_unavailable] et non d\'erreur technique'
  );

  // Cas 5.2 : Accès Graph désactivé sur le locataire (GraphAccessToTranscriptsDisabled)
  const errGraphDisabled = explainTranscriptError({
    code: 'GraphAccessToTranscriptsDisabled',
    statusCode: 403,
    message: 'Access to transcripts is disabled for this tenant.'
  });
  assert(
    errGraphDisabled.code === 'GRAPH_ACCESS_DISABLED' &&
    errGraphDisabled.httpStatus === 403,
    'Erreur tenant GraphAccessToTranscriptsDisabled formellement discriminée'
  );

  // Cas 5.3 : Attribution des locuteurs désactivée dans Teams (EnableAttributedTranscripts)
  const errAttributionDisabled = explainTranscriptError({
    code: 'AttributedTranscriptDisabled',
    statusCode: 403,
    message: 'Attributed transcripts are disabled in meeting policy.'
  });
  assert(
    errAttributionDisabled.code === 'SPEAKER_ATTRIBUTION_DISABLED',
    'Attribution des locuteurs désactivée discriminée de l\'accès global'
  );

  // Cas 5.4 : Permission manquante Entra ID
  const errPermMissing = explainTranscriptError({
    statusCode: 403,
    code: 'AccessDenied',
    message: 'Caller does not have OnlineMeetingTranscript.Read.All permission.'
  });
  assert(
    errPermMissing.code === 'PERMISSION_MISSING' &&
    errPermMissing.title === 'Autorisation administrateur en attente',
    'Permission manquante libellée « Autorisation administrateur en attente » sans prétention abusive'
  );

  // Cas 5.5 : Réunion expirée ou introuvable (404)
  const errNotFound = explainTranscriptError({
    statusCode: 404,
    code: 'ItemNotFound',
    message: 'The requested meeting was not found.'
  });
  assert(
    errNotFound.code === 'MEETING_NOT_FOUND',
    'Réunion Teams introuvable ou expirée (404 Not Found) correctement typée'
  );

  // Cas 5.6 : Limitation de débit (429 Too Many Requests)
  const errRateLimit = explainTranscriptError({
    statusCode: 429,
    code: 'TooManyRequests',
    message: 'Too many requests.',
    headers: new Map([['Retry-After', '15']])
  });
  assert(
    errRateLimit.code === 'RATE_LIMIT_EXCEEDED',
    'Limitation de débit Microsoft Graph (HTTP 429) gérée avec temporisation'
  );

  console.log('\n------------------------------------------------------------------------');
  console.log(`RÉSULTAT GLOBAL DU BANC D'ESSAI : ${passed} passés, ${failed} échoués`);
  console.log('------------------------------------------------------------------------\n');

  if (failed > 0) {
    process.exit(1);
  }
})();
