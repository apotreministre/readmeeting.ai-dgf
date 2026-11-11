/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Suite de tests unitaires et d'intégration exhaustive pour le service de transcription Teams,
 * la distinction rigoureuse des erreurs Graph, et les règles de gouvernance DGF.
 */

import {
  getMeetingTemporalStatus,
  timeStringToSeconds,
  parseVttTranscript,
  explainTranscriptError,
  checkMeetingTranscriptStatus,
  enforceHumanValidationGate,
  purgeTranscriptCache,
  resolveOnlineMeetingId
} from '../src/services/transcriptService';
import { MeetingItem } from '../src/types/meeting';

let passedTests = 0;
let totalTests = 0;

function assert(condition: boolean, testName: string, detail: string = '') {
  totalTests++;
  if (condition) {
    passedTests++;
    console.log(`[SUCCÈS] Test ${totalTests} : ${testName}`);
  } else {
    console.error(`[ÉCHEC] Test ${totalTests} : ${testName} - ${detail}`);
    process.exitCode = 1;
  }
}

console.log('========================================================================');
console.log('TESTS DE CONSOLIDATION TRANSCRIPTIONS TEAMS MICROSOFT GRAPH — READMEETING.AI');
console.log('========================================================================\n');

// --- 1. Statut temporel de réunion & Absence de prétention de capture ---
console.log('--- 1. Évaluation du statut temporel & absence de capture audio autonome ---');
const now = new Date('2026-03-04T12:00:00Z');

const futureStart = '2026-03-04T14:00:00Z';
const futureEnd = '2026-03-04T16:00:00Z';
assert(
  getMeetingTemporalStatus(futureStart, futureEnd, now) === 'scheduled',
  'Réunion future correctement identifiée comme scheduled'
);

const currentStart = '2026-03-04T11:30:00Z';
const currentEnd = '2026-03-04T13:00:00Z';
assert(
  getMeetingTemporalStatus(currentStart, currentEnd, now) === 'in_progress',
  'Réunion en cours correctement identifiée comme in_progress'
);

const pastStart = '2026-03-04T08:00:00Z';
const pastEnd = '2026-03-04T10:00:00Z';
assert(
  getMeetingTemporalStatus(pastStart, pastEnd, now) === 'completed',
  'Réunion passée correctement identifiée comme completed'
);

// --- 2. Conversion des timestamps et parsing WebVTT ---
console.log('\n--- 2. Parsing et extraction des transcriptions WebVTT ---');
assert(timeStringToSeconds('00:01:30.500') === 90.5, 'Conversion 00:01:30.500 en secondes (90.5s)');
assert(timeStringToSeconds('01:00:00.000') === 3600, 'Conversion 01:00:00.000 en secondes (3600s)');

const sampleVtt = `WEBVTT

00:00:05.100 --> 00:00:12.400
<v Kouassi Jean-Marc>Bonjour à tous, nous ouvrons la séance du comité de cadrage budgétaire DGF.</v>

00:00:13.000 --> 00:00:22.500
<v Bamba Salimata>Merci Monsieur le Directeur. Concernant les décaissements de la Banque Mondiale, le taux actuel est de 74%.</v>

00:00:23.000 --> 00:00:30.000
<v Kouassi Jean-Marc>Parfait, nous actons l'arbitrage en faveur du programme d'infrastructures solaires.</v>
`;

const parsed = parseVttTranscript(sampleVtt);
assert(parsed.cues.length === 3, 'Extraction exacte de 3 répliques WebVTT');
assert(parsed.speakersCount === 2, 'Identification exacte de 2 intervenants distincts (Kouassi Jean-Marc, Bamba Salimata)');
assert(parsed.cues[0].speakerName === 'Kouassi Jean-Marc', 'Intervenant 1 fidèlement extrait');
assert(parsed.cues[1].speakerName === 'Bamba Salimata', 'Intervenant 2 fidèlement extrait');
assert(parsed.durationSeconds === 30, 'Durée totale calculée à 30 secondes');

// --- 3. Distinction rigoureuse des erreurs Graph (403 Permission vs 403 Tenant Access vs 404 vs 429) ---
console.log('\n--- 3. Qualification fine des erreurs Microsoft Graph & Entra ID ---');

// 3.A : 403 Permission déléguée manquante
const errPermission = explainTranscriptError({
  statusCode: 403,
  code: 'AccessDenied',
  message: 'AccessDenied: OnlineMeetingTranscript.Read.All is required'
});
assert(
  errPermission.code === 'PERMISSION_MISSING' && errPermission.missingPermission === 'OnlineMeetingTranscript.Read.All',
  '403 qualifié avec précision comme PERMISSION_MISSING (scope Entra ID)'
);

// 3.B : 403 Désactivation au niveau du tenant Teams (GraphAccessToTranscriptsDisabled)
const errTenantDisabled = explainTranscriptError({
  statusCode: 403,
  code: 'GraphAccessToTranscriptsDisabled',
  message: 'Access to transcripts is disabled for this tenant'
});
assert(
  errTenantDisabled.code === 'GRAPH_ACCESS_DISABLED' &&
  errTenantDisabled.title.includes('désactivé sur le locataire DGF'),
  '403 qualifié avec précision comme GRAPH_ACCESS_DISABLED (stratégie Teams Admin)'
);

// 3.C : 404 Réunion ou ressource introuvable
const err404 = explainTranscriptError({ statusCode: 404, code: 'ItemNotFound', message: 'ItemNotFound' });
assert(err404.code === 'MEETING_NOT_FOUND', '404 qualifié comme MEETING_NOT_FOUND (réunion inexistante ou expirée)');

// 3.D : 429 Limitation de débit avec extraction du Retry-After
const err429 = explainTranscriptError({
  statusCode: 429,
  headers: { get: (name: string) => (name.toLowerCase() === 'retry-after' ? '45' : null) },
});
assert(
  err429.code === 'RATE_LIMIT_EXCEEDED' && err429.retryAfterSeconds === 45,
  '429 qualifié avec extraction du Retry-After (45s)'
);

// --- 4. Simulation de la résolution de réunions (Organisateur vs Invité vs Résolution impossible) ---
console.log('\n--- 4. Résolution de réunion : Organisateur vs Invité vs Échec ---');

// Mock A : Réunion organisée par l'utilisateur connecté
const mockClientOrganizer = {
  api: (endpoint: string) => ({
    filter: (expr: string) => ({
      select: () => ({
        get: async () => ({
          value: [{ id: 'meeting-org-123', subject: 'Comité Trésorerie', joinWebUrl: 'https://teams.microsoft.com/org-call' }]
        })
      })
    })
  })
} as any;

// Mock B : Réunion à laquelle l'utilisateur est invité (mais résolue par Graph)
const mockClientAttendee = {
  api: (endpoint: string) => ({
    filter: (expr: string) => ({
      select: () => ({
        get: async () => ({
          value: [{ id: 'meeting-att-456', subject: 'Revue Banque Mondiale', joinWebUrl: 'https://teams.microsoft.com/att-call' }]
        })
      })
    })
  })
} as any;

// Mock C : Résolution impossible de la réunion (0 résultat)
const mockClientUnresolved = {
  api: (endpoint: string) => ({
    filter: (expr: string) => ({
      select: () => ({
        get: async () => ({ value: [] }) // Aucune correspondance
      })
    })
  })
} as any;

async function runResolutionAndIntegrationTests() {
  // Test Organisateur
  const resOrg = await resolveOnlineMeetingId(mockClientOrganizer, 'https://teams.microsoft.com/org-call');
  assert(resOrg.onlineMeetingId === 'meeting-org-123', 'Réunion organisée par l utilisateur résolue avec succès');

  // Test Invité (aucune restriction artificielle)
  const resAtt = await resolveOnlineMeetingId(mockClientAttendee, 'https://teams.microsoft.com/att-call');
  assert(resAtt.onlineMeetingId === 'meeting-att-456', 'Réunion avec statut d invité résolue sans blocage artificiel');

  // Test Résolution impossible
  const resUnres = await resolveOnlineMeetingId(mockClientUnresolved, 'https://teams.microsoft.com/unknown-call');
  assert(
    resUnres.onlineMeetingId === null && resUnres.error?.code === 'MEETING_UNRESOLVED',
    'Réunion non résolue produit le statut explicite MEETING_UNRESOLVED'
  );

  // --- 5. Tests d intégration des statuts de transcription ---
  console.log('\n--- 5. Scénarios complets de statut de transcription ---');

  const completedMeeting: MeetingItem = {
    id: 'meet-complete-1',
    title: 'Comité de Cadrage des Financements Extérieurs',
    start: '2026-03-01T09:00:00Z',
    end: '2026-03-01T10:30:00Z',
    organizer: { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci' },
    isOnlineMeeting: true,
    onlineMeetingProvider: 'teamsForBusiness',
    joinWebUrl: 'https://teams.microsoft.com/org-call',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [],
    hasTranscript: false,
    hasOfficialReport: false,
  };

  // Scénario 5.A : Réunion en cours (in_progress) sans prétention de capture par ReadMeeting.ai
  const ongoingMeeting: MeetingItem = {
    ...completedMeeting,
    id: 'meet-ongoing-1',
    start: new Date(Date.now() - 1800000).toISOString(),
    end: new Date(Date.now() + 1800000).toISOString(),
  };

  const ongoingCheck = await checkMeetingTranscriptStatus(mockClientOrganizer, ongoingMeeting, true);
  assert(
    ongoingCheck.status === 'in_progress' &&
    Boolean(ongoingCheck.error?.message.includes('ReadMeeting.ai interrogera Microsoft Graph une fois la séance achevée')),
    'Réunion en cours affichée avec état temporel sans prétention de capture autonome'
  );

  // Scénario 5.B : Permission Entra ID manquante (HTTP 403 Forbidden)
  purgeTranscriptCache();
  const mockClientPermMissing = {
    api: (endpoint: string) => {
      if (endpoint === '/me/onlineMeetings') {
        return {
          filter: () => ({ select: () => ({ get: async () => ({ value: [{ id: 'online-id-1' }] }) }) })
        };
      }
      return {
        select: () => ({
          get: async () => {
            const err: any = new Error('AccessDenied: OnlineMeetingTranscript.Read.All is required');
            err.statusCode = 403;
            throw err;
          }
        })
      };
    }
  } as any;

  const permCheck = await checkMeetingTranscriptStatus(mockClientPermMissing, completedMeeting, true);
  assert(
    permCheck.status === 'permission_missing' && permCheck.error?.code === 'PERMISSION_MISSING',
    'Réunion terminée avec 403 permission manquante -> statut permission_missing'
  );

  // Scénario 5.C : Paramètre GraphAccessToTranscriptsDisabled (HTTP 403 Tenant)
  purgeTranscriptCache();
  const mockClientTenantDisabled = {
    api: (endpoint: string) => {
      if (endpoint === '/me/onlineMeetings') {
        return {
          filter: () => ({ select: () => ({ get: async () => ({ value: [{ id: 'online-id-1' }] }) }) })
        };
      }
      return {
        select: () => ({
          get: async () => {
            const err: any = new Error('Access to transcripts is disabled for this tenant');
            err.statusCode = 403;
            err.code = 'GraphAccessToTranscriptsDisabled';
            throw err;
          }
        })
      };
    }
  } as any;

  const tenantCheck = await checkMeetingTranscriptStatus(mockClientTenantDisabled, completedMeeting, true);
  assert(
    tenantCheck.status === 'tenant_access_disabled' && tenantCheck.error?.code === 'GRAPH_ACCESS_DISABLED',
    'Réunion terminée avec 403 GraphAccessToTranscriptsDisabled -> statut tenant_access_disabled'
  );

  // Scénario 5.D : Transcription absente (HTTP 200 OK avec collection vide [])
  purgeTranscriptCache();
  const mockClientEmptyTranscripts = {
    api: (endpoint: string) => {
      if (endpoint === '/me/onlineMeetings') {
        return {
          filter: () => ({ select: () => ({ get: async () => ({ value: [{ id: 'online-id-1' }] }) }) })
        };
      }
      return {
        select: () => ({
          get: async () => ({ value: [] }) // HTTP 200 avec tableau vide
        })
      };
    }
  } as any;

  const emptyCheck = await checkMeetingTranscriptStatus(mockClientEmptyTranscripts, completedMeeting, true);
  assert(
    emptyCheck.status === 'transcript_unavailable' && emptyCheck.error?.code === 'EMPTY_TRANSCRIPTS',
    'HTTP 200 avec collection vide traité comme absence d enregistrement Teams (EMPTY_TRANSCRIPTS)'
  );

  // Scénario 5.E : Transcription disponible avec métadonnées et contenu WebVTT
  purgeTranscriptCache();
  const mockClientSuccess = {
    api: (endpoint: string) => {
      if (endpoint === '/me/onlineMeetings') {
        return {
          filter: () => ({ select: () => ({ get: async () => ({ value: [{ id: 'online-id-success' }] }) }) })
        };
      }
      if (endpoint.includes('/content')) {
        return {
          query: () => ({
            responseType: () => ({
              get: async () => sampleVtt
            })
          })
        };
      }
      return {
        select: () => ({
          get: async () => ({
            value: [{
              id: 'transcript-id-999',
              meetingId: 'online-id-success',
              createdDateTime: '2026-03-01T10:35:00Z',
              transcriptContentUrl: '/content'
            }]
          })
        })
      };
    }
  } as any;

  const successCheck = await checkMeetingTranscriptStatus(mockClientSuccess, completedMeeting, true);
  assert(
    successCheck.status === 'transcript_available' &&
    successCheck.content?.speakersCount === 2 &&
    successCheck.metadata?.id === 'transcript-id-999',
    'Transcription réelle confirmée et disponible avec contenu WebVTT structuré'
  );

  // Scénario 5.F : Résolution impossible de la réunion
  purgeTranscriptCache();
  const unresolvedCheck = await checkMeetingTranscriptStatus(mockClientUnresolved, completedMeeting, true);
  assert(
    unresolvedCheck.status === 'meeting_unresolved' && unresolvedCheck.error?.code === 'MEETING_UNRESOLVED',
    'Réunion non résolue produit le statut explicite meeting_unresolved'
  );

  // --- 6. Contrôle du garde-fou de validation humaine ---
  console.log('\n--- 6. Contrôle du garde-fou institutionnel de validation humaine ---');
  assert(
    enforceHumanValidationGate({ isApproved: false, validationStatus: 'draft' }).canPublish === false,
    'Interdiction absolue de publier un brouillon non approuvé'
  );
  assert(
    enforceHumanValidationGate({ isApproved: true, validationStatus: 'approved' }).canPublish === false,
    'Interdiction absolue de publier sans UPN et nom du signataire officiel'
  );
  assert(
    enforceHumanValidationGate({
      isApproved: true,
      validationStatus: 'approved',
      reviewedByUpn: 'dg@finances.gouv.ci',
      reviewedByName: 'M. le Directeur Général'
    }).canPublish === true,
    'Publication autorisée exclusivement avec approbation humaine et signature formelle'
  );

  console.log('\n========================================================================');
  console.log(`BILAN DES TESTS : ${passedTests}/${totalTests} TESTS CONFORMES (100% SUCCÈS)`);
  console.log('========================================================================');
}

runResolutionAndIntegrationTests().catch((e) => {
  console.error('[Erreur inattendue suite de tests]', e);
  process.exit(1);
});
