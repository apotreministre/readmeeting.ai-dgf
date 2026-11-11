/**
 * ReadMeeting.ai | Direction Générale des Financements (DGF)
 * Ministère des Finances et du Budget - République de Côte d'Ivoire
 *
 * Banc d'essai et tests unitaires / d'intégration :
 * Extension de la récupération des transcriptions aux réunions où l'utilisateur
 * est participant invité (Scénarios A & B - Section 8).
 */

import {
  detectUserMeetingRole,
  isOrganizerExternal,
  explainTranscriptError,
  resolveOnlineMeetingId,
  fetchMeetingTranscripts,
  fetchTranscriptContent,
  checkMeetingTranscriptStatus,
  parseVttTranscript,
  purgeTranscriptCache
} from '../src/services/transcriptService';
import { MeetingItem } from '../src/types/meeting';
import { DgfUserProfile } from '../src/types/auth';

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
console.log('BANC D\'ESSAI : EXTENSION PARTICIPANTS INVITÉS (SCÉNARIOS A & B) — DGF');
console.log('========================================================================\n');

const mockCurrentUser: DgfUserProfile = {
  id: 'user-kouassi-id',
  displayName: 'Kouassi Jean-Marc',
  userPrincipalName: 'jm.kouassi@finances.gouv.ci',
  mail: 'jm.kouassi@finances.gouv.ci',
  jobTitle: 'Chargé d\'Études Senior',
  department: 'DFEE',
  officeLocation: 'Tour C, 14ème étage',
};

const SAMPLE_VTT_CONTENT = `WEBVTT

00:00:01.000 --> 00:00:08.500
<v Directeur Général des Financements>Bienvenue à tous pour cette séance trimestrielle de cadrage budgétaire DGF.</v>

00:00:09.000 --> 00:00:18.200
<v Kouassi Jean-Marc>Merci Monsieur le Directeur Général. Le taux de décaissement global atteint 78.4% ce trimestre.</v>
`;

(async () => {
  purgeTranscriptCache();

  // -----------------------------------------------------------------------------
  // TEST UNITAIRE : DÉTECTION DU RÔLE DE L'UTILISATEUR DANS LA RÉUNION
  // -----------------------------------------------------------------------------
  console.log('--- TEST 0 : Détection rigoureuse des rôles et appartenance tenant ---');

  const meetingOrg: MeetingItem = {
    id: 'm-org',
    title: 'Comité A',
    organizer: { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci' },
    start: '2026-03-01T10:00:00Z',
    end: '2026-03-01T11:00:00Z',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [],
    hasTranscript: true,
    hasOfficialReport: false,
  };
  assert(
    detectUserMeetingRole(meetingOrg, mockCurrentUser) === 'organizer',
    'Détection rôle : "organizer" quand l\'UPN correspond à l\'organisateur'
  );

  const meetingAttendee: MeetingItem = {
    id: 'm-att',
    title: 'Comité B',
    organizer: { name: 'Directeur Général', email: 'dg@finances.gouv.ci' },
    start: '2026-03-01T10:00:00Z',
    end: '2026-03-01T11:00:00Z',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [
      { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required', responseStatus: 'accepted' },
    ],
    hasTranscript: true,
    hasOfficialReport: false,
  };
  assert(
    detectUserMeetingRole(meetingAttendee, mockCurrentUser) === 'attendee',
    'Détection rôle : "attendee" quand l\'UPN est dans la liste des participants convoqués'
  );

  const meetingUndetermined: MeetingItem = {
    id: 'm-undet',
    title: 'Comité C',
    organizer: { name: 'Directeur Général', email: 'dg@finances.gouv.ci' },
    start: '2026-03-01T10:00:00Z',
    end: '2026-03-01T11:00:00Z',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [
      { name: 'Autre Personne', email: 'autre@finances.gouv.ci', type: 'required' },
    ],
    hasTranscript: true,
    hasOfficialReport: false,
  };
  assert(
    detectUserMeetingRole(meetingUndetermined, mockCurrentUser) === 'undetermined',
    'Détection rôle : "undetermined" quand la personne n\'est ni organisateur ni dans les participants convoqués'
  );

  assert(
    isOrganizerExternal('dg@finances.gouv.ci', mockCurrentUser.userPrincipalName) === false,
    'Détection domaine : finances.gouv.ci identifié comme interne au Ministère / DGF'
  );
  assert(
    isOrganizerExternal('s.toure@afdb.org', mockCurrentUser.userPrincipalName) === true,
    'Détection domaine : afdb.org identifié comme locataire externe (BAD)'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 1 : ORGANISATEUR AVEC TRANSCRIPTION ACCESSIBLE (SCÉNARIO A)
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 1 : Organisateur avec transcription accessible (Scénario A) ---');

  const meetingScenario1: MeetingItem = {
    id: 'meeting-sc-1',
    title: 'Comité Clôture Prêt BAD',
    organizer: { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci' },
    start: '2026-02-15T09:00:00Z',
    end: '2026-02-15T11:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sc1%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [{ name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required' }],
    hasTranscript: true,
    hasOfficialReport: false,
  };

  const mockClientScenario1: any = {
    api: (endpoint: string) => ({
      filter: () => ({
        select: () => ({
          get: async () => {
            if (endpoint === '/me/onlineMeetings') {
              return { value: [{ id: 'teams-meeting-sc1-id', subject: 'Comité Clôture' }] };
            }
            return { value: [] };
          }
        })
      }),
      select: () => ({
        get: async () => {
          if (endpoint === '/me/onlineMeetings/teams-meeting-sc1-id/transcripts') {
            return {
              value: [{
                id: 'transcript-sc1-001',
                meetingId: 'teams-meeting-sc1-id',
                createdDateTime: '2026-02-15T11:05:00Z',
                transcriptContentUrl: '/me/onlineMeetings/teams-meeting-sc1-id/transcripts/transcript-sc1-001/content'
              }]
            };
          }
          return { value: [] };
        }
      }),
      query: () => ({
        responseType: () => ({
          get: async () => SAMPLE_VTT_CONTENT
        })
      })
    })
  };

  const resScenario1 = await checkMeetingTranscriptStatus(
    mockClientScenario1,
    meetingScenario1,
    true,
    mockCurrentUser,
    true
  );

  assert(
    resScenario1.userRole === 'organizer',
    'Scénario 1 : Rôle correctement qualifié en "organizer"'
  );
  assert(
    resScenario1.status === 'transcript_available' && resScenario1.accessibility === 'accessible',
    'Scénario 1 : Statut "transcript_available" et accessibilité "accessible"'
  );
  assert(
    resScenario1.auditTrail !== undefined && resScenario1.auditTrail.requestedByUpn === 'jm.kouassi@finances.gouv.ci',
    'Scénario 1 : Piste d\'audit générée avec le demandeur organisateur'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 2 : INVITÉ AVEC TRANSCRIPTION ACCESSIBLE À TOUS (SCÉNARIO B.1)
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 2 : Invité avec transcription accessible à tous (Scénario B.1) ---');

  const meetingScenario2: MeetingItem = {
    id: 'meeting-sc-2',
    title: 'Revue Trimestrielle Portefeuille DGF',
    organizer: { name: 'Directeur Général', email: 'dg@finances.gouv.ci' },
    start: '2026-02-18T14:00:00Z',
    end: '2026-02-18T16:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sc2%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [
      { name: 'Directeur Général', email: 'dg@finances.gouv.ci', type: 'required' },
      { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required' }
    ],
    hasTranscript: true,
    hasOfficialReport: false,
  };

  const mockClientScenario2: any = {
    api: (endpoint: string) => ({
      filter: () => ({
        select: () => ({
          get: async () => {
            // Dans /me, pas trouvé directement, mais résolu sous /users/dg@finances.gouv.ci
            if (endpoint.includes('/users/dg@finances.gouv.ci/onlineMeetings')) {
              return { value: [{ id: 'teams-meeting-sc2-id', subject: 'Revue Trimestrielle' }] };
            }
            return { value: [] };
          }
        })
      }),
      select: () => ({
        get: async () => {
          if (endpoint.includes('transcripts')) {
            return {
              value: [{
                id: 'transcript-sc2-001',
                meetingId: 'teams-meeting-sc2-id',
                createdDateTime: '2026-02-18T16:05:00Z',
                transcriptContentUrl: '/transcripts/transcript-sc2-001/content'
              }]
            };
          }
          return { value: [] };
        }
      }),
      query: () => ({
        responseType: () => ({
          get: async () => SAMPLE_VTT_CONTENT
        })
      })
    })
  };

  const resScenario2 = await checkMeetingTranscriptStatus(
    mockClientScenario2,
    meetingScenario2,
    true,
    mockCurrentUser,
    true
  );

  assert(
    resScenario2.userRole === 'attendee',
    'Scénario 2 : Rôle correctement qualifié en "attendee"'
  );
  assert(
    resScenario2.status === 'transcript_available' && resScenario2.accessibility === 'accessible',
    'Scénario 2 : Transcription de réunion tiers récupérée avec succès'
  );
  assert(
    resScenario2.auditTrail?.organizerUpn === 'dg@finances.gouv.ci' &&
    resScenario2.auditTrail?.requestedByUpn === 'jm.kouassi@finances.gouv.ci',
    'Scénario 2 : Audit consigne explicitement l\'organisateur tiers et le demandeur participant'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 3 : INVITÉ EXCLU PAR LES PARAMÈTRES DE PARTAGE TEAMS (SCÉNARIO B.2)
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 3 : Invité exclu par les paramètres de partage Teams (Scénario B.2) ---');

  const meetingScenario3: MeetingItem = {
    id: 'meeting-sc-3',
    title: 'Comité Arbitrage Restreint',
    organizer: { name: 'Touré Ibrahim', email: 'i.toure@finances.gouv.ci' },
    start: '2026-02-20T10:00:00Z',
    end: '2026-02-20T11:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sc3%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [
      { name: 'Touré Ibrahim', email: 'i.toure@finances.gouv.ci', type: 'required' },
      { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required' }
    ],
    hasTranscript: true,
    hasOfficialReport: false,
  };

  const mockClientScenario3: any = {
    api: (endpoint: string) => ({
      filter: () => ({
        select: () => ({
          get: async () => ({ value: [{ id: 'teams-meeting-sc3-id' }] })
        })
      }),
      select: () => ({
        get: async () => {
          // Simulation du refus Teams : HTTP 403 Forbidden sur /transcripts alors que la permission est accordée
          const forbiddenErr: any = new Error('AccessDenied: Organizer restricted transcript access to organizers and co-organizers only');
          forbiddenErr.statusCode = 403;
          forbiddenErr.code = 'SharingRestriction';
          throw forbiddenErr;
        }
      })
    })
  };

  const resScenario3 = await checkMeetingTranscriptStatus(
    mockClientScenario3,
    meetingScenario3,
    true,
    mockCurrentUser,
    true
  );

  assert(
    resScenario3.status === 'sharing_restricted_by_organizer',
    'Scénario 3 : Qualification distincte en "sharing_restricted_by_organizer"'
  );
  assert(
    resScenario3.accessibility === 'restricted_by_sharing',
    'Scénario 3 : Accessibilité positionnée sur "restricted_by_sharing"'
  );
  assert(
    resScenario3.error?.message === "Cette transcription n'est pas accessible avec les droits de votre compte Microsoft 365. Veuillez demander à l'organisateur de vous accorder l'accès.",
    'Scénario 3 : Message institutionnel obligatoire DGF rigoureusement respecté'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 4 : INVITÉ DONT LA RÉUNION N'A PAS DE TRANSCRIPTION (200 OK VIDE)
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 4 : Invité dont la réunion n\'a pas de transcription (200 OK vide) ---');

  const meetingScenario4: MeetingItem = {
    id: 'meeting-sc-4',
    title: 'Point Étape SIGD',
    organizer: { name: 'Bamba Mamadou', email: 'm.bamba@finances.gouv.ci' },
    start: '2026-02-22T08:00:00Z',
    end: '2026-02-22T09:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sc4%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [
      { name: 'Bamba Mamadou', email: 'm.bamba@finances.gouv.ci', type: 'required' },
      { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required' }
    ],
    hasTranscript: false,
    hasOfficialReport: false,
  };

  const mockClientScenario4: any = {
    api: (endpoint: string) => ({
      filter: () => ({
        select: () => ({
          get: async () => ({ value: [{ id: 'teams-meeting-sc4-id' }] })
        })
      }),
      select: () => ({
        get: async () => {
          // HTTP 200 OK avec tableau vide []
          return { value: [] };
        }
      })
    })
  };

  const resScenario4 = await checkMeetingTranscriptStatus(
    mockClientScenario4,
    meetingScenario4,
    true,
    mockCurrentUser,
    true
  );

  assert(
    resScenario4.status === 'transcript_unavailable',
    'Scénario 4 : Statut "transcript_unavailable" sans lever d\'erreur 403 ou 404'
  );
  assert(
    resScenario4.accessibility === 'no_transcript',
    'Scénario 4 : Accessibilité "no_transcript"'
  );
  assert(
    resScenario4.error?.code === 'EMPTY_TRANSCRIPTS',
    'Scénario 4 : Code EMPTY_TRANSCRIPTS explicitant l\'absence d\'enregistrement dans Teams'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 5 : INVITÉ RÉUNION ORGANISÉE PAR UN AUTRE COLLÈGUE DGF
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 5 : Invité réunion organisée par un autre collaborateur DGF ---');

  const meetingScenario5: MeetingItem = {
    id: 'meeting-sc-5',
    title: 'Comité de Suivi des Emprunts Souverains',
    organizer: { name: 'Kone Amadou', email: 'a.kone@finances.gouv.ci' },
    start: '2026-02-25T11:00:00Z',
    end: '2026-02-25T12:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sc5%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [
      { name: 'Kone Amadou', email: 'a.kone@finances.gouv.ci', type: 'required' },
      { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required' }
    ],
    hasTranscript: true,
    hasOfficialReport: false,
  };

  const tracker = { usedUsersRoute: false };
  const mockClientScenario5: any = {
    api: (endpoint: string) => {
      if (endpoint.includes('/users/a.kone@finances.gouv.ci/onlineMeetings')) {
        tracker.usedUsersRoute = true;
      }
      return {
        filter: () => ({
          select: () => ({
            get: async () => {
              if (endpoint === '/me/onlineMeetings') return { value: [] };
              if (endpoint.includes('/users/a.kone@finances.gouv.ci')) {
                return { value: [{ id: 'teams-meeting-sc5-id' }] };
              }
              return { value: [] };
            }
          })
        }),
        select: () => ({
          get: async () => {
            if (endpoint.includes('transcripts')) {
              return { value: [{ id: 'tr-sc5', meetingId: 'teams-meeting-sc5-id' }] };
            }
            return { value: [] };
          }
        }),
        query: () => ({
          responseType: () => ({
            get: async () => SAMPLE_VTT_CONTENT
          })
        })
      };
    }
  };

  const resScenario5 = await checkMeetingTranscriptStatus(
    mockClientScenario5,
    meetingScenario5,
    true,
    mockCurrentUser,
    true
  );

  assert(
    tracker.usedUsersRoute === true,
    'Scénario 5 : Résolution via la route inter-comptes /users/{organizerEmail}/onlineMeetings'
  );
  assert(
    resScenario5.isExternalOrganizer === false,
    'Scénario 5 : Organisateur confirmé comme interne au même tenant DGF'
  );
  assert(
    resScenario5.accessibility === 'accessible',
    'Scénario 5 : Réunion entre collègues DGF accessible'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 6 : RÉUNION CONVOQUÉE PAR UN LOCATAIRE EXTERNE (CROSS-TENANT)
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 6 : Réunion convoquée par un locataire externe ---');

  const meetingScenario6: MeetingItem = {
    id: 'meeting-sc-6',
    title: 'Table Ronde Bailleurs BAD',
    organizer: { name: 'Dr. Salimata Touré', email: 's.toure@afdb.org', isExternal: true },
    start: '2026-02-26T15:00:00Z',
    end: '2026-02-26T17:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_afdb_sc6%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: false,
    attendees: [
      { name: 'Dr. Salimata Touré', email: 's.toure@afdb.org', type: 'required', isExternal: true },
      { name: 'Kouassi Jean-Marc', email: 'jm.kouassi@finances.gouv.ci', type: 'required' }
    ],
    hasTranscript: true,
    hasOfficialReport: false,
  };

  const resScenario6 = await checkMeetingTranscriptStatus(
    mockClientScenario1,
    meetingScenario6,
    true,
    mockCurrentUser,
    true
  );

  assert(
    resScenario6.isExternalOrganizer === true,
    'Scénario 6 : Détection formelle d\'un organisateur externe (afdb.org)'
  );
  assert(
    resScenario6.status === 'external_tenant_restricted',
    'Scénario 6 : Statut "external_tenant_restricted" bloquant toute tentative de requête cross-tenant indue'
  );
  assert(
    resScenario6.accessibility === 'external_tenant',
    'Scénario 6 : Accessibilité "external_tenant"'
  );
  assert(
    resScenario6.error?.code === 'EXTERNAL_TENANT_RESTRICTED',
    'Scénario 6 : Message explicatif clair sur les politiques de souveraineté et d\'accès inter-locataires'
  );

  // -----------------------------------------------------------------------------
  // SCÉNARIO 7 : RÉUNION DONT L'ID NE PEUT PAS ÊTRE RÉSOLU (SANS CRASH)
  // -----------------------------------------------------------------------------
  console.log('\n--- SCÉNARIO 7 : Réunion dont l\'ID ne peut pas être résolu (sans crash) ---');

  const meetingScenario7: MeetingItem = {
    id: 'meeting-sc-7',
    title: 'Réunion avec lien obsolète ou inaccessible',
    organizer: { name: 'Inconnu', email: 'inconnu@finances.gouv.ci' },
    start: '2026-02-27T10:00:00Z',
    end: '2026-02-27T11:00:00Z',
    isOnlineMeeting: true,
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_invalid%40thread.v2/0',
    status: 'completed',
    captureMethod: 'teams_native',
    isDgfMeeting: true,
    attendees: [],
    hasTranscript: true,
    hasOfficialReport: false,
  };

  const mockClientScenario7: any = {
    api: () => ({
      filter: () => ({
        select: () => ({
          get: async () => ({ value: [] }) // Aucune correspondance
        })
      })
    })
  };

  const resScenario7 = await checkMeetingTranscriptStatus(
    mockClientScenario7,
    meetingScenario7,
    true,
    mockCurrentUser,
    true
  );

  assert(
    resScenario7.status === 'meeting_unresolved',
    'Scénario 7 : Qualification propre en "meeting_unresolved" sans crash de l\'application'
  );
  assert(
    resScenario7.accessibility === 'unresolved',
    'Scénario 7 : Accessibilité marquée "unresolved"'
  );
  assert(
    resScenario7.error?.code === 'MEETING_UNRESOLVED',
    'Scénario 7 : Explication pédagogique invitant à vérifier le lien et les droits'
  );

  // -----------------------------------------------------------------------------
  // BILAN FINAL
  // -----------------------------------------------------------------------------
  console.log('\n========================================================================');
  console.log(`RÉSULTAT DU BANC D'ESSAI DES 7 SCÉNARIOS OBLIGATOIRES :`);
  console.log(`  Tests réussis : \x1b[32m${passed}\x1b[0m`);
  console.log(`  Tests échoués : \x1b[31m${failed}\x1b[0m`);
  console.log('========================================================================\n');

  if (failed > 0) {
    process.exit(1);
  } else {
    process.exit(0);
  }
})();
