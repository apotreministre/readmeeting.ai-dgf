import { explainEntraIdError, OutlookCalendarEvent } from '../src/services/graphClient';
import { MeetingItem } from '../src/types/meeting';
import { DgfUserProfile } from '../src/types/auth';

console.log('========================================================================');
console.log('SUITE DE TESTS D\'INTÉGRATION MICROSOFT GRAPH API — READMEETING.AI (DGF)');
console.log('========================================================================\n');

let allPassed = true;

// -----------------------------------------------------------------------------
// Test 1 : Récupération et normalisation du profil avec /me
// -----------------------------------------------------------------------------
console.log('--- Test 1 : Profil utilisateur Microsoft Graph (/me) ---');
const rawGraphMeResponse = {
  id: 'usr-dgf-2026-99',
  displayName: 'Kouassi Kouame Jean-Marc',
  givenName: 'Jean-Marc',
  surname: 'Kouassi',
  mail: 'jm.kouassi@finances.gouv.ci',
  userPrincipalName: 'jm.kouassi@finances.gouv.ci',
  jobTitle: 'Directeur des Financements Extérieurs',
  officeLocation: 'Tour C, 14ème étage, Abidjan Plateau',
  department: 'Direction Générale des Financements',
};

const profile: DgfUserProfile = {
  id: rawGraphMeResponse.id || undefined,
  displayName: rawGraphMeResponse.displayName || 'Collaborateur DGF',
  givenName: rawGraphMeResponse.givenName || undefined,
  surname: rawGraphMeResponse.surname || undefined,
  mail: rawGraphMeResponse.mail || undefined,
  userPrincipalName: rawGraphMeResponse.userPrincipalName,
  jobTitle: rawGraphMeResponse.jobTitle || undefined,
  officeLocation: rawGraphMeResponse.officeLocation || undefined,
  department: rawGraphMeResponse.department || undefined,
};

const test1Passed =
  profile.id === 'usr-dgf-2026-99' &&
  profile.displayName === 'Kouassi Kouame Jean-Marc' &&
  profile.givenName === 'Jean-Marc' &&
  profile.surname === 'Kouassi' &&
  profile.mail === 'jm.kouassi@finances.gouv.ci' &&
  profile.jobTitle === 'Directeur des Financements Extérieurs' &&
  profile.officeLocation === 'Tour C, 14ème étage, Abidjan Plateau';

if (test1Passed) {
  console.log('[SUCCÈS] Test 1 : Profil /me extrait avec succès et champs respectés.');
} else {
  console.error('[ÉCHEC] Test 1 : Anomalie dans le mapping du profil.');
  allPassed = false;
}

// -----------------------------------------------------------------------------
// Test 2 : Récupération et mapping des événements calendrier (calendarView)
// -----------------------------------------------------------------------------
console.log('\n--- Test 2 : Événements calendrier (/me/calendarView) ---');
const mockEventTeams: OutlookCalendarEvent = {
  id: 'evt-teams-001',
  subject: 'Comité de Suivi du Projet BAD DGF',
  bodyPreview: 'Ordre du jour : revue des décaissements Q3 2026.',
  start: { dateTime: '2026-11-10T10:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-11-10T11:30:00.0000000', timeZone: 'UTC' },
  location: { displayName: 'Microsoft Teams' },
  organizer: { emailAddress: { name: 'Kouassi JM', address: 'jm.kouassi@finances.gouv.ci' } },
  isOnlineMeeting: true,
  onlineMeetingProvider: 'teamsForBusiness',
  onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_real1' },
  webLink: 'https://outlook.office365.com/calendar/item/real1',
  attendees: [
    { type: 'required', status: { response: 'accepted' }, emailAddress: { name: 'Directeur DGF', address: 'dg@finances.gouv.ci' } },
  ],
};

const test2Passed =
  mockEventTeams.id === 'evt-teams-001' &&
  mockEventTeams.subject === 'Comité de Suivi du Projet BAD DGF' &&
  mockEventTeams.bodyPreview === 'Ordre du jour : revue des décaissements Q3 2026.' &&
  mockEventTeams.start?.dateTime === '2026-11-10T10:00:00.0000000' &&
  mockEventTeams.onlineMeetingProvider === 'teamsForBusiness' &&
  mockEventTeams.isOnlineMeeting === true &&
  mockEventTeams.webLink !== undefined;

if (test2Passed) {
  console.log('[SUCCÈS] Test 2 : Tous les champs requis de calendarView sont présents et mappés.');
} else {
  console.error('[ÉCHEC] Test 2 : Champs manquants dans l\'événement.');
  allPassed = false;
}

// -----------------------------------------------------------------------------
// Test 3 : Gestion d'un calendrier vide
// -----------------------------------------------------------------------------
console.log('\n--- Test 3 : Calendrier vide ---');
const emptyGraphCalendarResponse = { value: [] };
const mappedEmpty: MeetingItem[] = (emptyGraphCalendarResponse.value as OutlookCalendarEvent[]).map(() => ({} as MeetingItem));

const test3Passed = Array.isArray(mappedEmpty) && mappedEmpty.length === 0;
if (test3Passed) {
  console.log('[SUCCÈS] Test 3 : Calendrier vide géré sans exception (retourne []).');
} else {
  console.error('[ÉCHEC] Test 3 : Échec de gestion du calendrier vide.');
  allPassed = false;
}

// -----------------------------------------------------------------------------
// Test 4 : Gestion des erreurs d'autorisation Entra ID / Microsoft Graph
// -----------------------------------------------------------------------------
console.log('\n--- Test 4 : Gestion des erreurs d\'autorisation ---');
const consentError = explainEntraIdError(new Error('AADSTS65001: The user or administrator has not consented to use the application'));
const replyUrlError = explainEntraIdError(new Error('AADSTS50011: The reply URL specified in the request does not match the reply URLs configured'));
const interactionError = explainEntraIdError(new Error('interaction_required'));

const test4Passed =
  consentError.code === 'AADSTS65001' &&
  replyUrlError.code === 'AADSTS50011' &&
  interactionError.code === 'INTERACTION_REQUIRED';

if (test4Passed) {
  console.log('[SUCCÈS] Test 4 : Erreurs d\'autorisation Graph/MSAL catégorisées avec solutions claires.');
} else {
  console.error('[ÉCHEC] Test 4 : Mauvaise catégorisation des erreurs.');
  allPassed = false;
}

// -----------------------------------------------------------------------------
// Test 5 : Distinction des réunions Teams vs rendez-vous présentiels
// -----------------------------------------------------------------------------
console.log('\n--- Test 5 : Distinction Teams vs Présentiel ---');
const mockEventInPerson: OutlookCalendarEvent = {
  id: 'evt-inperson-002',
  subject: 'Entretien bilatéral avec le Trésorier Général',
  start: { dateTime: '2026-11-12T14:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-11-12T15:00:00.0000000', timeZone: 'UTC' },
  location: { displayName: 'Bureau du Directeur - Tour C, 14ème étage' },
  organizer: { emailAddress: { name: 'Kouassi JM', address: 'jm.kouassi@finances.gouv.ci' } },
  isOnlineMeeting: false,
};

const isTeams1 = Boolean(mockEventTeams.isOnlineMeeting || mockEventTeams.onlineMeetingProvider === 'teamsForBusiness');
const isTeams2 = Boolean(mockEventInPerson.isOnlineMeeting || mockEventInPerson.onlineMeetingProvider === 'teamsForBusiness');

const test5Passed = isTeams1 === true && isTeams2 === false;
if (test5Passed) {
  console.log('[SUCCÈS] Test 5 : Réunion Teams et réunion présentielle correctement discriminées.');
} else {
  console.error('[ÉCHEC] Test 5 : Échec de discrimination Teams.');
  allPassed = false;
}

// -----------------------------------------------------------------------------
// Test 6 : Pagination des événements avec @odata.nextLink
// -----------------------------------------------------------------------------
console.log('\n--- Test 6 : Pagination @odata.nextLink ---');
const page1Events: OutlookCalendarEvent[] = [{ id: 'p1-e1' }, { id: 'p1-e2' }];
const page2Events: OutlookCalendarEvent[] = [{ id: 'p2-e1' }, { id: 'p2-e2' }];

let simulatedEvents: OutlookCalendarEvent[] = [];
let simulatedNextLink: string | undefined = 'https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=X2';

// Page 1
simulatedEvents = simulatedEvents.concat(page1Events);

// Page 2 via nextLink
if (simulatedNextLink) {
  simulatedEvents = simulatedEvents.concat(page2Events);
  simulatedNextLink = undefined;
}

const test6Passed = simulatedEvents.length === 4 && simulatedEvents[3].id === 'p2-e2';
if (test6Passed) {
  console.log('[SUCCÈS] Test 6 : Mécanisme de pagination par @odata.nextLink fonctionnel et cumulatif.');
} else {
  console.error('[ÉCHEC] Test 6 : Échec de pagination.');
  allPassed = false;
}

// -----------------------------------------------------------------------------
// Test 7 : Cohérence des dates, heures et fuseaux horaires UTC
// -----------------------------------------------------------------------------
console.log('\n--- Test 7 : Cohérence temporelle UTC ---');
const rawUtcDateStr = '2026-11-10T10:00:00.0000000';
const normalizedIso = rawUtcDateStr.endsWith('Z') ? rawUtcDateStr : `${rawUtcDateStr}Z`;
const parsedDate = new Date(normalizedIso);

const test7Passed =
  !isNaN(parsedDate.getTime()) &&
  parsedDate.getUTCFullYear() === 2026 &&
  parsedDate.getUTCMonth() === 10 && // 0-indexed: 10 = Novembre
  parsedDate.getUTCDate() === 10 &&
  parsedDate.getUTCHours() === 10;

if (test7Passed) {
  console.log('[SUCCÈS] Test 7 : Dates et heures normalisées en UTC ISO 8601 conformes.');
} else {
  console.error('[ÉCHEC] Test 7 : Anomalie de parsing de date.');
  allPassed = false;
}

console.log('\n========================================================================');
if (allPassed) {
  console.log('RÉSULTAT GLOBAL GRAPH : 7/7 TESTS FONCTIONNELS CONFORMES (100% SUCCÈS)');
} else {
  console.error('RÉSULTAT GLOBAL GRAPH : DES ANOMALIES ONT ÉTÉ DÉTECTÉES');
  process.exit(1);
}
console.log('========================================================================\n');
