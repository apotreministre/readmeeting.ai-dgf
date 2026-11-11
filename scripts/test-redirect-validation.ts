import { validateMsalRedirectUri, saveMsalCredentials } from '../src/config/msalConfig';

interface TestCase {
  id: number;
  name: string;
  uri: string;
  activeOrigin?: string;
  expectedValid: boolean;
  expectedErrorKeyword?: string;
}

const ACTIVE_DEV_ORIGIN = 'https://ais-dev-2p5ysc4umjlywyqjtxzst3-540172577201.europe-west2.run.app';
const OLD_CLOUD_RUN_URI = 'https://ais-dev-5ypdjic7tckw5nwjsfbkdh-540172577201.europe-west2.run.app/redirect.html';

const uriValidationCases: TestCase[] = [
  {
    id: 1,
    name: '1. URI officielle de production DGF',
    uri: 'https://readmeeting-ai-dgf.ai.studio/redirect.html',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: true,
  },
  {
    id: 2,
    name: '2. URI avec sous-domaine frauduleux (contournement startsWith)',
    uri: 'https://readmeeting-ai-dgf.ai.studio.evil-example.com/redirect.html',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: false,
    expectedErrorKeyword: 'Origine non autorisée',
  },
  {
    id: 3,
    name: '3. URI avec identifiants utilisateur intégrés (username:password@)',
    uri: 'https://admin:SuperSecret123@readmeeting-ai-dgf.ai.studio/redirect.html',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: false,
    expectedErrorKeyword: 'identifiants utilisateur',
  },
  {
    id: 4,
    name: '4. URI sans identifiants utilisateur (critères conformes)',
    uri: 'https://readmeeting-ai-dgf.ai.studio/redirect.html',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: true,
  },
  {
    id: 5,
    name: '5. URI avec chemin incorrect (/autre.html)',
    uri: 'https://readmeeting-ai-dgf.ai.studio/autre.html',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: false,
    expectedErrorKeyword: 'Chemin non conforme',
  },
  {
    id: 6,
    name: '6. URI avec paramètres de requête (?test=1)',
    uri: 'https://readmeeting-ai-dgf.ai.studio/redirect.html?test=1',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: false,
    expectedErrorKeyword: 'Paramètres de requête',
  },
  {
    id: 7,
    name: '7. URI avec fragment (#code=123)',
    uri: 'https://readmeeting-ai-dgf.ai.studio/redirect.html#code=123',
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: false,
    expectedErrorKeyword: 'Fragments',
  },
  {
    id: 8,
    name: '8. URI localhost en HTTP',
    uri: 'http://localhost:3000/redirect.html',
    activeOrigin: 'http://localhost:3000',
    expectedValid: true,
  },
  {
    id: 9,
    name: '9. URI de développement active autorisée',
    uri: `${ACTIVE_DEV_ORIGIN}/redirect.html`,
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: true,
  },
  {
    id: 10,
    name: '10. URI Cloud Run obsolète et non autorisée',
    uri: OLD_CLOUD_RUN_URI,
    activeOrigin: ACTIVE_DEV_ORIGIN,
    expectedValid: false,
    expectedErrorKeyword: 'Origine non autorisée',
  },
];

console.log('========================================================================');
console.log('SUITE DE TESTS OFFICIELLE — VALIDATION STRICTE MSAL REDIRECT URI (DGF)');
console.log('========================================================================\n');

let allPassed = true;
const resultsTable: any[] = [];

for (const tc of uriValidationCases) {
  const result = validateMsalRedirectUri(tc.uri, tc.activeOrigin);
  const passed = result.isValid === tc.expectedValid &&
    (!tc.expectedErrorKeyword || (result.error && result.error.includes(tc.expectedErrorKeyword)));

  if (!passed) {
    allPassed = false;
  }

  resultsTable.push({
    '#': tc.id,
    'Cas de test': tc.name,
    'URI testée': tc.uri,
    'Attendu': tc.expectedValid ? 'ACCEPTÉE' : 'REJETÉE',
    'Obtenu': result.isValid ? 'ACCEPTÉE' : 'REJETÉE',
    'Détail / Motif': result.isValid ? result.sanitizedUri : result.error,
    'Statut': passed ? 'SUCCÈS' : 'ÉCHEC',
  });
}

console.table(resultsTable);

// Configuration d'un mock localStorage pour tester l'atomicité de saveMsalCredentials
console.log('\n--- Tests unitaires d\'atomicité : saveMsalCredentials() ---');
const fakeStorage: Record<string, string> = {};
(globalThis as any).window = {
  location: { origin: ACTIVE_DEV_ORIGIN },
};
(globalThis as any).localStorage = {
  getItem: (key: string) => fakeStorage[key] ?? null,
  setItem: (key: string, val: string) => { fakeStorage[key] = val; },
  removeItem: (key: string) => { delete fakeStorage[key]; },
};

// ===================================================================================
// Test 11 : Échec de validation -> AUCUNE modification de Client ID, Tenant ID ou URI
// ===================================================================================
const INITIAL_CLIENT_ID = 'ORIGINAL-CLIENT-ID-11111';
const INITIAL_TENANT_ID = 'ORIGINAL-TENANT-ID-22222';
const INITIAL_REDIRECT_URI = 'https://readmeeting-ai-dgf.ai.studio/redirect.html';

fakeStorage['readmeeting_entra_client_id'] = INITIAL_CLIENT_ID;
fakeStorage['readmeeting_entra_tenant_id'] = INITIAL_TENANT_ID;
fakeStorage['readmeeting_entra_redirect_uri'] = INITIAL_REDIRECT_URI;

// Tentative de modification avec une URI frauduleuse (sous-domaine malveillant)
const failedSave = saveMsalCredentials(
  'NEW-ATTACKER-CLIENT-ID',
  'NEW-ATTACKER-TENANT-ID',
  'https://readmeeting-ai-dgf.ai.studio.evil-example.com/redirect.html'
);

const clientIdUnchanged = fakeStorage['readmeeting_entra_client_id'] === INITIAL_CLIENT_ID;
const tenantIdUnchanged = fakeStorage['readmeeting_entra_tenant_id'] === INITIAL_TENANT_ID;
const uriUnchanged = fakeStorage['readmeeting_entra_redirect_uri'] === INITIAL_REDIRECT_URI;
const test11Passed = failedSave.success === false && clientIdUnchanged && tenantIdUnchanged && uriUnchanged;

if (test11Passed) {
  console.log('[SUCCÈS] Test 11 : Échec de validation -> Atomicité préservée (aucune modification partielle dans localStorage).');
} else {
  console.error('[ÉCHEC] Test 11 : Une valeur a été partiellement modifiée dans localStorage !', {
    storedClient: fakeStorage['readmeeting_entra_client_id'],
    storedTenant: fakeStorage['readmeeting_entra_tenant_id'],
    storedUri: fakeStorage['readmeeting_entra_redirect_uri'],
  });
  allPassed = false;
}

// ===================================================================================
// Test 12 : Réussite de validation -> Les 3 valeurs sont correctement enregistrées
// ===================================================================================
const VALID_NEW_CLIENT = 'VALID-CLIENT-ID-99999';
const VALID_NEW_TENANT = 'VALID-TENANT-ID-88888';
const VALID_NEW_URI = `${ACTIVE_DEV_ORIGIN}/redirect.html`;

const successfulSave = saveMsalCredentials(VALID_NEW_CLIENT, VALID_NEW_TENANT, VALID_NEW_URI);

const test12Passed =
  successfulSave.success === true &&
  fakeStorage['readmeeting_entra_client_id'] === VALID_NEW_CLIENT &&
  fakeStorage['readmeeting_entra_tenant_id'] === VALID_NEW_TENANT &&
  fakeStorage['readmeeting_entra_redirect_uri'] === VALID_NEW_URI;

if (test12Passed) {
  console.log('[SUCCÈS] Test 12 : Réussite de validation -> Les trois valeurs sont correctement enregistrées.');
} else {
  console.error('[ÉCHEC] Test 12 : Les valeurs n\'ont pas été enregistrées correctement !');
  allPassed = false;
}

console.log('\n========================================================================');
if (allPassed) {
  console.log('RÉSULTAT GLOBAL : 12/12 TESTS CONFORMES (100% SUCCÈS)');
} else {
  console.error('RÉSULTAT GLOBAL : ÉCHECS DÉTECTÉS');
  process.exit(1);
}
console.log('========================================================================\n');
