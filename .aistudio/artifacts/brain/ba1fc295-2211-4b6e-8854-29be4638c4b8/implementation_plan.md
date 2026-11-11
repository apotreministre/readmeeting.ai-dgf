# Plan d'Initialisation & Architecture Technique : ReadMeeting.ai (DGF)

> **Projet** : ReadMeeting.ai — Assistant Institutionnel Intelligent de Gestion et d'Exploitation des Réunions  
> **Bénéficiaire** : Direction Générale des Financements (DGF), Côte d'Ivoire  
> **Rôle** : Architecte Logiciel Senior & Tech Lead Microsoft Azure / Microsoft 365

---

## Synthèse & Décisions Validées

Conformément à la phase de cadrage et aux réponses apportées :
- **Structure de dépôt** : **Monorepo** unifié (`frontend/`, `backend/`, `infra/`, `docs/`) garantissant la cohérence des contrats d'interface (DTOs, modèles de données) et facilitant le déploiement CI/CD.
- **Environnement Microsoft** : Le tenant Microsoft Entra ID et l'enregistrement d'application (App Registration) sont déjà existants et prêts.
- **Priorité d'initialisation** : Mise en place du **socle d'authentification Microsoft Entra ID (MSAL)** et de la **configuration Microsoft Graph API** réelle (aucun mock ou fausse simulation).

---

## 1. Analyse des Exigences & Diagnostic d'Environnement

### A. Capacités et Contraintes de l'Environnement de Développement
- **Outil Git** : `git version 2.34.1` est disponible directement dans le système. Le dépôt peut être initialisé (`git init`), configuré avec les conventions de commit institutionnelles et relié à un dépôt distant GitHub via HTTPS/SSH.
- **Moteur d'exécution local** : L'environnement d'exécution immédiat AI Studio est basé sur Node.js 22 / TypeScript. Le runtime .NET (`dotnet`) n'est pas préinstallé dans ce conteneur web Linux.
- **Stratégie de développement** :
  1. Le code backend **ASP.NET Core (.NET 8/9)** sera rigoureusement rédigé selon les standards de la *Clean Architecture* C# (projets `.csproj`, solution `.sln`, contrôleurs, injection de dépendances, DbContext EF Core, services Graph).
  2. Le frontend **TypeScript** utilisera la bibliothèque officielle `@azure/msal-react` / `@azure/msal-browser` et `@microsoft/microsoft-graph-client` pour exécuter les flux OAuth 2.0 PKCE directs avec le tenant Entra ID de la DGF.

### B. Matrice des Prérequis & Blocages Techniques Identifiés

| Composant | Prérequis Technique | Risque / Blocage Identifié | Mesure de Contournement / Solution |
| :--- | :--- | :--- | :--- |
| **Entra ID (App Registration)** | Client ID, Tenant ID, Redirect URIs autorisés | Redirect URI de l'URL de test non enregistrée dans Entra ID | Configuration d'une interface de diagnostic montrant exactement l'URI d'origine requise et validation des tokens |
| **Transcriptions Teams natives** | Permissions `OnlineMeetingTranscript.Read.All` | **Consentement Administrateur requis** ; politique Teams autorisant la transcription | Procédure pas-à-pas pour l'administrateur M365 DGF ; mode dégradé clair en cas d'absence de transcription |
| **Bot de capture hybride** | Azure Bot Service + Calling SDK / Graph Communications API | Nécessite un certificat d'application, un endpoint public HTTPS et des permissions d'appel d'application | Prévu en phase 2 après stabilisation de la récupération native par Graph |
| **Microsoft 365 Copilot APIs** | Licences Copilot M365 actives sur le tenant | Coût et disponibilité des licences Copilot au sein du tenant DGF | Architecture hybride découplée : adaptateur Copilot API avec fallback Azure OpenAI Service pour la synthèse institutionnelle |
| **Génération Word DGF** | Modèle officiel `.dotx` / `.docx` avec entête et charte | Spécifications typographiques et mentions légales de l'administration ivoirienne | Gabarit OpenXML / docx structuré intégrant les métadonnées officielles de la DGF |

---

## 2. Structure Cible du Dépôt (Monorepo)

```
readmeeting-ai/
├── .github/
│   └── workflows/
│       ├── frontend-ci.yml              # CI/CD Next.js / Azure Static Web Apps
│       ├── backend-ci.yml               # CI/CD .NET / Azure App Service
│       └── bicep-deploy.yml             # Déploiement Infrastructure Azure
├── frontend/                            # Client Web Institutionnel (Next.js / TS)
│   ├── src/
│   │   ├── auth/                        # Configuration MSAL, msalConfig, MsalProvider
│   │   ├── services/                    # Clients Microsoft Graph (Calendriers, Teams, Transcripts)
│   │   ├── components/
│   │   │   ├── layout/                  # Navigation DGF, Header, Sidebar institutionnelle
│   │   │   ├── dashboard/               # Tableau de bord des réunions DGF
│   │   │   ├── transcript/              # Visualiseur de transcription & synthèse
│   │   │   └── reports/                 # Générateur & exportateur de comptes rendus Word
│   │   ├── types/                       # Interfaces TypeScript alignées sur Graph et DTOs C#
│   │   └── config/                      # Validation des variables d'environnement Entra ID
│   ├── package.json
│   └── tsconfig.json
├── backend/                             # API ASP.NET Core 8+ (Clean Architecture C#)
│   ├── ReadMeeting.sln
│   ├── src/
│   │   ├── ReadMeeting.Domain/          # Entités métiers : Meeting, Transcript, Report, HostProfile
│   │   ├── ReadMeeting.Application/     # CQRS / MediatR, Services d'analyse, Interfaces Graph
│   │   ├── ReadMeeting.Infrastructure/  # EF Core (Azure SQL), GraphServiceClient, Azure AI Search
│   │   └── ReadMeeting.Api/             # Contrôleurs REST, Auth JWT Bearer (Entra ID), Middlewares
│   └── tests/
├── infra/                               # Infrastructure as Code (Bicep)
│   ├── main.bicep                       # Orchestration globale des ressources Azure
│   ├── modules/
│   │   ├── sql.bicep                    # Azure SQL Database & règles de pare-feu
│   │   ├── appservice.bicep             # Azure App Service Linux pour le backend
│   │   ├── search.bicep                 # Azure AI Search pour la base de connaissances
│   │   └── servicebus.bicep             # Azure Service Bus pour la capture asynchrone
└── docs/                                # Documentation d'architecture & guides DGF
    ├── architecture.md                  # Schéma d'architecture détaillé
    ├── entra-id-setup.md                # Procédure de configuration Entra ID et consentements
    └── dgf-word-template-spec.md        # Spécifications du compte rendu officiel
```

---

## 3. Architecture Technique & Flux de Données

```
 ┌────────────────────────────────────────────────────────────────────────┐
 │                      UTILISATEUR DGF (Navigateur)                      │
 └───────────────────────────────────┬────────────────────────────────────┘
                                     │ 1. Auth PKCE / SSO
                                     ▼
                      ┌─────────────────────────────┐
                      │     MICROSOFT ENTRA ID      │
                      │  (Azure AD - Tenant DGF)    │
                      └──────────────┬──────────────┘
                                     │ Token JWT (Scopes: User.Read, Calendars.Read,
                                     │        OnlineMeetingTranscript.Read.All)
                                     ▼
 ┌────────────────────────────────────────────────────────────────────────┐
 │                    READMEETING.AI - FRONTEND WEB                       │
 │        - MSAL React Authentication Context                             │
 │        - Module de Diagnostic de Connectivité M365                     │
 │        - Tableau de bord des Réunions DGF                              │
 └───────────────────┬─────────────────────────────────┬──────────────────┘
                     │                                 │
     Appels Graph    │ Directs (Client-Side)           │ Appels API Sécurisés (Bearer)
                     ▼                                 ▼
 ┌───────────────────────────────┐     ┌──────────────────────────────────┐
 │     MICROSOFT GRAPH API       │     │     BACKEND ASP.NET CORE API     │
 │  - Calendriers Outlook        │     │  - Authentification JWT Entra ID │
 │  - Événements Teams           │     │  - Logique métier & filtrage DGF │
 │  - Transcriptions natives     │     │  - Pipeline d'analyse IA         │
 └───────────────────────────────┘     └─────────────────┬────────────────┘
                                                         │
                                    ┌────────────────────┴────────────────┐
                                    ▼                                     ▼
                     ┌─────────────────────────────┐       ┌──────────────────────────────┐
                     │    AZURE SQL DATABASE       │       │       SHAREPOINT ONLINE      │
                     │  - Profils hôtes & rôles    │       │  - Archivage des comptes     │
                     │  - Métadonnées des réunions │       │    rendus officiels Word     │
                     │  - Historique d'audit       │       │    au format officiel DGF    │
                     └─────────────────────────────┘       └──────────────────────────────┘
```

---

## 4. Plan de Réalisation Progressif (Étape par Étape)

### Étape 1 : Socle Fondamental & Connectivité Entra ID (Session Actuelle)
1. **Initialisation Git & Structure Monorepo** : Initialisation du dépôt local avec le `.gitignore` standard (Node, .NET, Azure, fichiers sensibles).
2. **Configuration & Variables d'Environnement** : Définition des variables obligatoires (`VITE_AZURE_CLIENT_ID`, `VITE_AZURE_TENANT_ID`, `VITE_AZURE_REDIRECT_URI`, `VITE_GRAPH_SCOPES`).
3. **Module MSAL React Réel** : Intégration de `@azure/msal-browser` et `@azure/msal-react` avec gestion des états d'authentification réels (Non connecté, En cours de connexion, Connecté, Erreur d'interaction, Consentement requis).
4. **Console de Diagnostic Technique DGF** : Écran technique permettant de vérifier en direct :
   - La validité de la configuration Entra ID renseignée.
   - Le test de connexion réelle au compte professionnel Microsoft 365 de l'utilisateur.
   - La vérification des permissions accordées (`User.Read`, `Calendars.Read`, etc.).
   - Aucun mock ni simulation : affichage transparent de l'état réel et guidage en cas d'erreur Microsoft.

### Étape 2 : Synchronisation des Réunions & Détection DGF (Prochaine Étape)
1. Requêtage Graph API de l'agenda Outlook (`/me/calendarView`).
2. Moteur de détection automatique des réunions institutionnelles DGF (mots-clés, participants internes `@finances.gouv.ci`, codes projets).
3. Affichage du tableau de bord de pilotage des réunions avec statuts de capture (Planifiée, En cours, Terminée, Transcrite).

### Étape 3 : Capture de Transcription & Architecture Backend
1. Échafaudage de la solution ASP.NET Core (`ReadMeeting.sln`) et des entités Azure SQL.
2. Service de récupération des transcriptions Teams natives.
3. Spécification de l'agent / bot de réunion pour les réunions non transcrites nativement.

### Étape 4 : Moteur de Génération des Comptes Rendus Word DGF
1. Implémentation du générateur OpenXML selon le modèle officiel de la DGF (République de Côte d'Ivoire, Union - Discipline - Travail).
2. Extraction automatique : Ordre du jour, Décisions actées, Plan d'actions avec responsables et échéances, Synthèse des débats.
3. Enregistrement direct sur la bibliothèque SharePoint institutionnelle.

---

## 5. Règles de Sécurité & Non-Simulation

1. **Zéro simulation** : Tous les appels d'authentification et de requêtage Microsoft Graph utilisent les endpoints officiels Microsoft (`login.microsoftonline.com`, `graph.microsoftcom`).
2. **Gestion rigoureuse des erreurs** : Si les identifiants ne sont pas encore configurés dans `.env` ou si le consentement administrateur n'est pas accordé, l'interface affiche explicitement le message d'erreur Microsoft réel (ex: `AADSTS65001`, `AADSTS50011`) avec les instructions exactes de résolution pour l'équipe réseau/sécurité DGF.
3. **Protection des secrets** : Les secrets d'application (Client Secret) ne transitent jamais côté client web ; seul le flux public PKCE (MSAL SPA) est utilisé sur le frontend.
