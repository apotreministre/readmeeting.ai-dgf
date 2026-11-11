import React, { useState } from 'react';
import {
  FolderTree,
  Server,
  Cloud,
  Shield,
  FileCode2,
  GitBranch,
  Terminal,
  CheckCircle2,
  ExternalLink,
  Layers,
  Database,
  Video
} from 'lucide-react';

export const ArchitectureView: React.FC = () => {
  const [selectedTab, setSelectedTab] = useState<'monorepo' | 'azure_topology' | 'github_cicd' | 'transcripts_pipeline' | 'security'>('monorepo');

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-16">
      {/* Header */}
      <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-base font-bold text-white">Dossier d'Architecture Technique (DAT) & Dépôt Source</h2>
            <span className="px-2 py-0.5 text-[10px] font-mono bg-emerald-950 text-emerald-300 border border-emerald-800 rounded">
              v1.0 DGF
            </span>
          </div>
          <p className="text-xs text-slate-400 mt-1">
            Documentation détaillée de la structure monorepo, des composants Azure et de l'intégration continue GitHub.
          </p>
        </div>

        {/* Tab Pills */}
        <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800">
          <button
            onClick={() => setSelectedTab('monorepo')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
              selectedTab === 'monorepo' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
            }`}
          >
            Arborescence Monorepo
          </button>
          <button
            onClick={() => setSelectedTab('azure_topology')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
              selectedTab === 'azure_topology' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
            }`}
          >
            Topologie Azure
          </button>
          <button
            onClick={() => setSelectedTab('github_cicd')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
              selectedTab === 'github_cicd' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
            }`}
          >
            GitHub & CI/CD
          </button>
          <button
            onClick={() => setSelectedTab('transcripts_pipeline')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
              selectedTab === 'transcripts_pipeline' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
            }`}
          >
            Pipeline Transcriptions DGF
          </button>
          <button
            onClick={() => setSelectedTab('security')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
              selectedTab === 'security' ? 'bg-emerald-600 text-white' : 'text-slate-400 hover:text-white'
            }`}
          >
            Gouvernance & Sécurité
          </button>
        </div>
      </div>

      {/* Tab 1 : Monorepo */}
      {selectedTab === 'monorepo' && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2 bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-4">
            <h3 className="text-sm font-semibold text-slate-200 flex items-center gap-2">
              <FolderTree className="w-4 h-4 text-emerald-400" />
              Organisation des Dossiers du Projet
            </h3>

            <div className="font-mono text-xs text-slate-300 bg-slate-950 p-4 rounded-xl border border-slate-800 space-y-1 overflow-x-auto leading-relaxed">
              <div className="text-emerald-400 font-bold">readmeeting-ai/ (Monorepo DGF)</div>
              <div className="text-slate-500">├── .github/workflows/</div>
              <div className="text-slate-400">│   ├── frontend-ci.yml           <span className="text-slate-500"># CI/CD client web Next.js / Azure SWA</span></div>
              <div className="text-slate-400">│   ├── backend-ci.yml            <span className="text-slate-500"># Build .NET 8, tests unitaires et déploiement App Service</span></div>
              <div className="text-slate-400">│   └── bicep-deploy.yml          <span className="text-slate-500"># Déploiement IaC Azure Resource Manager</span></div>
              <div className="text-slate-500">├── frontend/                     <span className="text-blue-400 font-bold"># Application Web Institutionnelle</span></div>
              <div className="text-slate-400">│   ├── src/auth/                 <span className="text-slate-500"># Intégration officielle MSAL React</span></div>
              <div className="text-slate-400">│   ├── src/services/             <span className="text-slate-500"># Client Microsoft Graph v1.0</span></div>
              <div className="text-slate-400">│   ├── src/components/           <span className="text-slate-500"># Tableaux de bord, diagnostics, visionneuse CR</span></div>
              <div className="text-slate-400">│   └── package.json</div>
              <div className="text-slate-500">├── backend/                      <span className="text-amber-400 font-bold"># API ASP.NET Core (Clean Architecture)</span></div>
              <div className="text-slate-400">│   ├── ReadMeeting.sln           <span className="text-slate-500"># Solution Visual Studio / dotnet</span></div>
              <div className="text-slate-400">│   ├── src/ReadMeeting.Domain/   <span className="text-slate-500"># Entités Meeting, MeetingReport, ActionItem</span></div>
              <div className="text-slate-400">│   ├── src/ReadMeeting.Application/ <span className="text-slate-500"># Interfaces Graph, CQRS, Règles DGF</span></div>
              <div className="text-slate-400">│   ├── src/ReadMeeting.Infrastructure/ <span className="text-slate-500"># EF Core Azure SQL & GraphServiceClient</span></div>
              <div className="text-slate-400">│   └── src/ReadMeeting.Api/      <span className="text-slate-500"># Contrôleurs REST protégés par Entra ID</span></div>
              <div className="text-slate-500">├── infra/                        <span className="text-purple-400 font-bold"># Infrastructure as Code (Bicep)</span></div>
              <div className="text-slate-400">│   └── main.bicep                <span className="text-slate-500"># Provisioning App Service, SQL, Service Bus, Search</span></div>
              <div className="text-slate-500">└── docs/                         <span className="text-teal-400 font-bold"># Spécifications & Guides Administrateur</span></div>
              <div className="text-slate-400">    ├── architecture.md           <span className="text-slate-500"># Document d'Architecture Technique</span></div>
              <div className="text-slate-400">    └── entra-id-setup.md         <span className="text-slate-500"># Guide pas-à-pas enregistrement App Entra ID</span></div>
            </div>
          </div>

          <div className="space-y-4">
            <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-3">
              <h4 className="text-xs font-semibold text-slate-200 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                Avantages du Monorepo pour la DGF
              </h4>
              <ul className="text-xs text-slate-400 space-y-2 leading-relaxed">
                <li className="flex items-start gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 shrink-0 mt-1.5"></span>
                  <span><strong>Synchronisation des contrats :</strong> Les modèles de réunions et de comptes rendus restent toujours alignés entre le backend C# et le client TypeScript.</span>
                </li>
                <li className="flex items-start gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 shrink-0 mt-1.5"></span>
                  <span><strong>Déploiement atomique :</strong> Un commit ou une pull-request valide conjointement l'infrastructure, l'API et le frontend.</span>
                </li>
                <li className="flex items-start gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 shrink-0 mt-1.5"></span>
                  <span><strong>Audit institutionnel :</strong> Facilité pour l'équipe sécurité de la DGF de suivre l'historique complet dans un seul référentiel.</span>
                </li>
              </ul>
            </div>
          </div>
        </div>
      )}

      {/* Tab 2 : Azure Topology */}
      {selectedTab === 'azure_topology' && (
        <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-5">
          <h3 className="text-sm font-semibold text-white flex items-center gap-2">
            <Cloud className="w-4 h-4 text-blue-400" />
            Topologie des Services Microsoft Azure DGF
          </h3>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-200">Microsoft Entra ID</span>
                <span className="text-[10px] text-blue-400 bg-blue-950 px-2 py-0.5 rounded border border-blue-800">Identité</span>
              </div>
              <p className="text-xs text-slate-400">Authentification SSO, gestion des rôles RBAC, flux OAuth 2.0 PKCE pour le client web et validation JWT sur l'API.</p>
            </div>

            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-200">Azure App Service Linux</span>
                <span className="text-[10px] text-emerald-400 bg-emerald-950 px-2 py-0.5 rounded border border-emerald-800">API .NET</span>
              </div>
              <p className="text-xs text-slate-400">Hébergement conteneurisé du backend ASP.NET Core avec TLS 1.3, AlwaysOn et Managed Identity pour accéder à Azure SQL.</p>
            </div>

            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-200">Azure SQL Database</span>
                <span className="text-[10px] text-purple-400 bg-purple-950 px-2 py-0.5 rounded border border-purple-800">Données</span>
              </div>
              <p className="text-xs text-slate-400">Stockage relationnel ACID des métadonnées de réunions, assignations d'actions, pistes d'audit et droits d'accès.</p>
            </div>

            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-200">SharePoint Online</span>
                <span className="text-[10px] text-teal-400 bg-teal-950 px-2 py-0.5 rounded border border-teal-800">Documents</span>
              </div>
              <p className="text-xs text-slate-400">Bibliothèque documentaire institutionnelle hébergeant les comptes rendus Word (.docx) générés selon le modèle officiel DGF.</p>
            </div>

            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-200">Azure AI Search</span>
                <span className="text-[10px] text-amber-400 bg-amber-950 px-2 py-0.5 rounded border border-amber-800">RAG Sémantique</span>
              </div>
              <p className="text-xs text-slate-400">Indexation vectorielle des transcriptions et comptes rendus pour l'assistant conversationnel « Demander à ReadMeeting.ai ».</p>
            </div>

            <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-200">Azure Service Bus</span>
                <span className="text-[10px] text-rose-400 bg-rose-950 px-2 py-0.5 rounded border border-rose-800">Orchestration</span>
              </div>
              <p className="text-xs text-slate-400">File de messages garantissant le traitement fiable et asynchrone des transcriptions et synthèses volumineuses.</p>
            </div>
          </div>
        </div>
      )}

      {/* Tab 3 : GitHub & CI/CD */}
      {selectedTab === 'github_cicd' && (
        <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-4">
          <h3 className="text-sm font-semibold text-white flex items-center gap-2">
            <GitBranch className="w-4 h-4 text-emerald-400" />
            Connectivité GitHub & Procédure de Synchronisation
          </h3>

          <p className="text-xs text-slate-300 leading-relaxed">
            Le système intègre l'utilitaire <code className="text-emerald-400 font-mono">git 2.34</code>. Pour relier cette base de code à un dépôt GitHub de la DGF :
          </p>

          <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 font-mono text-xs text-slate-200 space-y-2">
            <div className="text-slate-500"># 1. Initialiser le dépôt local et créer la branche principale</div>
            <div>git init -b main</div>
            <div className="text-slate-500 pt-2"># 2. Ajouter tous les fichiers du monorepo</div>
            <div>git add .</div>
            <div>git commit -m "feat: initialisation architecture ReadMeeting.ai DGF (Entra ID, Clean Architecture, Bicep)"</div>
            <div className="text-slate-500 pt-2"># 3. Lier au dépôt distant GitHub (ex: organisation DGF-CI)</div>
            <div>git remote add origin https://github.com/votre-organisation/readmeeting-ai.git</div>
            <div className="text-slate-500 pt-2"># 4. Pousser vers GitHub</div>
            <div>git push -u origin main</div>
          </div>

          <div className="p-4 bg-blue-950/40 border border-blue-800/60 rounded-lg text-xs text-blue-200 space-y-1">
            <span className="font-bold block text-white">Sécurité des identifiants :</span>
            Le fichier <code className="text-blue-300">.gitignore</code> est préconfiguré pour exclure tout fichier contenant des secrets (<code className="text-blue-300">.env</code>, <code className="text-blue-300">*.pfx</code>, <code className="text-blue-300">appsettings.Development.json</code>). Seul le template <code className="text-blue-300">.env.example</code> est versionné.
          </div>
        </div>
      )}

      {/* Tab Transcriptions & Pipeline Documentaire DGF */}
      {selectedTab === 'transcripts_pipeline' && (
        <div className="space-y-6">
          {/* 1. Audit technique et faisabilité */}
          <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                <Video className="w-4 h-4 text-indigo-400" />
                Audit Technique & Faisabilité de la Récupération des Transcriptions Teams
              </h3>
              <span className="text-[10px] text-amber-400 bg-amber-950 px-2 py-0.5 rounded border border-amber-800">
                Phase d'Étude & Cadrage
              </span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-xs text-slate-300">
              <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
                <span className="font-bold text-white block">1. Identification des réunions Teams</span>
                <p className="text-slate-400 leading-relaxed">
                  L'endpoint <code className="text-blue-300 font-mono">/me/calendarView</code> identifie les sessions Teams via les propriétés <code className="text-blue-300 font-mono">isOnlineMeeting: true</code>, <code className="text-blue-300 font-mono">onlineMeetingProvider: "teamsForBusiness"</code> et l'URL <code className="text-blue-300 font-mono">onlineMeeting.joinUrl</code>.
                </p>
              </div>

              <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
                <span className="font-bold text-white block">2. Non-interchangeabilité des identifiants</span>
                <p className="text-slate-400 leading-relaxed">
                  L'identifiant d'un événement Outlook (<code className="text-amber-300 font-mono">AAMkAG...</code>) est un ID de messagerie MAPI, <strong>strictement distinct</strong> de l'identifiant <code className="text-emerald-300 font-mono">onlineMeeting.id</code> (thread Teams encodé base64).
                </p>
              </div>

              <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
                <span className="font-bold text-white block">3. Résolution Organisateur vs Invités</span>
                <p className="text-slate-400 leading-relaxed">
                  ReadMeeting.ai interroge <code className="text-blue-300 font-mono">GET /me/onlineMeetings?$filter=JoinWebUrl eq '...'</code> sans restriction artificielle de rôle. Si l'ID ne peut être résolu (invitation externe ou stratégie de locataire restreinte), un statut explicite <code className="text-amber-300 font-mono">meeting_unresolved</code> est renvoyé.
                </p>
              </div>

              <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
                <span className="font-bold text-rose-300 block">4. Double Verrouillage Sécuritaire M365</span>
                <p className="text-slate-400 leading-relaxed">
                  L'accès aux transcriptions requiert <strong>deux autorisations cumulatives</strong> : la permission Entra ID <code className="text-emerald-400 font-mono">OnlineMeetingTranscript.Read.All</code> <strong>ET</strong> l'activation du paramètre Teams <code className="text-purple-300 font-mono">GraphAccessToTranscriptsDisabled = $false</code>.
                </p>
              </div>

              <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
                <span className="font-bold text-amber-300 block">5. Politiques Teams & Pas d'Automatisme</span>
                <p className="text-slate-400 leading-relaxed">
                  La transcription n'est pas automatique : elle dépend de la stratégie <code className="text-blue-300 font-mono">TeamsMeetingPolicy.AllowTranscription</code> et doit être <strong>effectivement déclenchée</strong> (« Démarrer la transcription ») par un participant lors de la réunion.
                </p>
              </div>

              <div className="p-4 bg-slate-950 rounded-xl border border-slate-800 space-y-2">
                <span className="font-bold text-blue-300 block">6. Démarcation Transcription vs Vidéo</span>
                <p className="text-slate-400 leading-relaxed">
                  ReadMeeting.ai exploite exclusivement la <strong>transcription textuelle WebVTT</strong> via Graph. L'enregistrement vidéo (.mp4) relève de SharePoint/OneDrive et ne fait l'objet d'aucune capture ni stockage par ReadMeeting.ai.
                </p>
              </div>
            </div>
          </div>

          {/* 2. Architecture du pipeline documentaire DGF */}
          <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-sm font-semibold text-white flex items-center gap-2">
                <Layers className="w-4 h-4 text-emerald-400" />
                Pipeline de Traitement Documentaire Institutionnel (9 Étapes)
              </h3>
              <span className="text-[10px] text-emerald-400 bg-emerald-950 px-2 py-0.5 rounded border border-emerald-800">
                Garde-Fou Validation Humaine Obligatoire
              </span>
            </div>

            <div className="space-y-2.5 text-xs">
              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">01</span>
                <div>
                  <strong className="text-white block">Extraction du flux brut Teams</strong>
                  Téléchargement du fichier WebVTT via <code className="text-blue-300 font-mono">/transcripts/{'{id}'}/content?$format=text/vtt</code> avec horodatages et étiquettes vocales.
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">02</span>
                <div>
                  <strong className="text-white block">Nettoyage & Normalisation Linguistique</strong>
                  Suppression des bruits d'élocution, rattachement des répliques aux intervenants convoqués et détection des pseudonymes Teams.
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">03</span>
                <div>
                  <strong className="text-white block">Segmentation Thématique selon l'Ordre du Jour</strong>
                  Alignement des temps de parole avec les points officiels de la convocation Outlook (bodyPreview).
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">04</span>
                <div>
                  <strong className="text-white block">Extraction Formelle des Décisions</strong>
                  Identification des arbitrages, des accords de principe, des montants financiers et des consensus actés.
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">05</span>
                <div>
                  <strong className="text-white block">Matrice d'Actions & Assignation des Responsabilités</strong>
                  Génération des tâches avec échéances précises et désignation des directions/agents responsables (DPE, DFEE, Trésor).
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">06</span>
                <div>
                  <strong className="text-white block">Génération du Projet de Compte Rendu (Brouillon)</strong>
                  Mise en page selon la charte typographique et administrative de la République de Côte d'Ivoire.
                </div>
              </div>

              {/* POINT D'ARRÊT CRITIQUE */}
              <div className="p-4 bg-gradient-to-r from-amber-950/80 via-slate-950 to-amber-950/80 border-2 border-amber-600/80 rounded-xl flex items-start gap-3 shadow-lg">
                <span className="font-mono text-amber-400 font-extrabold text-sm shrink-0">07</span>
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <strong className="text-amber-300 text-sm">POINT D'ARRÊT OBLIGATOIRE : Validation Humaine & Signature Institutionnelle</strong>
                    <span className="px-2 py-0.5 text-[9px] font-bold bg-amber-500 text-slate-950 rounded uppercase">Garde-Fou Inviolable</span>
                  </div>
                  <p className="text-xs text-amber-100 leading-relaxed">
                    Le compte rendu ne sera <strong>JAMAIS</strong> automatiquement considéré comme officiel ni diffusé sans l'approbation formelle, la relecture contradictoire et la signature numérique (avec UPN et horodatage UTC) du rapporteur ou du président de séance.
                  </p>
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">08</span>
                <div>
                  <strong className="text-white block">Validation & Attribution du Numéro de Référence Officiel</strong>
                  Enregistrement sous référence ministérielle unique (<code className="text-emerald-400 font-mono">CR-DGF-2026-XXX</code>).
                </div>
              </div>

              <div className="p-3 bg-slate-950 rounded-lg border border-slate-800 flex items-start gap-3">
                <span className="font-mono text-emerald-400 font-bold shrink-0">09</span>
                <div>
                  <strong className="text-white block">Génération Word (.docx) & Archivage SharePoint Online</strong>
                  Production du document officiel conforme et synchronisation dans la bibliothèque sécurisée de la DGF.
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Tab 4 : Security */}
      {selectedTab === 'security' && (
        <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 space-y-4">
          <h3 className="text-sm font-semibold text-white flex items-center gap-2">
            <Shield className="w-4 h-4 text-emerald-400" />
            Gouvernance des Données & Conformité DGF
          </h3>

          <div className="space-y-3 text-xs text-slate-300 leading-relaxed">
            <div className="p-3 bg-slate-950 rounded-lg border border-slate-800">
              <span className="font-bold text-white block mb-1">1. Souveraineté & Chiffrement au repos</span>
              Toutes les données de réunion et les comptes rendus sont hébergés dans le locataire Azure sous juridiction sécurisée, avec chiffrement transparent des données (TDE) sur Azure SQL et clés gérées par le client (CMK via Azure Key Vault).
            </div>
            <div className="p-3 bg-slate-950 rounded-lg border border-slate-800">
              <span className="font-bold text-white block mb-1">2. Séparation stricte des privilèges</span>
              Les transcriptions sensibles de la DGF ne sont jamais exportées vers des LLM publics non régis par l'accord d'entreprise Microsoft Enterprise Agreement.
            </div>
            <div className="p-3 bg-slate-950 rounded-lg border border-slate-800">
              <span className="font-bold text-white block mb-1">3. Piste d'audit inviolable</span>
              Chaque accès, chaque consultation de transcription et chaque génération de compte rendu est tracé avec l'identifiant UPN Entra ID de l'utilisateur, l'horodatage UTC et l'adresse IP.
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
