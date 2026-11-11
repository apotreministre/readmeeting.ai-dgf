# Architecture Système & Spécifications Techniques : ReadMeeting.ai

**Ministère des Finances et du Budget — République de Côte d'Ivoire**  
**Direction Générale des Financements (DGF)**  
*Document d'Architecture Technique (DAT) — Version 1.0*

---

## 1. Contexte & Vision Métier

ReadMeeting.ai est conçu pour moderniser, sécuriser et industrialiser la captation et l'exploitation des réunions stratégiques de la Direction Générale des Financements (DGF).

L'application traite des informations sensibles relatives aux financements publics, aux négociations bilatérales/multilatérales et aux projets de développement national. À ce titre, elle applique une doctrine de **souveraineté des données**, d'intégration native à l'environnement Microsoft 365 existant et de **traçabilité intégrale**.

---

## 2. Piles Technologiques & Rôles

| Échelon | Technologie Retenue | Justification & Rôle |
| :--- | :--- | :--- |
| **Frontend** | Next.js / React 19 + TypeScript | Interface web institutionnelle, réactive, conforme aux normes d'accessibilité et sécurisée par MSAL (OAuth 2.0 PKCE). |
| **Backend API** | ASP.NET Core (.NET 8/9 C#) | Moteur d'orchestration robuste, typage fort, intégration directe du SDK Microsoft Graph et protection JWT Microsoft Entra ID. |
| **Identité & Accès** | Microsoft Entra ID (ex Azure AD) | SSO institutionnel DGF, gestion des rôles (RBAC), isolation multilocataire et révocation centralisée. |
| **Connecteur M365** | Microsoft Graph API (v1.0 & Beta) | Récupération des événements de calendrier Outlook, métadonnées Teams et flux de transcription natifs. |
| **Stockage Relationnel** | Azure SQL Database | Persistance ACID des profils utilisateurs, métadonnées de réunions, actions actées et pistes d'audit. |
| **Archivage Documents** | SharePoint Online (Graph Drive) | Stockage sécurisé des comptes rendus officiels Word (`.docx`) dans les bibliothèques documentaires de la DGF. |
| **Indexation & Recherche** | Azure AI Search | Moteur sémantique vectoriel pour interroger les délibérations et décisions passées en langage naturel. |
| **Traitement Asynchrone** | Azure Functions + Service Bus | File d'attente pour la transcription, l'analyse par Copilot / Azure OpenAI et la conversion de documents sans bloquer l'UI. |

---

## 3. Matrice des Rôles & Habilitations (RBAC)

1. **Super Administrateur DGF (IT / DSI)** :
   - Gestion de l'App Registration, consentements Microsoft Graph, surveillance des journaux d'audit.
2. **Direction Générale / Directeurs de Département** :
   - Accès complet aux réunions de leur direction, validation formelle des comptes rendus, attribution des plans d'actions.
3. **Chargés de Missions / Hôtes de Réunion** :
   - Planification de la capture, révision du compte rendu généré, génération du document Word officiel.
4. **Participants & Collaborateurs** :
   - Consultation des comptes rendus validés et suivi des tâches qui leur sont assignées.

---

## 4. Stratégie de Capture Hybride des Réunions

1. **Mode Primaire (Natif Teams)** :
   - L'organisateur active la transcription Teams lors de la réunion.
   - À la clôture, ReadMeeting.ai interroge `/me/onlineMeetings/{id}/transcripts` via Microsoft Graph.
   - Avantage : Aucun bot intrusif, qualité audio maximale traitée par l'infrastructure Microsoft, conformité totale RGPD/protection des données.
2. **Mode Secondaire (Bot de Réunion)** :
   - Utilisé si la réunion n'a pas été transcrite ou si l'organisateur est externe.
   - Service Azure Calling / Bot Framework rejoignant l'appel Teams en tant qu'auditeur silencieux sur convocation.
