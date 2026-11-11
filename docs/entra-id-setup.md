# Guide de Configuration Microsoft Entra ID (DGF)

Ce document décrit les étapes obligatoires pour l'administrateur du tenant Microsoft 365 de la **Direction Générale des Financements (DGF)** afin d'activer l'authentification et l'accès Microsoft Graph pour ReadMeeting.ai.

---

## 1. Enregistrement de l'Application dans Microsoft Entra ID

1. Rendez-vous sur le portail [Microsoft Entra admin center](https://entra.microsoft.com/) ou [Azure Portal](https://portal.azure.com/).
2. Accédez à **Identité** > **Applications** > **Inscriptions d'applications** > **Nouvelle inscription**.
3. Renseignez les champs suivants :
   - **Nom** : `ReadMeeting.ai - DGF`
   - **Types de comptes pris en charge** : `Comptes dans cet annuaire d'organisation uniquement (Direction Générale des Financements)`
   - **URI de redirection** :
     - Sélectionnez la plateforme : **Application monopage (SPA)** (Important : Ne pas choisir "Web", choisir **SPA** pour activer le flux OAuth 2.0 PKCE).
     - Renseignez l'URL de votre application (ex: `https://ais-dev-2p5ysc4umjlywyqjtxzst3-540172577201.europe-west2.run.app` ou `http://localhost:3000`).
4. Cliquez sur **Inscrire**.

---

## 2. Récupération des Identifiants

Sur la page de présentation (Overview) de votre application enregistrée, notez :
- **ID d'application (client)** (`Application (client) ID`)
- **ID de l'annuaire (locataire)** (`Directory (tenant) ID`)

Ces deux valeurs doivent être reportées dans les variables d'environnement de ReadMeeting.ai :
```bash
VITE_AZURE_CLIENT_ID="<Votre-ID-Application-Client>"
VITE_AZURE_TENANT_ID="<Votre-ID-Annuaire-Tenant>"
```

---

## 3. Configuration des Autorisations API (Microsoft Graph)

Dans le menu latéral de votre application Entra ID :
1. Cliquez sur **Autorisations de l’API** > **Ajouter une autorisation** > **Microsoft Graph**.
2. Sélectionnez **Autorisations déléguées** (Delegated Permissions) :
   - `User.Read` : Connexion et lecture du profil de l'utilisateur DGF.
   - `Calendars.Read` : Lecture des événements du calendrier Outlook.
   - `Calendars.ReadWrite` : Mise à jour des statuts de réunion et synchronisation.
   - `OnlineMeetings.Read` : Accès aux métadonnées des réunions Teams.
   - `OnlineMeetingTranscript.Read.All` : Récupération des transcriptions des réunions Teams (**Nécessite le consentement de l'administrateur**).
   - `Files.ReadWrite.All` : Sauvegarde des comptes rendus Word dans SharePoint Online.
3. Cliquez sur **Accorder un consentement d'administrateur pour [Nom de votre organisation]**.

---

## 4. Codes d'Erreurs Fréquents & Dépannage

- **`AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application`**
  - *Cause* : L'URL active dans le navigateur n'est pas déclarée dans la section SPA de l'enregistrement d'application.
  - *Correction* : Ajouter l'URL exacte affichée dans la console de diagnostic ReadMeeting.ai dans les Redirect URIs de l'application Entra ID.
- **`AADSTS65001: The user or administrator has not consented to use the application`**
  - *Cause* : Les permissions déléguées (notamment `OnlineMeetingTranscript.Read.All`) nécessitent la validation de l'administrateur général du tenant.
  - *Correction* : Cliquer sur le bouton « Accorder un consentement d'administrateur » dans le portail Entra ID.
