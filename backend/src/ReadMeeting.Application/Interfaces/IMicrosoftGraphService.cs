using ReadMeeting.Domain.Entities;

namespace ReadMeeting.Application.Interfaces;

/// <summary>
/// Contrat du service d'intégration avec l'écosystème Microsoft 365 via Microsoft Graph.
/// </summary>
public interface IMicrosoftGraphService
{
    /// <summary>
    /// Récupère la liste des réunions Outlook dans une plage temporelle pour un utilisateur ou une boîte partagée DGF.
    /// </summary>
    Task<IEnumerable<Meeting>> GetCalendarEventsAsync(string userUpn, DateTimeOffset start, DateTimeOffset end, CancellationToken cancellationToken = default);

    /// <summary>
    /// Récupère la transcription native d'une réunion Teams en ligne.
    /// Nécessite la permission OnlineMeetingTranscript.Read.All avec consentement admin.
    /// </summary>
    Task<string?> GetOnlineMeetingTranscriptAsync(string meetingGraphId, CancellationToken cancellationToken = default);

    /// <summary>
    /// Dépose le compte rendu officiel généré au format Word (.docx) sur la bibliothèque documentaire SharePoint Online.
    /// </summary>
    Task<string> UploadReportToSharePointAsync(string siteId, string driveId, string fileName, byte[] docxContent, CancellationToken cancellationToken = default);
}
