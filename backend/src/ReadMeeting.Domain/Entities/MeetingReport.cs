namespace ReadMeeting.Domain.Entities;

/// <summary>
/// Représente le compte rendu officiel DGF généré à partir de la réunion.
/// </summary>
public class MeetingReport
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public Guid MeetingId { get; set; }
    public Meeting Meeting { get; set; } = null!;

    public string Title { get; set; } = string.Empty;

    /// <summary>
    /// Numéro d'enregistrement officiel (ex: CR-DGF-2026-XXXX).
    /// </summary>
    public string ReferenceNumber { get; set; } = string.Empty;

    public string ExecutiveSummary { get; set; } = string.Empty;

    public string DecisionsAgreed { get; set; } = string.Empty;

    public string KeyDiscussionPoints { get; set; } = string.Empty;

    /// <summary>
    /// URL du document Word officiel généré et stocké dans SharePoint Online.
    /// </summary>
    public string? SharePointDocxUrl { get; set; }

    public string? SharePointFileId { get; set; }

    /// <summary>
    /// Statut de validation institutionnelle du compte rendu.
    /// </summary>
    public ReportApprovalStatus ApprovalStatus { get; set; } = ReportApprovalStatus.Draft;

    public string? ValidatedByEmail { get; set; }

    public DateTimeOffset? ValidatedAt { get; set; }

    public DateTimeOffset GeneratedAt { get; set; } = DateTimeOffset.UtcNow;
}

public enum ReportApprovalStatus
{
    Draft = 0,
    UnderReview = 1,
    Approved = 2,
    Published = 3,
    Rejected = 4
}
