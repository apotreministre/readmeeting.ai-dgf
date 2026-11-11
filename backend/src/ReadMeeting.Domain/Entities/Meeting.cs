namespace ReadMeeting.Domain.Entities;

/// <summary>
/// Représente une réunion captée depuis Microsoft Teams ou Outlook pour la DGF.
/// </summary>
public class Meeting
{
    public Guid Id { get; set; } = Guid.NewGuid();

    /// <summary>
    /// Identifiant unique d'événement Outlook / Teams (iCalUId ou Graph Id).
    /// </summary>
    public string MicrosoftGraphId { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public string? Description { get; set; }

    public DateTimeOffset StartTime { get; set; }

    public DateTimeOffset EndTime { get; set; }

    public string OrganizerEmail { get; set; } = string.Empty;

    public string OrganizerName { get; set; } = string.Empty;

    /// <summary>
    /// Code ou intitulé de la direction DGF concernée (ex: DPE, DFEE, DAF).
    /// </summary>
    public string? DepartmentCode { get; set; }

    /// <summary>
    /// Lien vers la réunion Microsoft Teams.
    /// </summary>
    public string? JoinWebUrl { get; set; }

    /// <summary>
    /// Statut du cycle de vie de la captation.
    /// </summary>
    public MeetingCaptureStatus Status { get; set; } = MeetingCaptureStatus.Scheduled;

    /// <summary>
    /// Méthode de captation effective (Graph Native vs Bot de réunion).
    /// </summary>
    public CaptureMethod CaptureMethod { get; set; } = CaptureMethod.TeamsNative;

    /// <summary>
    /// Données brutes de transcription textuelle horodatée.
    /// </summary>
    public string? RawTranscript { get; set; }

    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;

    public DateTimeOffset? UpdatedAt { get; set; }

    // Relations
    public MeetingReport? OfficialReport { get; set; }
    public ICollection<ActionItem> ActionItems { get; set; } = new List<ActionItem>();
}

public enum MeetingCaptureStatus
{
    Scheduled = 0,
    InProgress = 1,
    Completed = 2,
    Transcribing = 3,
    Transcribed = 4,
    ReportGenerated = 5,
    Archived = 6,
    Failed = 99
}

public enum CaptureMethod
{
    TeamsNative = 1,
    MeetingBot = 2,
    ManualImport = 3
}
