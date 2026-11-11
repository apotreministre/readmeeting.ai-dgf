namespace ReadMeeting.Domain.Entities;

/// <summary>
/// Tâche d'action issue des délibérations de la réunion avec porteur et date limite.
/// </summary>
public class ActionItem
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public Guid MeetingId { get; set; }
    public Meeting Meeting { get; set; } = null!;

    public string Description { get; set; } = string.Empty;

    public string AssigneeName { get; set; } = string.Empty;

    public string AssigneeEmail { get; set; } = string.Empty;

    public DateTimeOffset? DueDate { get; set; }

    public ActionPriority Priority { get; set; } = ActionPriority.Medium;

    public ActionStatus Status { get; set; } = ActionStatus.Pending;

    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;
}

public enum ActionPriority
{
    Low = 0,
    Medium = 1,
    High = 2,
    Critical = 3
}

public enum ActionStatus
{
    Pending = 0,
    InProgress = 1,
    Completed = 2,
    Deferred = 3
}
