using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ReadMeeting.Application.Interfaces;
using ReadMeeting.Domain.Entities;

namespace ReadMeeting.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize] // Protégé par token JWT émis par Microsoft Entra ID
public class MeetingsController : ControllerBase
{
    private readonly IMicrosoftGraphService _graphService;
    private readonly ILogger<MeetingsController> _logger;

    public MeetingsController(IMicrosoftGraphService graphService, ILogger<MeetingsController> logger)
    {
        _graphService = graphService;
        _logger = logger;
    }

    /// <summary>
    /// Récupère la liste des réunions synchronisées depuis Outlook et filtrées pour la DGF.
    /// </summary>
    [HttpGet]
    [ProducesResponseType(typeof(IEnumerable<Meeting>), StatusCodes.Status200OK)]
    public async Task<IActionResult> GetMeetings([FromQuery] DateTimeOffset? start, [FromQuery] DateTimeOffset? end, CancellationToken ct)
    {
        var userUpn = User.Identity?.Name ?? "anonymous@finances.gouv.ci";
        var startDate = start ?? DateTimeOffset.UtcNow.AddDays(-7);
        var endDate = end ?? DateTimeOffset.UtcNow.AddDays(14);

        _logger.LogInformation("Récupération des réunions pour {UserUpn} du {Start} au {End}", userUpn, startDate, endDate);
        var meetings = await _graphService.GetCalendarEventsAsync(userUpn, startDate, endDate, ct);
        return Ok(meetings);
    }

    /// <summary>
    /// Déclenche la récupération de la transcription Teams native pour une réunion donnée.
    /// </summary>
    [HttpPost("{id}/transcript/sync")]
    public async Task<IActionResult> SyncTranscript(string id, CancellationToken ct)
    {
        _logger.LogInformation("Synchronisation de transcription pour la réunion Graph ID: {Id}", id);
        var transcript = await _graphService.GetOnlineMeetingTranscriptAsync(id, ct);
        
        if (string.IsNullOrEmpty(transcript))
        {
            return NotFound(new { message = "Aucune transcription trouvée ou permission non accordée par l'administrateur du tenant." });
        }

        return Ok(new { meetingId = id, transcriptLength = transcript.Length, status = "Transcribed" });
    }
}
