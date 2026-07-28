using Microsoft.AspNetCore.Mvc;
using ServvistaWebAppAPI.Services;

namespace ServvistaWebAppAPI.Controllers
{
    public class JobNotifyRequest
    {
        public string DjId { get; set; } = string.Empty;
    }

    // Deliberately NOT [Authorize] with the JWT scheme - the legacy AddCall
    // controller is a server calling another server, not a technician's
    // browser, so it can't produce a user JWT. Protected by a shared secret
    // instead (set the same value in both apps' config).
    [ApiController]
    [Route("api/internal/notify")]
    public class InternalNotifyController : ControllerBase
    {
        private readonly JobNotificationService _notifier;
        private readonly IConfiguration _configuration;
        private readonly ILogger<InternalNotifyController> _logger;

        public InternalNotifyController(
            JobNotificationService notifier,
            IConfiguration configuration,
            ILogger<InternalNotifyController> logger)
        {
            _notifier = notifier;
            _configuration = configuration;
            _logger = logger;
        }

        [HttpPost("job-assigned")]
        public async Task<IActionResult> JobAssigned(
            [FromBody] JobNotifyRequest request,
            [FromHeader(Name = "X-Internal-Key")] string? key,
            CancellationToken cancellationToken)
        {
            var expectedKey = _configuration["InternalApi:Key"];
            if (string.IsNullOrEmpty(expectedKey) || key != expectedKey)
                return Unauthorized();

            if (string.IsNullOrWhiteSpace(request.DjId))
                return BadRequest("djId is required.");

            var job = await _notifier.LoadJobAsync(request.DjId, cancellationToken);
            if (job is null)
            {
                _logger.LogWarning("Instant notify called for unknown DJ_ID {DjId}.", request.DjId);
                return NotFound();
            }

            var pushDelivered = await _notifier.NotifyJobAsync(job, cancellationToken);

            // Only mark fully notified if push actually succeeded (or there was
            // nothing to push to). Otherwise leave it for the background
            // poller to retry within its 30s window - covers a technician who
            // just granted permission and whose subscription hasn't saved yet.
            if (pushDelivered)
            {
                await _notifier.MarkNotifiedAsync(request.DjId, cancellationToken);
            }

            return Ok(new { message = "Notified", pushDelivered });
        }
    }
}
