using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ServvistaWebAppAPI.Models;
using ServvistaWebAppAPI.Repositories;

namespace ServvistaWebAppAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PushSubscriptionController : ControllerBase
    {
        private readonly PushSubscriptionRepository _repo;

        public PushSubscriptionController(PushSubscriptionRepository repo)
        {
            _repo = repo;
        }

        // Called once the browser has created a PushSubscription (see usePushNotifications.ts)
        [HttpPost("subscribe")]
        [Authorize]
        public async Task<IActionResult> Subscribe([FromBody] PushSubscriptionRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.TechCode) || string.IsNullOrWhiteSpace(request.Endpoint))
                return BadRequest("techCode and endpoint are required.");

            await _repo.SaveAsync(request);
            return Ok(new { message = "Subscribed" });
        }

        // Called when the technician logs out or disables notifications.
        [HttpPost("unsubscribe")]
        [Authorize]
        public async Task<IActionResult> Unsubscribe([FromBody] PushUnsubscribeRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Endpoint))
                return BadRequest("endpoint is required.");

            await _repo.RemoveAsync(request.Endpoint);
            return Ok(new { message = "Unsubscribed" });
        }
    }
}
