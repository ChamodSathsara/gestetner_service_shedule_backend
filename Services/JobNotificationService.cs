using Dapper;
using ServvistaWebAppAPI.Models;
using ServvistaWebAppAPI.Repositories;
using System.Data.SqlClient;
using System.Text.Json;
using WebPush;

namespace ServvistaWebAppAPI.Services
{
    // Push-only now - SignalR was removed since the instant push (triggered
    // from the legacy AddCall endpoint) reliably beats it, and this avoids
    // keeping a live WebSocket connection open per technician for nothing.
    public class JobNotificationService
    {
        private readonly ILogger<JobNotificationService> _logger;
        private readonly IConfiguration _configuration;
        private readonly PushSubscriptionRepository _pushRepo;
        private static readonly WebPushClient _webPushClient = new();

        public JobNotificationService(
            ILogger<JobNotificationService> logger,
            IConfiguration configuration,
            PushSubscriptionRepository pushRepo)
        {
            _logger = logger;
            _configuration = configuration;
            _pushRepo = pushRepo;
        }

        public async Task<BreakdownModel?> LoadJobAsync(string djId, CancellationToken cancellationToken = default)
        {
            var connectionString = _configuration.GetConnectionString("DefaultConnection");
            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            const string query = """
                SELECT DJ_ID, DJ_DATE, SERIAL_NO, MACHINE_REF_NO, NOTE, CONTACT_TYPE,
                       CUS_NAME, CUS_ADD1, CUS_ADD2, CUS_ADD3, CUS_CONTACT, CUS_TEL_NO, CUS_SMS_NO,
                       TEAM_ID, TEAM_NAME, MACHINE_MODEL_ID, MACHINE_MODEL_NAME, CUS_STATUS,
                       TECH_CODE, TECH_MOBILE, IS_TECH_NOTIFIED
                FROM TBL_DAILY_JOBS
                WHERE DJ_ID = @DjId
                """;

            return await connection.QuerySingleOrDefaultAsync<BreakdownModel>(
                new CommandDefinition(query, new { DjId = djId }, cancellationToken: cancellationToken));
        }

        public async Task MarkNotifiedAsync(string djId, CancellationToken cancellationToken = default)
        {
            var connectionString = _configuration.GetConnectionString("DefaultConnection");
            await using var connection = new SqlConnection(connectionString);
            await connection.ExecuteAsync(new CommandDefinition(
                "UPDATE TBL_DAILY_JOBS SET IS_TECH_NOTIFIED = '1' WHERE DJ_ID = @DjId",
                new { DjId = djId }, cancellationToken: cancellationToken));
        }

        // Returns true if push actually got delivered (or there was nothing
        // to deliver to), so the caller can decide whether it's safe to mark
        // IS_TECH_NOTIFIED = 1 or leave it for the background poller to retry.
        public async Task<bool> NotifyJobAsync(BreakdownModel job, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(job.TECH_CODE))
                return true; // nothing to notify, don't keep retrying

            return await SendWebPushAsync(job.TECH_CODE.Trim(), job, cancellationToken);
        }

        private async Task<bool> SendWebPushAsync(string techCode, BreakdownModel job, CancellationToken cancellationToken)
        {
            var subscriptions = await _pushRepo.GetByTechCodeAsync(techCode);
            if (subscriptions.Count == 0)
            {
                _logger.LogInformation("No push subscriptions stored for technician {TechCode} (yet).", techCode);
                return false;
            }

            var vapidPublicKey = _configuration["Vapid:PublicKey"];
            var vapidPrivateKey = _configuration["Vapid:PrivateKey"];
            var vapidSubject = _configuration["Vapid:Subject"];

            if (string.IsNullOrWhiteSpace(vapidPublicKey) || string.IsNullOrWhiteSpace(vapidPrivateKey))
            {
                _logger.LogWarning("Vapid keys are not configured; skipping web push.");
                return false;
            }

            var location = $"{job.CUS_ADD1} {job.CUS_ADD2} {job.CUS_ADD3}".Trim();
            var payload = JsonSerializer.Serialize(new
            {
                title = "New job assigned",
                body = string.IsNullOrWhiteSpace(location) ? $"Job #{job.DJ_ID}" : $"Job #{job.DJ_ID} - {location}",
                jobId = job.DJ_ID,
                url = $"/technician-dashboard?jobId={job.DJ_ID}",
                // Full detail, used by usePushNotifications.ts to build the
                // in-app notification list/popup when the tab is focused.
                DJ_ID = job.DJ_ID,
                DJ_DATE = job.DJ_DATE,
                SERIAL_NO = job.SERIAL_NO,
                MACHINE_REF_NO = job.MACHINE_REF_NO,
                CUS_NAME = job.CUS_NAME,
                CUS_ADD1 = job.CUS_ADD1,
                CUS_ADD2 = job.CUS_ADD2,
                CUS_ADD3 = job.CUS_ADD3,
                CUS_CONTACT = job.CUS_CONTACT,
                CUS_TEL_NO = job.CUS_TEL_NO,
                TEAM_ID = job.TEAM_ID,
                TEAM_NAME = job.TEAM_NAME,
                TECH_CODE = job.TECH_CODE,
                TECH_MOBILE = job.TECH_MOBILE,
                MACHINE_MODEL_ID = job.MACHINE_MODEL_ID,
                MACHINE_MODEL_NAME = job.MACHINE_MODEL_NAME,
                CUS_STATUS = job.CUS_STATUS,
                NOTE = job.NOTE
            });

            var vapidDetails = new VapidDetails(vapidSubject, vapidPublicKey, vapidPrivateKey);
            var options = new Dictionary<string, object>
            {
                ["vapidDetails"] = vapidDetails,
                ["TTL"] = 60,
                ["headers"] = new Dictionary<string, object> { ["Urgency"] = "high" }
            };

            var anyDelivered = false;

            foreach (var sub in subscriptions)
            {
                var pushSubscription = new WebPush.PushSubscription(sub.ENDPOINT, sub.P256DH, sub.AUTH);
                try
                {
                    await _webPushClient.SendNotificationAsync(pushSubscription, payload, options, cancellationToken);
                    _logger.LogInformation("Push notification sent for job {JobId}, technician {TechCode}.", job.DJ_ID, techCode);
                    anyDelivered = true;
                }
                catch (WebPushException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound
                                                || ex.StatusCode == System.Net.HttpStatusCode.Gone)
                {
                    _logger.LogInformation("Removing stale push subscription {Id} for tech {Tech}", sub.ID, techCode);
                    await _pushRepo.RemoveByIdAsync(sub.ID);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Push notification failed for job {JobId}, technician {TechCode}.", job.DJ_ID, techCode);
                }
            }

            return anyDelivered;
        }
    }
}
