using Dapper;
using Microsoft.AspNetCore.SignalR;
using ServvistaWebAppAPI.Classes;
using ServvistaWebAppAPI.Models;
using System.Data.SqlClient;

namespace ServvistaWebAppAPI.Services
{
    public class TechNotificationBackgroundService : BackgroundService
    {
        private readonly ILogger<TechNotificationBackgroundService> _logger;
        private readonly IConfiguration _configuration;
        private readonly IHubContext<NotificationHub> _hubContext;
        private readonly IServiceScopeFactory _scopeFactory;

        public TechNotificationBackgroundService(
            ILogger<TechNotificationBackgroundService> logger,
            IConfiguration configuration,
            IHubContext<NotificationHub> hubContext,
            IServiceScopeFactory scopeFactory)
        {
            _logger = logger;
            _configuration = configuration;
            _hubContext = hubContext;
            _scopeFactory = scopeFactory;
        }

        protected override async Task ExecuteAsync(
            CancellationToken stoppingToken)
        {
            _logger.LogInformation(
                "Technician notification background service started.");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await UpdateEmptyNotificationStatusAsync(
                        stoppingToken);

                    await ProcessAndSendNotificationsAsync(
                        stoppingToken);
                }
                catch (OperationCanceledException)
                    when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(
                        ex,
                        "Unexpected error occurred while processing technician notifications.");
                }

                try
                {
                    await Task.Delay(
                        TimeSpan.FromSeconds(5),
                        stoppingToken);
                }
                catch (OperationCanceledException)
                    when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
            }

            _logger.LogInformation(
                "Technician notification background service stopped.");
        }

        private async Task UpdateEmptyNotificationStatusAsync(
            CancellationToken cancellationToken)
        {
            var connectionString =
                _configuration.GetConnectionString(
                    "DefaultConnection");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "DefaultConnection is not configured.");
            }

            await using var connection =
                new SqlConnection(connectionString);

            await connection.OpenAsync(cancellationToken);

            const string query = """
                UPDATE TBL_DAILY_JOBS
                SET
                    IS_TECH_NOTIFIED =
                        ISNULL(IS_TECH_NOTIFIED, 0),

                    IS_SIGNALR_SENT =
                        ISNULL(IS_SIGNALR_SENT, 0),

                    IS_PUSH_SENT =
                        ISNULL(IS_PUSH_SENT, 0)

                WHERE JOB_STATUS = 'TECH ALLOCATED'
                  AND DJ_DATE >= CAST(GETDATE() AS DATE)
                  AND
                  (
                      IS_TECH_NOTIFIED IS NULL
                      OR IS_SIGNALR_SENT IS NULL
                      OR IS_PUSH_SENT IS NULL
                  );
                """;

            await connection.ExecuteAsync(
                new CommandDefinition(
                    query,
                    cancellationToken: cancellationToken));
        }

        private async Task ProcessAndSendNotificationsAsync(
            CancellationToken cancellationToken)
        {
            var connectionString =
                _configuration.GetConnectionString(
                    "DefaultConnection");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "DefaultConnection is not configured.");
            }

            await using var connection =
                new SqlConnection(connectionString);

            await connection.OpenAsync(cancellationToken);

            const string query = """
                SELECT
                    DJ_ID,
                    DJ_DATE,
                    SERIAL_NO,
                    MACHINE_REF_NO,
                    NOTE,
                    CONTACT_TYPE,
                    CUS_ADD1,
                    CUS_ADD2,
                    CUS_ADD3,
                    CUS_CONTACT,
                    CUS_TEL_NO,
                    CUS_SMS_NO,
                    MACHINE_MODEL_ID,
                    MACHINE_MODEL_NAME,
                    CUS_STATUS,
                    TECH_CODE,
                    IS_TECH_NOTIFIED
                FROM TBL_DAILY_JOBS
                WHERE ISNULL(IS_TECH_NOTIFIED, 0) = 0
                  AND JOB_STATUS = 'TECH ALLOCATED'
                  AND TECH_CODE IS NOT NULL
                  AND LTRIM(RTRIM(TECH_CODE)) <> '';
                """;

            var jobs = (
                await connection.QueryAsync<BreakdownModel>(
                    new CommandDefinition(
                        query,
                        cancellationToken:
                            cancellationToken)))
                .ToList();

            if (jobs.Count == 0)
            {
                return;
            }

            foreach (var job in jobs)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(job.TECH_CODE))
                {
                    continue;
                }

                var signalRSent = false;
                var pushSent = false;

                signalRSent =
                    await SendSignalRNotificationAsync(
                        job,
                        cancellationToken);

                pushSent =
                    await SendPushNotificationAsync(
                        job,
                        cancellationToken);

                await UpdateNotificationStatusAsync(
                    connection,
                    job,
                    signalRSent,
                    pushSent,
                    cancellationToken);
            }
        }

        private async Task<bool> SendSignalRNotificationAsync(
            BreakdownModel job,
            CancellationToken cancellationToken)
        {
            try
            {
                var technicianCode =
                    job.TECH_CODE.Trim();

                await _hubContext.Clients
                    .Group(technicianCode)
                    .SendAsync(
                        "ReceivingNotifications",
                        job,
                        cancellationToken:
                            cancellationToken);

                _logger.LogInformation(
                    "SignalR notification processed for job {JobId}, technician {TechCode}.",
                    job.DJ_ID,
                    technicianCode);

                return true;
            }
            catch (OperationCanceledException)
                when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "SignalR notification failed for job {JobId}, technician {TechCode}.",
                    job.DJ_ID,
                    job.TECH_CODE);

                return false;
            }
        }

        private async Task<bool> SendPushNotificationAsync(
            BreakdownModel job,
            CancellationToken cancellationToken)
        {
            try
            {
                using var scope =
                    _scopeFactory.CreateScope();

                var pushNotificationService =
                    scope.ServiceProvider
                        .GetRequiredService<
                            IPushNotificationService>();

                var pushSent =
                    await pushNotificationService
                        .SendJobAssignedNotificationAsync(
                            job,
                            cancellationToken);

                if (pushSent)
                {
                    _logger.LogInformation(
                        "Push notification sent for job {JobId}, technician {TechCode}.",
                        job.DJ_ID,
                        job.TECH_CODE);
                }
                else
                {
                    _logger.LogWarning(
                        "Push notification was not sent for job {JobId}, technician {TechCode}. No active device token may be available.",
                        job.DJ_ID,
                        job.TECH_CODE);
                }

                return pushSent;
            }
            catch (OperationCanceledException)
                when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "Push notification failed for job {JobId}, technician {TechCode}.",
                    job.DJ_ID,
                    job.TECH_CODE);

                return false;
            }
        }

        private async Task UpdateNotificationStatusAsync(
            SqlConnection connection,
            BreakdownModel job,
            bool signalRSent,
            bool pushSent,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(job.DJ_ID))
            {
                _logger.LogWarning(
                    "Notification status could not be updated because DJ_ID is empty.");

                return;
            }

            if (signalRSent)
            {
                const string updateSignalRQuery = """
                    UPDATE TBL_DAILY_JOBS
                    SET
                        IS_SIGNALR_SENT = 1,
                        SIGNALR_SENT_AT = GETDATE()
                    WHERE DJ_ID = @JobId;
                    """;

                await connection.ExecuteAsync(
                    new CommandDefinition(
                        updateSignalRQuery,
                        new
                        {
                            JobId = job.DJ_ID
                        },
                        cancellationToken:
                            cancellationToken));
            }

            if (pushSent)
            {
                const string updatePushQuery = """
                    UPDATE TBL_DAILY_JOBS
                    SET
                        IS_PUSH_SENT = 1,
                        PUSH_SENT_AT = GETDATE()
                    WHERE DJ_ID = @JobId;
                    """;

                await connection.ExecuteAsync(
                    new CommandDefinition(
                        updatePushQuery,
                        new
                        {
                            JobId = job.DJ_ID
                        },
                        cancellationToken:
                            cancellationToken));
            }

            /*
             * SignalR හෝ Push notification එකක් successful නම්
             * overall notification status එක complete ලෙස mark කරනවා.
             */
            if (signalRSent || pushSent)
            {
                const string updateOverallQuery = """
                    UPDATE TBL_DAILY_JOBS
                    SET IS_TECH_NOTIFIED = 1
                    WHERE DJ_ID = @JobId;
                    """;

                await connection.ExecuteAsync(
                    new CommandDefinition(
                        updateOverallQuery,
                        new
                        {
                            JobId = job.DJ_ID
                        },
                        cancellationToken:
                            cancellationToken));

                _logger.LogInformation(
                    "Notification status updated for job {JobId}. SignalR: {SignalRSent}, Push: {PushSent}.",
                    job.DJ_ID,
                    signalRSent,
                    pushSent);
            }
            else
            {
                _logger.LogWarning(
                    "Both SignalR and Push notification failed for job {JobId}. It will be retried.",
                    job.DJ_ID);
            }
        }
    }
}