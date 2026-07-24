using Dapper;
using FirebaseAdmin.Messaging;
using Microsoft.Data.SqlClient;
using ServvistaWebAppAPI.Classes;
using ServvistaWebAppAPI.Models;

namespace ServvistaWebAppAPI.Services
{
    public class FirebasePushNotificationService : IPushNotificationService
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<FirebasePushNotificationService> _logger;

        public FirebasePushNotificationService(
            IConfiguration configuration,
            ILogger<FirebasePushNotificationService> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<bool> SendJobAssignedNotificationAsync(
            BreakdownModel job,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(job.TECH_CODE))
            {
                return false;
            }

            var connectionString =
                _configuration.GetConnectionString("DefaultConnection");

            await using var connection =
                new SqlConnection(connectionString);

            const string tokenQuery = """
                SELECT FCM_TOKEN
                FROM TBL_TECH_PUSH_DEVICE
                WHERE TECH_CODE = @TechCode
                  AND IS_ACTIVE = 1
                  AND FCM_TOKEN IS NOT NULL
                  AND LTRIM(RTRIM(FCM_TOKEN)) <> ''
                """;

            var tokens = (await connection.QueryAsync<string>(
                tokenQuery,
                new
                {
                    TechCode = job.TECH_CODE.Trim()
                }))
                .Distinct()
                .ToList();

            if (tokens.Count == 0)
            {
                _logger.LogInformation(
                    "No active FCM token found for technician {TechCode}",
                    job.TECH_CODE);

                return false;
            }

            var customerName =
                string.IsNullOrWhiteSpace(job.CUS_CONTACT)
                    ? "Customer"
                    : job.CUS_CONTACT;

            var machineName =
                string.IsNullOrWhiteSpace(job.MACHINE_MODEL_NAME)
                    ? job.SERIAL_NO
                    : job.MACHINE_MODEL_NAME;

            var message = new MulticastMessage
            {
                Tokens = tokens,

                Notification = new Notification
                {
                    Title = "New Service Job Assigned",
                    Body = $"{customerName} - {machineName}"
                },

                Data = new Dictionary<string, string>
                {
                    ["type"] = "JOB_ASSIGNED",
                    ["jobId"] = job.DJ_ID ?? string.Empty,
                    ["techCode"] = job.TECH_CODE ?? string.Empty,
                    ["serialNo"] = job.SERIAL_NO ?? string.Empty,
                    ["machineRefNo"] = job.MACHINE_REF_NO ?? string.Empty,
                    ["machineModel"] = job.MACHINE_MODEL_NAME ?? string.Empty,
                    ["url"] =
                        $"/technician-dashboard?jobId={Uri.EscapeDataString(job.DJ_ID ?? string.Empty)}"
                },

                Webpush = new WebpushConfig
                {
                    Headers = new Dictionary<string, string>
                    {
                        ["Urgency"] = "high"
                    },

                    Notification = new WebpushNotification
                    {
                        Icon = "/icon.png",
                        Badge = "/icon-dark-32x32.png",
                        RequireInteraction = true,
                        Tag = $"job-{job.DJ_ID}"
                    },

                    FcmOptions = new WebpushFcmOptions
                    {
                        Link =
                            $"/technician-dashboard?jobId={Uri.EscapeDataString(job.DJ_ID ?? string.Empty)}"
                    }
                }
            };

            try
            {
                var response =
                    await FirebaseMessaging.DefaultInstance.SendEachForMulticastAsync(
                        message,
                        cancellationToken);

                _logger.LogInformation(
                    "FCM notification sent for job {JobId}. Success: {SuccessCount}, Failed: {FailureCount}",
                    job.DJ_ID,
                    response.SuccessCount,
                    response.FailureCount);

                await DisableInvalidTokensAsync(
                    tokens,
                    response,
                    connection,
                    cancellationToken);

                return response.SuccessCount > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "FCM send failed for job {JobId}",
                    job.DJ_ID);

                return false;
            }
        }

        private static async Task DisableInvalidTokensAsync(
            IReadOnlyList<string> tokens,
            BatchResponse response,
            SqlConnection connection,
            CancellationToken cancellationToken)
        {
            var invalidTokens = new List<string>();

            for (var index = 0; index < response.Responses.Count; index++)
            {
                var sendResponse = response.Responses[index];

                if (sendResponse.IsSuccess)
                {
                    continue;
                }

                var messagingException =
                    sendResponse.Exception as FirebaseMessagingException;

                if (messagingException?.MessagingErrorCode is
                    MessagingErrorCode.Unregistered or
                    MessagingErrorCode.InvalidArgument)
                {
                    invalidTokens.Add(tokens[index]);
                }
            }

            if (invalidTokens.Count == 0)
            {
                return;
            }

            const string disableSql = """
                UPDATE TBL_TECH_PUSH_DEVICE
                SET IS_ACTIVE = 0,
                    UPDATED_AT = GETDATE()
                WHERE FCM_TOKEN IN @Tokens
                """;

            await connection.ExecuteAsync(
                new CommandDefinition(
                    disableSql,
                    new { Tokens = invalidTokens },
                    cancellationToken: cancellationToken));
        }
    }
}