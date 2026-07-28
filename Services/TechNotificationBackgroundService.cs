using Dapper;
using ServvistaWebAppAPI.Models;
using System.Data.SqlClient;

namespace ServvistaWebAppAPI.Services
{
    // Safety net only: catches any job whose instant push (triggered from the
    // legacy AddCall endpoint) never arrived, or where push failed because the
    // technician's subscription hadn't saved yet at the time.
    public class TechNotificationBackgroundService : BackgroundService
    {
        private readonly ILogger<TechNotificationBackgroundService> _logger;
        private readonly IConfiguration _configuration;
        private readonly IServiceScopeFactory _scopeFactory;

        public TechNotificationBackgroundService(
            ILogger<TechNotificationBackgroundService> logger,
            IConfiguration configuration,
            IServiceScopeFactory scopeFactory)
        {
            _logger = logger;
            _configuration = configuration;
            _scopeFactory = scopeFactory;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("Technician notification background service started.");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await UpdateEmptyTechNotifiedAsync(stoppingToken);
                    await ProcessAndSendNotificationsAsync(stoppingToken);
                }
                catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Unexpected error occurred while processing technician notifications.");
                }

                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken);
                }
                catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
            }

            _logger.LogInformation("Technician notification background service stopped.");
        }

        private async Task UpdateEmptyTechNotifiedAsync(CancellationToken cancellationToken)
        {
            var connectionString = _configuration.GetConnectionString("DefaultConnection");
            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            const string query = """
                UPDATE TBL_DAILY_JOBS
                SET IS_TECH_NOTIFIED = '0'
                WHERE IS_TECH_NOTIFIED IS NULL
                  AND JOB_STATUS = 'TECH ALLOCATED'
                  AND DJ_DATE >= CAST(GETDATE() AS DATE)
                """;

            await connection.ExecuteAsync(new CommandDefinition(query, cancellationToken: cancellationToken));
        }

        private async Task ProcessAndSendNotificationsAsync(CancellationToken cancellationToken)
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
                WHERE IS_TECH_NOTIFIED = '0'
                """;

            var jobs = (await connection.QueryAsync<BreakdownModel>(
                new CommandDefinition(query, cancellationToken: cancellationToken))).ToList();

            if (jobs.Count == 0) return;

            var sentJobIds = new List<string>();

            foreach (var job in jobs)
            {
                if (cancellationToken.IsCancellationRequested) break;
                if (string.IsNullOrWhiteSpace(job.TECH_CODE)) continue;

                using var scope = _scopeFactory.CreateScope();
                var notifier = scope.ServiceProvider.GetRequiredService<JobNotificationService>();

                var pushHandled = await notifier.NotifyJobAsync(job, cancellationToken);
                var isStale = (DateTime.Now - job.DJ_DATE) > TimeSpan.FromSeconds(30);

                if (pushHandled || isStale)
                {
                    sentJobIds.Add(job.DJ_ID);
                }
            }

            if (sentJobIds.Count == 0) return;

            await connection.ExecuteAsync(new CommandDefinition(
                "UPDATE TBL_DAILY_JOBS SET IS_TECH_NOTIFIED = '1' WHERE DJ_ID IN @Ids",
                new { Ids = sentJobIds.Distinct().ToList() },
                cancellationToken: cancellationToken));
        }
    }
}
