using Dapper;
using ServvistaWebAppAPI.Models;
using System.Data.SqlClient;

namespace ServvistaWebAppAPI.Repositories
{
    public class PushSubscriptionRepository
    {
        private readonly IConfiguration _configuration;

        public PushSubscriptionRepository(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString => _configuration.GetConnectionString("DefaultConnection")!;

        public async Task SaveAsync(PushSubscriptionRequest request)
        {
            using var connection = new SqlConnection(ConnectionString);

            // One row per endpoint. If the same device re-subscribes (e.g. after
            // clearing site data) we just refresh the keys against that endpoint.
            const string query = @"
                MERGE TBL_PUSH_SUBSCRIPTIONS AS target
                USING (SELECT @Endpoint AS ENDPOINT) AS src
                ON target.ENDPOINT = src.ENDPOINT
                WHEN MATCHED THEN
                    UPDATE SET TECH_CODE = @TechCode, P256DH = @P256dh, AUTH = @Auth,
                               USER_AGENT = @UserAgent, CR_DATE = GETDATE()
                WHEN NOT MATCHED THEN
                    INSERT (TECH_CODE, ENDPOINT, P256DH, AUTH, USER_AGENT, CR_DATE)
                    VALUES (@TechCode, @Endpoint, @P256dh, @Auth, @UserAgent, GETDATE());";

            await connection.ExecuteAsync(query, new
            {
                request.TechCode,
                request.Endpoint,
                request.Keys.P256dh,
                request.Keys.Auth,
                request.UserAgent
            });
        }

        public async Task RemoveAsync(string endpoint)
        {
            using var connection = new SqlConnection(ConnectionString);
            await connection.ExecuteAsync(
                "DELETE FROM TBL_PUSH_SUBSCRIPTIONS WHERE ENDPOINT = @Endpoint",
                new { Endpoint = endpoint });
        }

        public async Task<List<StoredPushSubscription>> GetByTechCodeAsync(string techCode)
        {
            using var connection = new SqlConnection(ConnectionString);
            var result = await connection.QueryAsync<StoredPushSubscription>(
                "SELECT ID, TECH_CODE, ENDPOINT, P256DH, AUTH FROM TBL_PUSH_SUBSCRIPTIONS WHERE TECH_CODE = @TechCode",
                new { TechCode = techCode });
            return result.ToList();
        }

        // Called when the push service reports a subscription is dead (HTTP 404/410).
        public async Task RemoveByIdAsync(int id)
        {
            using var connection = new SqlConnection(ConnectionString);
            await connection.ExecuteAsync("DELETE FROM TBL_PUSH_SUBSCRIPTIONS WHERE ID = @Id", new { Id = id });
        }
    }
}
