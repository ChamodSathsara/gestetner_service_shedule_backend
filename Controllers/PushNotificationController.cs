using Dapper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using ServvistaWebAppAPI.Models.Dtos;

namespace ServvistaWebAppAPI.Controllers
{
    [ApiController]
    [Route("api/push-notifications")]
    [Authorize]
    public class PushNotificationController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<PushNotificationController> _logger;

        public PushNotificationController(
            IConfiguration configuration,
            ILogger<PushNotificationController> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        [HttpPost("register")]
        public async Task<IActionResult> RegisterDevice(
            RegisterPushDeviceRequest request)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            /*
             * Final implementation එකේ TechCode request body එකෙන්
             * ගන්නවාට වඩා JWT claim එකෙන් ගන්න.
             */
            var techCode = request.TechCode.Trim();
            var token = request.FcmToken.Trim();

            var connectionString =
                _configuration.GetConnectionString("DefaultConnection");

            await using var connection =
                new SqlConnection(connectionString);

            const string sql = """
                IF EXISTS
                (
                    SELECT 1
                    FROM TBL_TECH_PUSH_DEVICE
                    WHERE FCM_TOKEN = @FcmToken
                )
                BEGIN
                    UPDATE TBL_TECH_PUSH_DEVICE
                    SET TECH_CODE = @TechCode,
                        COMPANY_ID = @CompanyId,
                        DEVICE_NAME = @DeviceName,
                        DEVICE_TYPE = @DeviceType,
                        IS_ACTIVE = 1,
                        UPDATED_AT = GETDATE(),
                        LAST_USED_AT = GETDATE()
                    WHERE FCM_TOKEN = @FcmToken
                END
                ELSE
                BEGIN
                    INSERT INTO TBL_TECH_PUSH_DEVICE
                    (
                        TECH_CODE,
                        COMPANY_ID,
                        FCM_TOKEN,
                        DEVICE_NAME,
                        DEVICE_TYPE,
                        IS_ACTIVE,
                        CREATED_AT,
                        LAST_USED_AT
                    )
                    VALUES
                    (
                        @TechCode,
                        @CompanyId,
                        @FcmToken,
                        @DeviceName,
                        @DeviceType,
                        1,
                        GETDATE(),
                        GETDATE()
                    )
                END
                """;

            await connection.ExecuteAsync(sql, new
            {
                TechCode = techCode,
                request.CompanyId,
                FcmToken = token,
                request.DeviceName,
                request.DeviceType
            });

            return Ok(new
            {
                message = "Push notification device registered successfully"
            });
        }

        [HttpPost("unregister")]
        public async Task<IActionResult> UnregisterDevice(
            UnregisterPushDeviceRequest request)
        {
            var connectionString =
                _configuration.GetConnectionString("DefaultConnection");

            await using var connection =
                new SqlConnection(connectionString);

            const string sql = """
                UPDATE TBL_TECH_PUSH_DEVICE
                SET IS_ACTIVE = 0,
                    UPDATED_AT = GETDATE()
                WHERE FCM_TOKEN = @FcmToken
                """;

            await connection.ExecuteAsync(sql, new
            {
                FcmToken = request.FcmToken.Trim()
            });

            return Ok(new
            {
                message = "Push notification device unregistered successfully"
            });
        }
    }
}