using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ServvistaWebAppAPI.Models;
using System.Data.SqlClient;

namespace ServvistaWebAppAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ServiceSchedule01Controller : ControllerBase
    {
        [HttpGet("getMonthlyVisits")]
        public IActionResult GetMonthlyVisits(string techCode)
        {
            try
            {
                List<ServiceVisitMonthlyInfo> serviceVisitMonthlyInfos = new List<ServiceVisitMonthlyInfo>();
                string connectionString = @"Data Source=Gagan_GD\SQLEXPRESS;Initial Catalog=DB_A67CC4_Servvistagcp;User ID=gestetner;Password=gocl248;";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    string query = @"
                    SELECT 
                      SV.VisitID AS rowId,
                      MT.CUS_CODE AS customerID, 
                      MT.CUS_NAME AS customerName, 
                      MT.M_LOC_CONTACT_NAME AS contactPerson, 
                      MT.M_LOC_CONTACT_NO AS customerTelephone,
                      MT.M_LOC1 AS machineLocation01, 
                      MT.M_LOC2 AS machineLocation02, 
                      MT.M_LOC3 AS machineLocation03, 
                      MT.SERIAL_NO AS serialNo, 
                      MT.MACHINE_REF_CODE AS machineRefNo, 
                      MT.MACHINE_MODE AS machineModel, SV.ExpectedDate, SV.VisitNo
                      FROM TBL_MACHINE_TRANSACTION MT
                      INNER JOIN TBL_SERVICE_SCHEDULE SS
                      INNER JOIN SERVICE_VISITS SV
                      ON SV.ScheduleID = SS.scheduleID
                      ON MT.MACHINE_REF_CODE = SS.machineRefNo
                    WHERE MT.TO_CODE = @techCode AND ExpectedDate BETWEEN @startDate AND @endDate";

                    DateTime today = DateTime.Now.Date;
                    DateTime startDate = new DateTime(today.Year, today.Month, 1); 
                    DateTime endDate = startDate.AddMonths(1).AddDays(-1);
                    var result = connection.Query<ServiceVisitMonthlyInfo>(query, new { 
                        startDate = startDate,
                        endDate = endDate,
                        techCode = techCode
                    });

                    serviceVisitMonthlyInfos = result.ToList();

                    connection.Open(); 
                }
                return Ok(serviceVisitMonthlyInfos); 
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
