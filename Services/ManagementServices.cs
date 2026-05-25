using Dapper;
using ServvistaWebAppAPI.Models;
using ServvistaWebAppAPI.Models.Dtos;
using ServvistaWebAppAPI.Repositories;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;


namespace ServvistaWebAppAPI.Services
{
    public class ManagementServices : IManagementRepository
    {
        private readonly string _connectionString;
        private readonly ITenantService _tenantService;

        public ManagementServices(IConfiguration config, ITenantService tenantService)
        {
            _tenantService = tenantService;
            _connectionString = _tenantService.GetConnectionString();
        }

        public JobCountAndRateDto GetJobCountAndRate()
        {
            string jobCountQuery = @"
            SELECT COUNT(*) 
            FROM TBL_DAILY_JOBS
            WHERE DJ_DATE >= DATEADD(DAY, -30, GETDATE())";

            string successCountQuery = @"
            SELECT COUNT(*) 
            FROM TBL_DAILY_JOBS 
            WHERE (JOB_STATUS = 'COMPLETED' OR JOB_STATUS = 'COMPLETE')
            AND DJ_DATE >= DATEADD(DAY, -30, GETDATE())";

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                int jobCount = 0;
                int successCount = 0;

                using (var cmd = new SqlCommand(jobCountQuery, connection))
                {
                    jobCount = (int)cmd.ExecuteScalar();
                }

                using (var cmd = new SqlCommand(successCountQuery, connection))
                {
                    successCount = (int)cmd.ExecuteScalar();
                }

                double jobRate = jobCount > 0 ? (double)successCount / jobCount * 100 : 0;

                return new JobCountAndRateDto
                {
                    JobCount = jobCount,
                    SuccessCount = successCount,
                    JobRate = Math.Round(jobRate, 2)
                };
            }
        }

        public JobCountAndRateDto GetServiceCountAndRate()
        {
            string servicesCountQuery = @"
                SELECT COUNT(*)
                FROM TBL_SERVICE_SCEDULE_UPDATE
                WHERE
                    (
                        EXPT_SV1 BETWEEN DATEADD(DAY, -30, GETDATE()) AND GETDATE()
                    )
                    OR
                    (
                        EXPT_SV2 BETWEEN DATEADD(DAY, -30, GETDATE()) AND GETDATE()
                    )
                    OR
                    (
                        EXPT_SV3 BETWEEN DATEADD(DAY, -30, GETDATE()) AND GETDATE()
                    )
                    OR
                    (
                        EXPT_SV4 BETWEEN DATEADD(DAY, -30, GETDATE()) AND GETDATE()
                    )
                    OR
                    (
                        EXPT_SV5 BETWEEN DATEADD(DAY, -30, GETDATE()) AND GETDATE()
                    )
                    OR
                    (
                        EXPT_SV6 BETWEEN DATEADD(DAY, -30, GETDATE()) AND GETDATE()
                    );
                    ";

            string successCountQuery = @"
                SELECT COUNT(*) AS completeRow
                FROM TBL_SERVICE_SCEDULE_UPDATE s
                INNER JOIN MTBL_CUSTOMER_MASTER c
                    ON c.CUS_CODE = s.CUS_ID
                CROSS APPLY
                (
                    VALUES
                        ('EXPT_SV1', s.EXPT_SV1, s.SV1),
                        ('EXPT_SV2', s.EXPT_SV2, s.SV2),
                        ('EXPT_SV3', s.EXPT_SV3, s.SV3),
                        ('EXPT_SV4', s.EXPT_SV4, s.SV4),
                        ('EXPT_SV5', s.EXPT_SV5, s.SV5),
                        ('EXPT_SV6', s.EXPT_SV6, s.SV6)
                ) v (VisitNo, ExpectedDate, ActualVisit)
                WHERE v.ExpectedDate IS NOT NULL
                  AND v.ExpectedDate >= DATEADD(DAY, -30, GETDATE())
                  AND v.ActualVisit IS NOT NULL;
            ";


            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                int jobCount = 0;
                int successCount = 0;

                using (var cmd = new SqlCommand(servicesCountQuery, connection))
                {
                    jobCount = (int)cmd.ExecuteScalar();
                }

                using (var cmd = new SqlCommand(successCountQuery, connection))
                {
                    successCount = (int)cmd.ExecuteScalar();
                }

                double jobRate = jobCount > 0 ? (double)successCount / jobCount * 100 : 0;

                return new JobCountAndRateDto
                {
                    JobCount = jobCount,
                    SuccessCount = successCount,
                    JobRate = Math.Round(jobRate, 2)
                };
            }
        }

        public List<BreakdownModel> GetPendingJobs()
        {
            string query = @"
            SELECT DJ_ID, SERIAL_NO, MACHINE_REF_NO, CUS_NAME, CUS_ADD1, CUS_ADD2, CUS_ADD3, CUS_CONTACT, 
            CUS_SMS_NO AS CUS_TEL_NO, TEAM_ID, TEAM_NAME, DJ_DATE, TECH_CODE, TECH_MOBILE, MACHINE_MODEL_ID, 
            MACHINE_MODEL_NAME, CUS_STATUS, JOB_STATUS, NOTE
            FROM TBL_DAILY_JOBS 
            WHERE JOB_STATUS = 'TECH ALLOCATED'";

            var jobs = new List<BreakdownModel>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var cmd = new SqlCommand(query, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        jobs.Add(new BreakdownModel
                        {
                            DJ_ID = reader["DJ_ID"]?.ToString(),
                            SERIAL_NO = reader["SERIAL_NO"]?.ToString(),
                            MACHINE_REF_NO = reader["MACHINE_REF_NO"]?.ToString(),
                            CUS_NAME = reader["CUS_NAME"]?.ToString(),
                            CUS_ADD1 = reader["CUS_ADD1"]?.ToString(),
                            CUS_ADD2 = reader["CUS_ADD2"]?.ToString(),
                            CUS_ADD3 = reader["CUS_ADD3"]?.ToString(),
                            CUS_CONTACT = reader["CUS_CONTACT"]?.ToString(),
                            CUS_TEL_NO = reader["CUS_TEL_NO"]?.ToString(),
                            TEAM_ID = reader["TEAM_ID"]?.ToString(),
                            TEAM_NAME = reader["TEAM_NAME"]?.ToString(),
                            DJ_DATE = reader["DJ_DATE"] == DBNull.Value ? DateTime.MinValue : Convert.ToDateTime(reader["DJ_DATE"]),
                            TECH_CODE = reader["TECH_CODE"]?.ToString(),
                            TECH_NAME = null,         // not in query, default null
                            TECH_MOBILE = reader["TECH_MOBILE"]?.ToString(),
                            MACHINE_MODEL_ID = reader["MACHINE_MODEL_ID"]?.ToString(),
                            MACHINE_MODEL_NAME = reader["MACHINE_MODEL_NAME"]?.ToString(),
                            CUS_STATUS = reader["CUS_STATUS"]?.ToString(),
                            NOTE = reader["NOTE"]?.ToString(),
                            JOB_STATUS = reader["JOB_STATUS"]?.ToString(),
                            IS_TECH_NOTIFIED = false,       // not in query, default false
                            TYPE = null          // not in query, default null
                        });
                    }
                }
            }

            return jobs;
        }

        public List<ServiceVisitMonthlyInfo> GetPendingServices()
        {
            string query = @"
        SELECT
            s.T_ID AS rowId,
            s.CUS_ID AS customerID,
            c.CUS_NAME AS customerName,    
            c.CONTACT_PERSON AS contactPerson,
            c.TEL_NO AS customerTelephone,
            s.M_LOC1 AS machineLocation01, 
            s.M_LOC2 AS machineLocation02, 
            s.M_LOC3 AS machineLocation03, 
            s.MACHINE_REF AS machineRefNo,
            s.TECH_NAME AS techName,
            s.SERIAL_NO AS serialNo,
            v.VisitNo AS expectedVisitNo,
            CONVERT(char(10), v.ExpectedDate, 120) AS expectedVisitDate,
            CASE
                WHEN v.ActualVisit IS NOT NULL THEN 'COMPLETED'
                ELSE 'PENDING'
            END AS VisitStatus
        FROM TBL_SERVICE_SCEDULE_UPDATE s
        INNER JOIN MTBL_CUSTOMER_MASTER c
            ON c.CUS_CODE = s.CUS_ID
        CROSS APPLY
        (
            VALUES
                ('EXPT_SV1', s.EXPT_SV1, s.SV1),
                ('EXPT_SV2', s.EXPT_SV2, s.SV2),
                ('EXPT_SV3', s.EXPT_SV3, s.SV3),
                ('EXPT_SV4', s.EXPT_SV4, s.SV4),
                ('EXPT_SV5', s.EXPT_SV5, s.SV5),
                ('EXPT_SV6', s.EXPT_SV6, s.SV6)
        ) v(VisitNo, ExpectedDate, ActualVisit)
        WHERE v.ExpectedDate IS NOT NULL
        AND v.ExpectedDate >= DATEADD(DAY, -7, GETDATE())
        ORDER BY v.ExpectedDate, v.VisitNo";

            var services = new List<ServiceVisitMonthlyInfo>();

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var cmd = new SqlCommand(query, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        services.Add(new ServiceVisitMonthlyInfo
                        {
                            RowId = reader["rowId"] == DBNull.Value ? 0 : Convert.ToInt32(reader["rowId"]),
                            customerID = reader["customerID"]?.ToString(),
                            customerName = reader["customerName"]?.ToString(),
                            contactPerson = reader["contactPerson"] == DBNull.Value ? null : reader["contactPerson"].ToString(),
                            customerTelephone = reader["customerTelephone"]?.ToString(),
                            machineLocation01 = reader["machineLocation01"]?.ToString(),
                            machineLocation02 = reader["machineLocation02"]?.ToString(),
                            machineLocation03 = reader["machineLocation03"]?.ToString(),
                            machineRefNo = reader["machineRefNo"]?.ToString(),
                            techName = reader["techName"]?.ToString(),
                            serialNo = reader["serialNo"]?.ToString(),
                            expectedVisitNo = reader["expectedVisitNo"]?.ToString(),
                            expectedVisitDate = reader["expectedVisitDate"] == DBNull.Value
                            ? DateTime.MinValue
                            : DateTime.Parse(reader["expectedVisitDate"].ToString()),
                            VisitStatus = reader["VisitStatus"]?.ToString(),

                            // Not in query — set defaults
                            expectedVisitCount = 0,
                            machineModel = null
                        });
                    }
                }
            }

            return services;
        }

        public List<CompleteAndPendingPercentageDto> GetCompleteAndPendingJobPercentage()
        {
            string query = @"
            SELECT 
                SUM(CASE WHEN JOB_STATUS IN ('COMPLETE', 'COMPLETED') THEN 1 ELSE 0 END) AS completed,
                SUM(CASE WHEN JOB_STATUS = 'STARTED'        THEN 1 ELSE 0 END) AS started,
                SUM(CASE WHEN JOB_STATUS = 'TECH ALLOCATED' THEN 1 ELSE 0 END) AS pending,
	            SUM(CASE WHEN JOB_STATUS = 'CANCELLED' THEN 1 ELSE 0 END) AS cancel,
                COUNT(*) AS total
            FROM TBL_DAILY_JOBS WITH (NOLOCK)";


            Console.WriteLine("[GetCompleteAndPendingJobPercentage] Starting...");

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    Console.WriteLine("[GetCompleteAndPendingJobPercentage] Opening connection...");
                    connection.Open();
                    Console.WriteLine("[GetCompleteAndPndingJobPercentage] Connection opened successfully.");

                    using (var cmd = new SqlCommand(query, connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        Console.WriteLine("[GetCompleteAndPendingJobPercentage] Query executed.");

                        if (reader.Read())
                        {
                            int total = reader.GetInt32(reader.GetOrdinal("total"));
                            int completed = reader.GetInt32(reader.GetOrdinal("completed"));
                            int started = reader.GetInt32(reader.GetOrdinal("started"));
                            int pending = reader.GetInt32(reader.GetOrdinal("pending"));
                            int cansel = reader.GetInt32(reader.GetOrdinal("cancel"));

                            Console.WriteLine($"[GetCompleteAndPendingJobPercentage] Raw values => Total: {total}, Completed: {completed}, Started: {started}, Pending: {pending}");

                            int completedPct = total > 0 ? (int)Math.Round((double)completed * 100 / total) : 0;
                            int startedPct = total > 0 ? (int)Math.Round((double)started * 100 / total) : 0;
                            int pendingPct = total > 0 ? (int)Math.Round((double)pending * 100 / total) : 0;
                            int canselPct = total > 0 ? (int)Math.Round((double)cansel * 100 / total) : 0;

                            Console.WriteLine($"[GetCompleteAndPendingJobPercentage] Percentages => Completed: {completedPct}%, Started: {startedPct}%, Pending: {pendingPct}%");

                            return new List<CompleteAndPendingPercentageDto>
                            {
                                new CompleteAndPendingPercentageDto { name = "Completed", value = completedPct },
                                new CompleteAndPendingPercentageDto { name = "Started",   value = startedPct   },
                                new CompleteAndPendingPercentageDto { name = "Pending",   value = pendingPct   },
                                new CompleteAndPendingPercentageDto { name = "Cansel",   value = canselPct   }
                            };
                        }
                        else
                        {
                            Console.WriteLine("[GetCompleteAndPendingJobPercentage] WARNING: Query returned no rows.");
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                Console.WriteLine($"[GetCompleteAndPendingJobPercentage] SQL ERROR: {sqlEx.Message}");
                Console.WriteLine($"[GetCompleteAndPendingJobPercentage] SQL Error Number: {sqlEx.Number}");
                Console.WriteLine($"[GetCompleteAndPendingJobPercentage] SQL State: {sqlEx.State}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetCompleteAndPendingJobPercentage] UNEXPECTED ERROR: {ex.Message}");
                Console.WriteLine($"[GetCompleteAndPendingJobPercentage] StackTrace: {ex.StackTrace}");
                throw;
            }

            Console.WriteLine("[GetCompleteAndPendingJobPercentage] Returning fallback zeros.");
            return new List<CompleteAndPendingPercentageDto>
                {
                    new CompleteAndPendingPercentageDto { name = "Completed", value = 0 },
                    new CompleteAndPendingPercentageDto { name = "Started",   value = 0 },
                    new CompleteAndPendingPercentageDto { name = "Pending",   value = 0 }
                };
        }


        public List<CompleteAndPendingPercentageDto> GetCompleteAndPendingServicesPercentage()
        {
            string query = @"
            SELECT
                COUNT(*) AS total,
                SUM(CASE WHEN v.ActualVisit IS NOT NULL THEN 1 ELSE 0 END) AS completed,
                SUM(CASE WHEN v.ActualVisit IS NULL     THEN 1 ELSE 0 END) AS pending
            FROM TBL_SERVICE_SCEDULE_UPDATE s WITH (NOLOCK)
            INNER JOIN MTBL_CUSTOMER_MASTER c WITH (NOLOCK)
                ON c.CUS_CODE = s.CUS_ID
            CROSS APPLY
            (
                VALUES
                    ('EXPT_SV1', s.EXPT_SV1, s.SV1),
                    ('EXPT_SV2', s.EXPT_SV2, s.SV2),
                    ('EXPT_SV3', s.EXPT_SV3, s.SV3),
                    ('EXPT_SV4', s.EXPT_SV4, s.SV4),
                    ('EXPT_SV5', s.EXPT_SV5, s.SV5),
                    ('EXPT_SV6', s.EXPT_SV6, s.SV6)
            ) v (VisitNo, ExpectedDate, ActualVisit)
            WHERE v.ExpectedDate IS NOT NULL";

            Console.WriteLine("[GetCompleteAndPendingServicesPercentage] Starting...");

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    Console.WriteLine("[GetCompleteAndPendingServicesPercentage] Opening connection...");
                    connection.Open();
                    Console.WriteLine("[GetCompleteAndPendingServicesPercentage] Connection opened successfully.");

                    using (var cmd = new SqlCommand(query, connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        Console.WriteLine("[GetCompleteAndPendingServicesPercentage] Query executed.");

                        if (reader.Read())
                        {
                            int total = reader.GetInt32(reader.GetOrdinal("total"));
                            int completed = reader.GetInt32(reader.GetOrdinal("completed"));
                            int pending = reader.GetInt32(reader.GetOrdinal("pending"));

                            Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] Raw values => Total: {total}, Completed: {completed}, Pending: {pending}");

                            int completedPct = total > 0 ? (int)Math.Round((double)completed * 100 / total) : 0;
                            int pendingPct = total > 0 ? (int)Math.Round((double)pending * 100 / total) : 0;


                            Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] Percentages => Completed: {completedPct}%, Pending: {pendingPct}%");

                            return new List<CompleteAndPendingPercentageDto>
                            {
                                new CompleteAndPendingPercentageDto { name = "Completed", value = completedPct },
                                new CompleteAndPendingPercentageDto { name = "Started",   value = 0            },
                                new CompleteAndPendingPercentageDto { name = "Pending",   value = pendingPct   }
                            };
                        }
                        else
                        {
                            Console.WriteLine("[GetCompleteAndPendingServicesPercentage] WARNING: Query returned no rows.");
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] SQL ERROR: {sqlEx.Message}");
                Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] SQL Error Number: {sqlEx.Number}");
                Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] SQL State: {sqlEx.State}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] UNEXPECTED ERROR: {ex.Message}");
                Console.WriteLine($"[GetCompleteAndPendingServicesPercentage] StackTrace: {ex.StackTrace}");
                throw;
            }

            Console.WriteLine("[GetCompleteAndPendingServicesPercentage] Returning fallback zeros.");
            return new List<CompleteAndPendingPercentageDto>
            {
                new CompleteAndPendingPercentageDto { name = "Completed", value = 0 },
                new CompleteAndPendingPercentageDto { name = "Started",   value = 0 },
                new CompleteAndPendingPercentageDto { name = "Pending",   value = 0 }
            };
        }

        public List<LastYearJobPerformanceDto> GetLastWeekJobPerformance()
        {
            string query = @"
        SELECT 
            FORMAT(DJ_DATE, 'MM/dd') AS date,
            COUNT(CASE WHEN JOB_STATUS = 'TECH ALLOCATED'          THEN 1 END) AS pending,
            COUNT(CASE WHEN JOB_STATUS IN ('COMPLETED', 'COMPLETE') THEN 1 END) AS completed,
            COUNT(CASE WHEN JOB_STATUS = 'STARTED'                 THEN 1 END) AS started,
            COUNT(CASE WHEN JOB_STATUS = 'CANCELLED'               THEN 1 END) AS cancel,
            COUNT(*) AS total
        FROM TBL_DAILY_JOBS WITH (NOLOCK)
        WHERE DJ_DATE >= DATEADD(DAY, -10, GETDATE())
        GROUP BY CAST(DJ_DATE AS DATE)
        ORDER BY CAST(DJ_DATE AS DATE)";

            Console.WriteLine("[GetLastWeekJobPerformance] Starting...");

            var result = new List<LastYearJobPerformanceDto>();

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    Console.WriteLine("[GetLastWeekJobPerformance] Opening connection...");
                    connection.Open();
                    Console.WriteLine("[GetLastWeekJobPerformance] Connection opened successfully.");

                    using (var cmd = new SqlCommand(query, connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        Console.WriteLine("[GetLastWeekJobPerformance] Query executed.");

                        int idxDate = reader.GetOrdinal("date");
                        int idxPending = reader.GetOrdinal("pending");
                        int idxCompleted = reader.GetOrdinal("completed");
                        int idxStarted = reader.GetOrdinal("started");
                        int idxCancel = reader.GetOrdinal("cancel");
                        int idxTotal = reader.GetOrdinal("total");

                        while (reader.Read())
                        {
                            var row = new LastYearJobPerformanceDto
                            {
                                date = reader.IsDBNull(idxDate) ? null : reader.GetString(idxDate),
                                pending = reader.IsDBNull(idxPending) ? 0 : reader.GetInt32(idxPending),
                                completed = reader.IsDBNull(idxCompleted) ? 0 : reader.GetInt32(idxCompleted),
                                started = reader.IsDBNull(idxStarted) ? 0 : reader.GetInt32(idxStarted),
                                cancel = reader.IsDBNull(idxCancel) ? 0 : reader.GetInt32(idxCancel),
                                total = reader.IsDBNull(idxTotal) ? 0 : reader.GetInt32(idxTotal)
                            };

                            Console.WriteLine($"[GetLastWeekJobPerformance] Row => Date: {row.date}, Pending: {row.pending}, Completed: {row.completed}, Started: {row.started}, Cancel: {row.cancel}, Total: {row.total}");

                            result.Add(row);
                        }

                        Console.WriteLine($"[GetLastWeekJobPerformance] Total rows fetched: {result.Count}");
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                Console.WriteLine($"[GetLastWeekJobPerformance] SQL ERROR: {sqlEx.Message}");
                Console.WriteLine($"[GetLastWeekJobPerformance] SQL Error Number: {sqlEx.Number}");
                Console.WriteLine($"[GetLastWeekJobPerformance] SQL State: {sqlEx.State}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GetLastWeekJobPerformance] UNEXPECTED ERROR: {ex.Message}");
                Console.WriteLine($"[GetLastWeekJobPerformance] StackTrace: {ex.StackTrace}");
                throw;
            }

            return result;
        }

        public List<TechnicianPerformenceDto> GetTechniciansPerformence()
        {
            try
            {
                List<TechnicianPerformenceDto> technicianPerformance = new List<TechnicianPerformenceDto>();
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    string getTechniciansQuery = @"
                    SELECT TOP 5
                        t.TECH_CODE AS tech_id,
                        t.TECH_NAME AS name,

                        COALESCE(j.completedJobs, 0) AS completedJobs,
                        COALESCE(f.rating, 0) AS rating,
                        COALESCE(s.services, 0) AS services,
                        COALESCE(d.breakdowns, 0) AS breakdowns

                    FROM MTBL_TECH_OFFICERS t

                    -- Completed Jobs
                    LEFT JOIN (
                        SELECT 
                            COMPLETED_BY,
                            COUNT(*) AS completedJobs
                        FROM TBL_SCHEDULE_ACTIVITY
                        WHERE TYPE IN ('Job Recall', 'Job')
                        GROUP BY COMPLETED_BY
                    ) j ON j.COMPLETED_BY = t.TECH_CODE

                    -- Rating (FIXED HERE)
                    LEFT JOIN (
                        SELECT 
                            TECH_CODE,
                            AVG(TRY_CAST(RATING AS FLOAT)) AS rating
                        FROM TBL_SV_CUSTOMER_JOB_FEEDBACKS
                        GROUP BY TECH_CODE
                    ) f ON f.TECH_CODE = t.TECH_CODE

                    -- Services
                    LEFT JOIN (
                        SELECT 
                            COMPLETED_BY,
                            COUNT(*) AS services
                        FROM TBL_SCHEDULE_ACTIVITY
                        GROUP BY COMPLETED_BY
                    ) s ON s.COMPLETED_BY = t.TECH_CODE

                    -- Breakdowns
                    LEFT JOIN (
                        SELECT 
                            TECH_CODE,
                            COUNT(*) AS breakdowns
                        FROM TBL_DAILY_JOBS
                        GROUP BY TECH_CODE
                    ) d ON d.TECH_CODE = t.TECH_CODE

                    ORDER BY rating DESC";

                    technicianPerformance = connection.Query<TechnicianPerformenceDto>(getTechniciansQuery).ToList();
                }
                return technicianPerformance;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        //public List<TechnicianPerformenceDto> GetTechniciansPerformence()
        //{
        //    return new List<TechnicianPerformenceDto>
        //    {
        //        new TechnicianPerformenceDto
        //        {
        //            tech_id = "0000127",
        //            name = "Sudesh",
        //            completedJobs = 20, //total completed jobs - counts 
        //            rating = 4.3, //Average rating from customers - average counts
        //            services = 10, //total completed visits - counts
        //            breakdowns = 5 //total breakdowns assigned to technician - counts 
        //        },
        //        new TechnicianPerformenceDto
        //        {
        //            tech_id = "0000128",
        //            name = "Nimal",
        //            completedJobs = 15,
        //            rating = 3.8,
        //            services = 8,
        //            breakdowns = 7
        //        },
        //        new TechnicianPerformenceDto
        //        {
        //            tech_id = "0000129",
        //            name = "Kamal",
        //            completedJobs = 25,
        //            rating = 5.0,
        //            services = 12,
        //            breakdowns = 3
        //        },
        //        new TechnicianPerformenceDto
        //        {
        //            tech_id = "0000125",
        //            name = "Sudesh",
        //            completedJobs = 20,
        //            rating = 4.3,
        //            services = 10,
        //            breakdowns = 5
        //        },
        //        new TechnicianPerformenceDto
        //        {
        //            tech_id = "0000124",
        //            name = "Nimal",
        //            completedJobs = 15,
        //            rating = 3.8,
        //            services = 8,
        //            breakdowns = 7
        //        },
        //        new TechnicianPerformenceDto
        //        {
        //            tech_id = "0000123",
        //            name = "Kamal",
        //            completedJobs = 25,
        //            rating = 5.0,
        //            services = 12,
        //            breakdowns = 3
        //        }
        //    };
        //}

        public List<OldestDueDto> GetOldestDueJobs()
        {
            var results = new List<OldestDueDto>();

            string query = @"
            SELECT TOP 20
                DJ_ID, 
                SERIAL_NO, 
                MACHINE_REF_NO, 
                CUS_NAME, 
                CUS_ADD1, 
                CUS_ADD2, 
                CUS_ADD3, 
                CUS_CONTACT,
                CUS_SMS_NO AS CUS_TEL_NO, 
                TEAM_ID, 
                TEAM_NAME, 
                DJ_DATE, 
                TECH_CODE, 
                TECH_MOBILE, 
                MACHINE_MODEL_ID,
                MACHINE_MODEL_NAME, 
                CUS_STATUS, 
                JOB_STATUS, 
                NOTE
            FROM TBL_DAILY_JOBS
            WHERE JOB_STATUS = 'TECH ALLOCATED'
            AND DJ_DATE >= '2026-04-01'
            AND DJ_DATE < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))
            ORDER BY DJ_DATE ASC";

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var dto = new OldestDueDto
                            {
                                DjId = reader["DJ_ID"].ToString(),
                                SerialNo = reader["SERIAL_NO"].ToString(),
                                MachineRefNo = reader["MACHINE_REF_NO"].ToString(),
                                CustomerName = reader["CUS_NAME"].ToString(),
                                CustomerAddress = CombineAddress(
                                    reader["CUS_ADD1"].ToString(),
                                    reader["CUS_ADD2"].ToString(),
                                    reader["CUS_ADD3"].ToString()
                                ),
                                CustomerContact = reader["CUS_CONTACT"].ToString(),
                                CustomerTelNo = reader["CUS_TEL_NO"].ToString(),
                                TeamId = reader["TEAM_ID"].ToString(),
                                TeamName = reader["TEAM_NAME"].ToString(),
                                DjDate = Convert.ToDateTime(reader["DJ_DATE"]).ToString("yyyy-MM-dd"),
                                TechCode = reader["TECH_CODE"].ToString(),
                                TechMobile = reader["TECH_MOBILE"].ToString(),
                                MachineModelId = reader["MACHINE_MODEL_ID"].ToString(),
                                MachineModelName = reader["MACHINE_MODEL_NAME"].ToString(),
                                CustomerStatus = reader["CUS_STATUS"].ToString(),
                                JobStatus = reader["JOB_STATUS"].ToString(),
                                Note = reader["NOTE"].ToString(),
                                DaysOverdue = CalculateDaysOverdue(Convert.ToDateTime(reader["DJ_DATE"]))
                            };

                            results.Add(dto);
                        }
                    }
                }
            }

            return results;
        }

        // Helper method to combine address fields
        private string CombineAddress(string add1, string add2, string add3)
        {
            var addressParts = new[] { add1, add2, add3 }
                .Where(a => !string.IsNullOrWhiteSpace(a) && a != "-")
                .ToArray();

            return string.Join(", ", addressParts);
        }

        // Helper method to calculate days overdue
        private string CalculateDaysOverdue(DateTime djDate)
        {
            var daysOverdue = (DateTime.Now.Date - djDate.Date).Days;
            return daysOverdue.ToString();
        }

        public List<WarrantyDetailsDto> GetWarrantyDetails()
        {
            var results = new List<WarrantyDetailsDto>();

            string query = @"
        SELECT 
            TEAM,
            COUNT(CASE WHEN CUS_STATUS = 'NS' THEN 1 END) AS NS_COUNT,
            COUNT(CASE WHEN CUS_STATUS = 'MA' THEN 1 END) AS MA_COUNT,
            COUNT(CASE WHEN CUS_STATUS = 'FS' THEN 1 END) AS FS_COUNT,
            COUNT(CASE WHEN CUS_STATUS = 'EX' THEN 1 END) AS EX_COUNT,
            COUNT(*) AS TOTAL_COUNT
        FROM TBL_MACHINE_TRANSACTION
        WHERE MACHINE_TYPE IN ('MFP', 'DCP')
        AND YEAR(M_WARRANTY_END_DATE) = YEAR(GETDATE())
        GROUP BY TEAM
        ORDER BY TEAM";

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var dto = new WarrantyDetailsDto
                            {
                                Area = reader["TEAM"] != DBNull.Value
                                    ? reader["TEAM"].ToString()
                                    : "Unknown",  // Handle NULL team values
                                Ns = Convert.ToInt32(reader["NS_COUNT"]),
                                Ma = Convert.ToInt32(reader["MA_COUNT"]),
                                Fs = Convert.ToInt32(reader["FS_COUNT"]),
                                Ex = Convert.ToInt32(reader["EX_COUNT"]),
                                Total = Convert.ToInt32(reader["TOTAL_COUNT"])
                            };

                            results.Add(dto);
                        }
                    }
                }
            }

            return results;
        }


        public List<TechnicianDetailsDto> GetTechnicianDetails()
        {
            var result = new List<TechnicianDetailsDto>();
            string query = @"
        SELECT 
            -- Basic Info from MTBL_TECH_OFFICERS
            t.TECH_CODE as techCode,
            t.TECH_NAME as name,
            t.EMAIL as email,
            t.MOBILE_NO as phone,
            t.AREA as zone,
    
            -- Job Statistics from TBL_DAILY_JOBS
            ISNULL(j.allJobs, 0) as allJobs,
            ISNULL(j.pendingJobs, 0) as pendingJobs,
            CASE 
                WHEN ISNULL(j.allJobs, 0) > 0 
                THEN ROUND(CAST(ISNULL(j.completedJobs, 0) AS FLOAT) / j.allJobs * 100, 2)
                ELSE 0 
            END as jobPerformancePercentage,
    
            -- Service Statistics from TBL_SERVICE_SCEDULE_UPDATE
            ISNULL(s.allServices, 0) as allServices,
            ISNULL(s.pendingServices, 0) as pendingServices,
            CASE 
                WHEN ISNULL(s.allServices, 0) > 0 
                THEN ROUND(CAST(ISNULL(s.completedServices, 0) AS FLOAT) / s.allServices * 100, 2)
                ELSE 0 
            END as servicePerformancePercentage
        FROM MTBL_TECH_OFFICERS t
        -- LEFT JOIN for Job Statistics
        LEFT JOIN (
            SELECT 
                TECH_CODE,
                COUNT(*) as allJobs,
                SUM(CASE WHEN JOB_STATUS = 'TECH ALLOCATED' THEN 1 ELSE 0 END) as pendingJobs,
                SUM(CASE WHEN JOB_STATUS IN ('COMPLETE', 'COMPLETED') THEN 1 ELSE 0 END) as completedJobs
            FROM TBL_DAILY_JOBS
            GROUP BY TECH_CODE
        ) j ON t.TECH_CODE = j.TECH_CODE
        -- LEFT JOIN for Service Statistics
        LEFT JOIN (
            SELECT 
                TECH_CODE,
                COUNT(DISTINCT T_ID) as allServices,
                SUM(CASE WHEN SV1_STATUS = 'pending' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV2_STATUS = 'pending' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV3_STATUS = 'pending' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV4_STATUS = 'pending' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV5_STATUS = 'pending' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV6_STATUS = 'pending' THEN 1 ELSE 0 END) as pendingServices,
                SUM(CASE WHEN SV1_STATUS = 'COMPLETED' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV2_STATUS = 'COMPLETED' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV3_STATUS = 'COMPLETED' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV4_STATUS = 'COMPLETED' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV5_STATUS = 'COMPLETED' THEN 1 ELSE 0 END) +
                SUM(CASE WHEN SV6_STATUS = 'COMPLETED' THEN 1 ELSE 0 END) as completedServices
            FROM TBL_SERVICE_SCEDULE_UPDATE
            GROUP BY TECH_CODE
        ) s ON t.TECH_CODE = s.TECH_CODE
        WHERE t.IS_ACTIVE = '1'
        ORDER BY t.TECH_NAME";

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var technician = new TechnicianDetailsDto
                            {
                                Id = reader["techCode"]?.ToString() ?? string.Empty,
                                Name = reader["name"]?.ToString() ?? string.Empty,
                                Email = reader["email"]?.ToString() ?? string.Empty,
                                Phone = reader["phone"]?.ToString() ?? string.Empty,
                                Zone = reader["zone"]?.ToString() ?? string.Empty,
                                AllJobs = reader["allJobs"] != DBNull.Value ? Convert.ToInt32(reader["allJobs"]) : 0,
                                PendingJobs = reader["pendingJobs"] != DBNull.Value ? Convert.ToInt32(reader["pendingJobs"]) : 0,
                                JobPerformancePercentage = reader["jobPerformancePercentage"] != DBNull.Value ? Convert.ToDouble(reader["jobPerformancePercentage"]) : 0.0,
                                AllServices = reader["allServices"] != DBNull.Value ? Convert.ToInt32(reader["allServices"]) : 0,
                                PendingServices = reader["pendingServices"] != DBNull.Value ? Convert.ToInt32(reader["pendingServices"]) : 0,
                                ServicePerformancePercentage = reader["servicePerformancePercentage"] != DBNull.Value ? Convert.ToDouble(reader["servicePerformancePercentage"]) : 0.0
                            };

                            result.Add(technician);
                        }
                    }
                }
            }

            return result;

        }

        public List<TechnicianDetailsByTechCodeDto> GetTechnicianDetailsByTechCode(string techCode)
        {
            var result = new List<TechnicianDetailsByTechCodeDto>();
            string query = @"
        SELECT 
            a.[TECH_CODE]                          AS techCode,
            a.[TECH_NAME]                          AS name,
            a.[EMAIL]                              AS email,
            a.[MOBILE_NO]                          AS phone,
            a.[AREA]                               AS zone,

            -- ── JOB metrics (from TBL_DAILY_JOBS) ──────────────────────────────
            ISNULL(j.ALL_JOBS,        0)           AS allJobs,
            ISNULL(j.PENDING_JOBS,    0)           AS pendingJobs,
            ISNULL(j.COMPLETE_JOBS,   0)           AS completeJobs,
            ISNULL(j.STARTED_JOBS,    0)           AS startedJobs,
            CASE 
                WHEN ISNULL(j.ALL_JOBS, 0) = 0 THEN 0
                ELSE ROUND(
                    CAST(ISNULL(j.COMPLETE_JOBS, 0) AS FLOAT) 
                    / CAST(j.ALL_JOBS AS FLOAT) * 100, 2)
            END                                    AS jobPerformancePercentage,

            -- ── SERVICE metrics (from TBL_SERVICE_SCEDULE_UPDATE) ───────────────
            ISNULL(s.ALL_SERVICES,      0)         AS allServices,
            ISNULL(s.COMPLETE_SERVICES, 0)         AS completeServices,
            ISNULL(s.PENDING_SERVICES,  0)         AS pendingServices,
            ISNULL(s.STARTED_SERVICES,  0)         AS startedServices,
            CASE 
                WHEN ISNULL(s.ALL_SERVICES, 0) = 0 THEN 0
                ELSE ROUND(
                    CAST(ISNULL(s.COMPLETE_SERVICES, 0) AS FLOAT) 
                    / CAST(s.ALL_SERVICES AS FLOAT) * 100, 2)
            END                                    AS servicePerformancePercentage

        FROM MTBL_TECH_OFFICERS a

        LEFT JOIN (
            SELECT
                [TECH_CODE],
                SUM([JOB_COUNT]) AS ALL_JOBS,
                SUM(CASE WHEN [JOB_STATUS] = 'TECH ALLOCATED'              THEN [JOB_COUNT] ELSE 0 END) AS PENDING_JOBS,
                SUM(CASE WHEN [JOB_STATUS] IN ('COMPLETE','COMPLETED')     THEN [JOB_COUNT] ELSE 0 END) AS COMPLETE_JOBS,
                SUM(CASE WHEN [JOB_STATUS] = 'started'                     THEN [JOB_COUNT] ELSE 0 END) AS STARTED_JOBS
            FROM (
                SELECT [TECH_CODE], [JOB_STATUS], COUNT(*) AS [JOB_COUNT]
                FROM TBL_DAILY_JOBS
                GROUP BY [TECH_CODE], [JOB_STATUS]
            ) raw_jobs
            GROUP BY [TECH_CODE]
        ) j ON a.[TECH_CODE] = j.[TECH_CODE]

        LEFT JOIN (
            SELECT
                [TECH_CODE],
                COUNT(DISTINCT [T_ID]) AS ALL_SERVICES,
                SUM(CASE WHEN [SV1_STATUS] = 'COMPLETED' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV2_STATUS] = 'COMPLETED' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV3_STATUS] = 'COMPLETED' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV4_STATUS] = 'COMPLETED' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV5_STATUS] = 'COMPLETED' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV6_STATUS] = 'COMPLETED' THEN 1 ELSE 0 END) AS COMPLETE_SERVICES,
                SUM(CASE WHEN [SV1_STATUS] = 'pending' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV2_STATUS] = 'pending' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV3_STATUS] = 'pending' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV4_STATUS] = 'pending' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV5_STATUS] = 'pending' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV6_STATUS] = 'pending' THEN 1 ELSE 0 END) AS PENDING_SERVICES,
                SUM(CASE WHEN [SV1_STATUS] = 'started' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV2_STATUS] = 'started' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV3_STATUS] = 'started' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV4_STATUS] = 'started' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV5_STATUS] = 'started' THEN 1 ELSE 0 END)
              + SUM(CASE WHEN [SV6_STATUS] = 'started' THEN 1 ELSE 0 END) AS STARTED_SERVICES
            FROM TBL_SERVICE_SCEDULE_UPDATE
            GROUP BY [TECH_CODE]
        ) s ON a.[TECH_CODE] = s.[TECH_CODE]

        WHERE a.[IS_ACTIVE] = '1'
          AND a.[TECH_CODE]  = @techCode";

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    result = connection.Query<TechnicianDetailsByTechCodeDto>(
                        query,
                        new { techCode }   // Dapper safely binds the parameter
                    ).ToList();
                }
            }
            catch (SqlException ex)
            {
                // Log SQL-specific errors
                Console.WriteLine($"[SQL ERROR] GetTechnicianDetailsByTechCode: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                // Log general errors
                Console.WriteLine($"[ERROR] GetTechnicianDetailsByTechCode: {ex.Message}");
                throw;
            }

            return result;
        }

        public List<JobDto> GetTechnitianPendingJobs(string techCode)
        {
            var result = new List<JobDto>();
            string query = @"
        SELECT [COM_ID]
              ,[DJ_ID]
              ,[DJ_DATE]
              ,[SERIAL_NO]
              ,[MACHINE_REF_NO]
              ,[TEAM_ID]
              ,[TEAM_NAME]
              ,[NOTE]
              ,[INFORM_BY]
              ,[CONTACT_TYPE]
              ,[CUS_ADD1]
              ,[CUS_ADD2]
              ,[CUS_ADD3]
              ,[CUS_CONTACT]
              ,[CUS_TEL_NO]
              ,[CUS_SMS_NO]
              ,[MACHINE_MODEL_ID]
              ,[MACHINE_MODEL_NAME]
              ,[CUS_STATUS]
              ,[CUS_TYPE]
              ,[JOB_STATUS]
              ,[CUS_ID]
              ,[CUS_NAME]
              ,[TECH_CODE]
              ,[TECH_MOBILE]
              ,[CR_BY]
              ,[CR_DATE]
              ,[COMPLETE_BY]
              ,[COMPLETED_DATE]
              ,[CANCELLED_BY]
              ,[CANCELLED_DATE]
              ,[SMS_DEL_REPORT_ID]
              ,[SMS_DELIVERED]
              ,[SMS_DELIVERED_DATETIME]
              ,[IS_TECH_NOTIFIED]
              ,[STARTED_BY]
              ,[STARTED_DATE]
              ,[CANCELLED_REASON]
              ,[SOLUTION_CATEGORY]
              ,[COMPLETE_SOLUTION]
              ,[RECALL_ID]
              ,[START_WORK_NOTE]
          FROM TBL_DAILY_JOBS where TECH_CODE = @techCode AND JOB_STATUS = 'TECH ALLOCATED'
        ";

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    result = connection.Query<JobDto>(
                        query,
                        new { techCode }   // Dapper safely binds the parameter
                    ).ToList();
                }
            }
            catch (SqlException ex)
            {
                // Log SQL-specific errors
                Console.WriteLine($"[SQL ERROR] GetTechnitianPendingJobs: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                // Log general errors
                Console.WriteLine($"[ERROR] GetTechnitianPendingJobs: {ex.Message}");
                throw;
            }

            return result;
        }

        public List<ServiceScheduleDto> GetTechnicianPendingServices(string techCode, string firstday, string lastday)
        {
            var result = new List<ServiceScheduleDto>();
            DateTime fromDate = DateTime.Parse(firstday);
            DateTime toDate = DateTime.Parse(lastday);


            string query = @"
                SELECT 
                T_ID,
                COM_ID,
                SERIAL_NO,
                SERVICE_DATE,
                TECH_CODE,
                TECH_NAME,
                MACHINE_REF,
                CUS_ID,
                CUS_NAME,

                SV1_STATUS,
                SV2_STATUS,
                SV3_STATUS,
                SV4_STATUS,
                SV5_STATUS,
                SV6_STATUS,

                EXPT_SV1,
                EXPT_SV2,
                EXPT_SV3,
                EXPT_SV4,
                EXPT_SV5,
                EXPT_SV6

            FROM TBL_SERVICE_SCEDULE_UPDATE

            WHERE TECH_CODE = @techCode
            AND
            (
                (
                    (SV1_STATUS IS NULL OR SV1_STATUS = '')
                    AND EXPT_SV1 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    (SV2_STATUS IS NULL OR SV2_STATUS = '')
                    AND EXPT_SV2 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    (SV3_STATUS IS NULL OR SV3_STATUS = '')
                    AND EXPT_SV3 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    (SV4_STATUS IS NULL OR SV4_STATUS = '')
                    AND EXPT_SV4 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    (SV5_STATUS IS NULL OR SV5_STATUS = '')
                    AND EXPT_SV5 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    (SV6_STATUS IS NULL OR SV6_STATUS = '')
                    AND EXPT_SV6 BETWEEN @fromDate AND @toDate
                )
            )

            ORDER BY SERVICE_DATE DESC;
        ";

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    result = connection.Query<ServiceScheduleDto>(
                        query,
                        new { techCode, fromDate, toDate }
                    ).ToList();
                }
            }
            catch (SqlException ex)
            {
                // Log SQL-specific errors
                Console.WriteLine($"[SQL ERROR] Pending Services: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                // Log general errors
                Console.WriteLine($"[ERROR] Pending Services: {ex.Message}");
                throw;
            }

            return result;
        }

        public List<JobDto> GetTechnitianAllJobs(string techCode)
        {
            var result = new List<JobDto>();
            string query = @"
        SELECT [COM_ID]
              ,[DJ_ID]
              ,[DJ_DATE]
              ,[SERIAL_NO]
              ,[MACHINE_REF_NO]
              ,[TEAM_ID]
              ,[TEAM_NAME]
              ,[NOTE]
              ,[INFORM_BY]
              ,[CONTACT_TYPE]
              ,[CUS_ADD1]
              ,[CUS_ADD2]
              ,[CUS_ADD3]
              ,[CUS_CONTACT]
              ,[CUS_TEL_NO]
              ,[CUS_SMS_NO]
              ,[MACHINE_MODEL_ID]
              ,[MACHINE_MODEL_NAME]
              ,[CUS_STATUS]
              ,[CUS_TYPE]
              ,[JOB_STATUS]
              ,[CUS_ID]
              ,[CUS_NAME]
              ,[TECH_CODE]
              ,[TECH_MOBILE]
              ,[CR_BY]
              ,[CR_DATE]
              ,[COMPLETE_BY]
              ,[COMPLETED_DATE]
              ,[CANCELLED_BY]
              ,[CANCELLED_DATE]
              ,[SMS_DEL_REPORT_ID]
              ,[SMS_DELIVERED]
              ,[SMS_DELIVERED_DATETIME]
              ,[IS_TECH_NOTIFIED]
              ,[STARTED_BY]
              ,[STARTED_DATE]
              ,[CANCELLED_REASON]
              ,[SOLUTION_CATEGORY]
              ,[COMPLETE_SOLUTION]
              ,[RECALL_ID]
              ,[START_WORK_NOTE]
          FROM TBL_DAILY_JOBS where TECH_CODE = @techCode
        ";

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    result = connection.Query<JobDto>(
                        query,
                        new { techCode }   // Dapper safely binds the parameter
                    ).ToList();
                }
            }
            catch (SqlException ex)
            {
                // Log SQL-specific errors
                Console.WriteLine($"[SQL ERROR] GetTechnitianPendingJobs: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                // Log general errors
                Console.WriteLine($"[ERROR] GetTechnitianPendingJobs: {ex.Message}");
                throw;
            }

            return result;
        }

        public List<ServiceScheduleDto> GetTechnicianAllServices(string techCode, string firstday, string lastday)
        {
            var result = new List<ServiceScheduleDto>();
            DateTime fromDate = DateTime.Parse(firstday);
            DateTime toDate = DateTime.Parse(lastday);


            string query = @"
                SELECT 
                T_ID,
                COM_ID,
                SERIAL_NO,
                SERVICE_DATE,
                TECH_CODE,
                TECH_NAME,
                MACHINE_REF,
                CUS_ID,
                CUS_NAME,

                SV1_STATUS,
                SV2_STATUS,
                SV3_STATUS,
                SV4_STATUS,
                SV5_STATUS,
                SV6_STATUS,

                EXPT_SV1,
                EXPT_SV2,
                EXPT_SV3,
                EXPT_SV4,
                EXPT_SV5,
                EXPT_SV6

            FROM TBL_SERVICE_SCEDULE_UPDATE


            WHERE TECH_CODE = @techCode
            AND
            (
                (
                    EXPT_SV1 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    EXPT_SV2 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    EXPT_SV3 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    EXPT_SV4 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    EXPT_SV5 BETWEEN @fromDate AND @toDate
                )
                OR
                (
                    EXPT_SV6 BETWEEN @fromDate AND @toDate
                )
            )

            ORDER BY SERVICE_DATE DESC;
        ";

            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    result = connection.Query<ServiceScheduleDto>(
                        query,
                        new { techCode, fromDate, toDate }
                    ).ToList();
                }
            }
            catch (SqlException ex)
            {
                // Log SQL-specific errors
                Console.WriteLine($"[SQL ERROR] Pending Services: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                // Log general errors
                Console.WriteLine($"[ERROR] Pending Services: {ex.Message}");
                throw;
            }

            return result;

        }
    }
}
