using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ServvistaWebAppAPI.Repositories;

namespace ServvistaWebAppAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        private readonly IReportRepository _reportService;

        public ReportController(IReportRepository reportRepository)
        {
            _reportService = reportRepository;
        }

        // ──────────────────────────────────────────────────────────────────────
        // GET DATA  →  JSON
        // GET /api/report/getNSReportData?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("getNSReportData")]
        public IActionResult GetNSReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetNsReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // EXCEL DOWNLOAD
        // GET /api/report/downloadNSReportExcel?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadNSReportExcel")]
        public IActionResult DownloadNSReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadNsReportExcel(startDate, lastDate);
                var fileName = $"NS_Report_{startDate}_to_{lastDate}.xlsx";

                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // PDF DOWNLOAD
        // GET /api/report/downloadNSReportPdf?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadNSReportPdf")]
        public IActionResult DownloadNSReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadNsReportPdf(startDate, lastDate);
                var fileName = $"NS_Report_{startDate}_to_{lastDate}.pdf";

                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }


        // ──────────────────────────────────────────────────────────────────────
        // FS REPORT: GET DATA  →  JSON
        // GET /api/report/getFSReportData?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("getFSReportData")]
        public IActionResult GetFSReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetFsReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // FS REPORT: EXCEL DOWNLOAD
        // GET /api/report/downloadFSReportExcel?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadFSReportExcel")]
        public IActionResult DownloadFSReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadFsReportExcel(startDate, lastDate);
                var fileName = $"FS_Report_{startDate}_to_{lastDate}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // FS REPORT: PDF DOWNLOAD
        // GET /api/report/downloadFSReportPdf?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadFSReportPdf")]
        public IActionResult DownloadFSReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadFsReportPdf(startDate, lastDate);
                var fileName = $"FS_Report_{startDate}_to_{lastDate}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }


        // ──────────────────────────────────────────────────────────────────────
        // MA REPORT: GET DATA  →  JSON
        // GET /api/report/getMAReportData?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("getMAReportData")]
        public IActionResult GetMAReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetMaReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // MA REPORT: EXCEL DOWNLOAD
        // GET /api/report/downloadMAReportExcel?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadMAReportExcel")]
        public IActionResult DownloadMAReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadMaReportExcel(startDate, lastDate);
                var fileName = $"MA_Report_{startDate}_to_{lastDate}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // MA REPORT: PDF DOWNLOAD
        // GET /api/report/downloadMAReportPdf?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadMAReportPdf")]
        public IActionResult DownloadMAReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadMaReportPdf(startDate, lastDate);
                var fileName = $"MA_Report_{startDate}_to_{lastDate}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // EX REPORT: GET DATA  →  JSON
        // GET /api/report/getExReportData?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("getExReportData")]
        public IActionResult GetExReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetExReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // EX REPORT: EXCEL DOWNLOAD
        // GET /api/report/downloadExReportExcel?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadExReportExcel")]
        public IActionResult DownloadExReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadExReportExcel(startDate, lastDate);
                var fileName = $"EX_Report_{startDate}_to_{lastDate}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // EX REPORT: PDF DOWNLOAD
        // GET /api/report/downloadExReportPdf?startDate=2026-01-01&lastDate=2026-12-31
        // ──────────────────────────────────────────────────────────────────────
        [Authorize]
        [HttpGet("downloadExReportPdf")]
        public IActionResult DownloadExReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadExReportPdf(startDate, lastDate);
                var fileName = $"EX_Report_{startDate}_to_{lastDate}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }


        [Authorize]
        [HttpGet("getDailyJobReportData")]
        public IActionResult GetDailyJobReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetDailyJobReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadDailyJobReportExcel")]
        public IActionResult DownloadDailyJobReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadDailyJobReportExcel(startDate, lastDate);
                var fileName = $"DailyJobReport_{startDate}_to_{lastDate}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadDailyReportReportPdf")]
        public IActionResult DownloadDailyReportReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadDailyJobReportPdf(startDate, lastDate);
                var fileName = $"DailyJobReport_{startDate}_to_{lastDate}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }


        [Authorize]
        [HttpGet("getCustomerReportData")]
        public IActionResult GetCustomerReportData()
        {
            try
            {
                var data = _reportService.GetCustomerReportData();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadCustomerReportExcel")]
        public IActionResult DownloadCustomerReportExcel()
        {
            try
            {
                var bytes = _reportService.DownloadCustomerReportExcel();
                var fileName = $"CustomerReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadCustomerReportPdf")]
        public IActionResult DownloadCustomerReportPdf()
        {
            try
            {
                var bytes = _reportService.DownloadCustomerReportPdf();
                var fileName = $"CustomerReport_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getCustomerMachineData")]
        public IActionResult GetCustomerMachineReportData(string customerCode)
        {
            try
            {
                var data = _reportService.GetMachineList(customerCode);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }


        [Authorize]
        [HttpGet("getServiceVisitReportData")]
        public IActionResult GetServiceVisitReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetMonthlyServiceReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadServiceVisitReportExcel")]
        public IActionResult DownloadServiceVisitReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadMonthlyServiceReportExcel(startDate, lastDate);
                var fileName = $"service visit Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadServiceVisitReportPdf")]
        public IActionResult DownloadServiceVisitReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadMonthlyServiceReportPdf(startDate, lastDate);
                var fileName = $"Service visit Report_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getFeedbackReportData")]
        public IActionResult GetFeedbackReportData(string startDate, string lastDate)
        {
            try
            {
                var data = _reportService.GetFeedbackReportData(startDate, lastDate);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadFeedbackReportExcel")]
        public IActionResult DownloadFeedbackReportExcel(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadFeedbackReportExcel(startDate, lastDate);
                var fileName = $"feedback Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("downloadFeedbackReportPdf")]
        public IActionResult DownloadFeedbackReportPdf(string startDate, string lastDate)
        {
            try
            {
                var bytes = _reportService.DownloadFeedbackReportPdf(startDate, lastDate);
                var fileName = $"feedback Report_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                return File(bytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

    }
}