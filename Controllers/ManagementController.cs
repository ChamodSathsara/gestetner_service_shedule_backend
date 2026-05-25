using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using ServvistaWebAppAPI.Repositories;
using ServvistaWebAppAPI.Services;

namespace ServvistaWebAppAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ManagementController : ControllerBase
    {
        private readonly IManagementRepository _managementService;

        public ManagementController(IManagementRepository managementRepository, ITenantService tenantService)
        {
            _managementService = managementRepository;
        }

        [Authorize]
        [HttpGet("getJobCountAndRate")]
        public IActionResult GetJobCountAndRate()
        {
            try
            {
                // Placeholder for actual logic to retrieve job count and rate
                var data = _managementService.GetJobCountAndRate();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getServiceCountAndRate")]
        public IActionResult GetServiceCountAndRate()
        {
            try
            {
                // Placeholder for actual logic to retrieve job count and rate
                var data = _managementService.GetServiceCountAndRate();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getPendingJobs")]
        public IActionResult GetPendingJobs()
        {
            try
            {
                var data = _managementService.GetPendingJobs();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getPendingServices")]
        public IActionResult GetPendingServices()
        {
            try
            {
                var data = _managementService.GetPendingServices();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getCompleteAndPendingJobPercentage")]
        public IActionResult GetCompleteAndPendingJobPercentage()
        {
            try
            {
                var data = _managementService.GetCompleteAndPendingJobPercentage();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getCompleteAndPendingServicesPercentage")]
        public IActionResult GetCompleteAndPendingServicesPercentage()
        {
            try
            {
                var data = _managementService.GetCompleteAndPendingServicesPercentage();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getLastWeekJobPerformence")]
        public IActionResult GetLastWeekJobPerformence()
        {
            try
            {
                var data = _managementService.GetLastWeekJobPerformance();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getTechniciansPerformance")]
        public IActionResult GetTechniciansPerformence()
        {
            try
            {
                var data = _managementService.GetTechniciansPerformence();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [Authorize]
        [HttpGet("getOldestDueJobs")]
        public IActionResult GetOlderstDueJobs()
        {
            try
            {
                var data = _managementService.GetOldestDueJobs();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getCustomerWarranty")]
        public IActionResult GetWarrantyDetails()
        {
            try
            {
                var data = _managementService.GetWarrantyDetails();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }


        [Authorize]
        [HttpGet("getTechnitianDetails")]
        public IActionResult GetTechnitianDetails()
        {
            try
            {
                var data = _managementService.GetTechnicianDetails();
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getTechnitianDetailsByTechCode")]
        public IActionResult GetTechnitianDetailsByTechCode(string techCode)
        {
            try
            {
                var data = _managementService.GetTechnicianDetailsByTechCode(techCode);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getTechnitianPendingJobs")]
        public IActionResult GetTechnitianPendingJobs(string techCode)
        {
            try
            {
                var data = _managementService.GetTechnitianPendingJobs(techCode);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getTechnitianPendingServices")]
        public IActionResult GetTechnitianPendingServices(string techCode , string firstday, string lastday)
        {
            try
            {
                var data = _managementService.GetTechnicianPendingServices(techCode , firstday, lastday);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getTechnitianAllJobs")]
        public IActionResult GetTechnitianAllJobs(string techCode)
        {
            try
            {
                var data = _managementService.GetTechnitianAllJobs(techCode);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }

        [Authorize]
        [HttpGet("getTechnitianAllServices")]
        public IActionResult GetTechnitianAllServices(string techCode, string firstday, string lastday)
        {
            try
            {
                var data = _managementService.GetTechnicianAllServices(techCode, firstday, lastday);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }

        }
    }
}
