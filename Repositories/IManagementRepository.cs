using ServvistaWebAppAPI.Models;
using ServvistaWebAppAPI.Models.Dtos;
namespace ServvistaWebAppAPI.Repositories
{
    public interface IManagementRepository
    {
        JobCountAndRateDto GetJobCountAndRate();
        JobCountAndRateDto GetServiceCountAndRate();
        List<BreakdownModel> GetPendingJobs();
        List<ServiceVisitMonthlyInfo> GetPendingServices();
        List<CompleteAndPendingPercentageDto> GetCompleteAndPendingJobPercentage();
        List<CompleteAndPendingPercentageDto> GetCompleteAndPendingServicesPercentage();

        List<LastYearJobPerformanceDto> GetLastWeekJobPerformance();
        List<TechnicianPerformenceDto> GetTechniciansPerformence();

        List<OldestDueDto> GetOldestDueJobs();
        List<WarrantyDetailsDto> GetWarrantyDetails();

        List<TechnicianDetailsDto> GetTechnicianDetails();
        List<TechnicianDetailsByTechCodeDto> GetTechnicianDetailsByTechCode(string techCode);

        List<JobDto> GetTechnitianPendingJobs(string techCode);
        List<ServiceScheduleDto> GetTechnicianPendingServices(string techCode, string firstday, string lastday);
        List<JobDto> GetTechnitianAllJobs(string techCode);

        List<ServiceScheduleDto> GetTechnicianAllServices(string techCode, string firstday, string lastday);
    }
}
