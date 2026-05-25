namespace ServvistaWebAppAPI.Models.Dtos
{
    public class TechnicianDetailsDto
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Zone { get; set; }

        // Jobs
        public int AllJobs { get; set; }
        public int PendingJobs { get; set; }
        public double JobPerformancePercentage { get; set; }

        // Services
        public int AllServices { get; set; }
        public int PendingServices { get; set; }
        public double ServicePerformancePercentage { get; set; }
    }
}
