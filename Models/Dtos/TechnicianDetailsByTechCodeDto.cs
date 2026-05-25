namespace ServvistaWebAppAPI.Models.Dtos
{
    public class TechnicianDetailsByTechCodeDto
    {
        // ── Technician Info ─────────────────────────────
        public string TechCode { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Zone { get; set; }

        // ── Job Metrics ─────────────────────────────────
        public int AllJobs { get; set; }
        public int PendingJobs { get; set; }
        public int CompleteJobs { get; set; }
        public int StartedJobs { get; set; }
        public double JobPerformancePercentage { get; set; }

        // ── Service Metrics ──────────────────────────────
        public int AllServices { get; set; }
        public int CompleteServices { get; set; }
        public int PendingServices { get; set; }
        public int StartedServices { get; set; }
        public double ServicePerformancePercentage { get; set; }
    }
}
