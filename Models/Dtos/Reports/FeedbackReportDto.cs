namespace ServvistaWebAppAPI.Models.Dtos.Reports
{
    public class FeedbackReportDto
    {
        public int TID { get; set; }
        public string? COM_ID { get; set; }
        public DateTime FB_DATE { get; set; }
        public string? MOBILE_NO { get; set; }
        public string? CUS_CODE { get; set; }
        public string? CUS_NAME { get; set; }
        public int RATING { get; set; }
        public string? FULL_MSG { get; set; }
        public int? JOB_ID { get; set; }
        public string? SERIAL_NO { get; set; }
        public string? MACHINE_REF_NO { get; set; }
        public string? CUSTOMER_REVIEW { get; set; }
        public string? TYPE { get; set; }
        public int? VISIT_NO { get; set; }
        public string? TECH_CODE { get; set; }
    }
}