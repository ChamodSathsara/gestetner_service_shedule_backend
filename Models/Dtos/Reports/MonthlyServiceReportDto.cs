namespace ServvistaWebAppAPI.Models.Dtos.Reports
{
    public class MonthlyServiceReportDto
    {
        public int T_ID { get; set; }
        public string? COM_ID { get; set; }
        public string? SERIAL_NO { get; set; }
        public string? AREA { get; set; }
        public string? CUS_ID { get; set; }
        public string? CUS_NAME { get; set; }
        public string? TECH_CODE { get; set; }
        public string? TECH_NAME { get; set; }
        public string? MACHINE_REF { get; set; }
        public string? M_MODEL { get; set; }
        public int VisitNo { get; set; }
        public DateTime? ExpectedVisitDate { get; set; }
        public DateTime? ServiceDate { get; set; }
        public string? MeterReading { get; set; }
        public string? VisitStatus { get; set; }
    }
}