namespace ServvistaWebAppAPI.Models.Dtos.Reports
{
    public class MonthlyServiceScheduleDto
    {
        public int T_ID { get; set; }
        public string? COM_ID { get; set; }
        public string? SERIAL_NO { get; set; }
        public string? Customer_Code { get; set; }
        public string? Customer_Name { get; set; }
        public string? TECH_CODE { get; set; }
        public string? TECH_NAME { get; set; }
        public string? AREA { get; set; }
        public string? MACHINE_REF { get; set; }
        public string? M_MODEL { get; set; }
        public string? agrementType { get; set; }
        public string? TEL_NO { get; set; }
        public string? M_LOC1 { get; set; }
        public string? M_LOC2 { get; set; }
        public string? M_LOC3 { get; set; }
        public DateTime? machine_instalation_or_period_start_date { get; set; }
        public int VisitNo { get; set; }
        public DateTime? ExpectedVisitDate { get; set; }
        public DateTime? ServiceDate { get; set; }
        public string? MeterReading { get; set; }
        public string? VisitStatus { get; set; }
    }
}