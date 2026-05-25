namespace ServvistaWebAppAPI.Models.Dtos
{
    public class OldestDueDto
    {
        public string DjId { get; set; }
        public string SerialNo { get; set; }
        public string MachineRefNo { get; set; }
        public string CustomerName { get; set; }
        public string CustomerAddress { get; set; }  // Combined address
        public string CustomerContact { get; set; }
        public string CustomerTelNo { get; set; }
        public string TeamId { get; set; }
        public string TeamName { get; set; }
        public string DjDate { get; set; }
        public string TechCode { get; set; }
        public string TechMobile { get; set; }
        public string MachineModelId { get; set; }
        public string MachineModelName { get; set; }
        public string CustomerStatus { get; set; }
        public string JobStatus { get; set; }
        public string Note { get; set; }
        public string DaysOverdue { get; set; }  // Calculated field
    }
}