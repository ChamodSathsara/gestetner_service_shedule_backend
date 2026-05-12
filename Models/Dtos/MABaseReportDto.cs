using System.Security.Principal;

namespace ServvistaWebAppAPI.Models.Dtos
{
    public class MABaseReportDto
    {
        public string CustomerName { get; set; }
        public string CusAddress01 { get; set; }
        public string CusAddress02 {  get; set; }   
        public string CusAddress03 { get; set; }
        public string Model { get; set; }
        public string Serial { get; set; }
        public string MachineRefNo { get; set; }
        public DateTime InstallationDate { get; set; }
        public string CustomerGrade { get; set; }
        public string Team { get; set; }
        public string CustomerStatus { get; set; }
        public DateTime maPeriodStart { get; set; }
        public DateTime maPeriodEnd { get; set; }
        public string TechOfficer { get; set; } //Main Technician
        public string TechName { get; set; } //Main Technician Name
        public string VC { get; set; }
        public string Type { get; set; }
        public string MachineInvoiceNumber { get; set; }
    }
}
