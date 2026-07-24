using System.ComponentModel.DataAnnotations;

namespace ServvistaWebAppAPI.Models.Dtos
{
    public class RegisterPushDeviceRequest
    {
        [Required]
        public string TechCode { get; set; } = string.Empty;

        public string? CompanyId { get; set; }

        [Required]
        public string FcmToken { get; set; } = string.Empty;

        public string? DeviceName { get; set; }

        public string DeviceType { get; set; } = "PWA";
    }
}