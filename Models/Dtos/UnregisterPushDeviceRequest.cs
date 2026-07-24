using System.ComponentModel.DataAnnotations;

namespace ServvistaWebAppAPI.Models.Dtos
{
    public class UnregisterPushDeviceRequest
    {
        [Required]
        public string FcmToken { get; set; } = string.Empty;
    }
}