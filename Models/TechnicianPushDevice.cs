namespace ServvistaWebAppAPI.Models
{
    public class TechnicianPushDevice
    {
        public long ID { get; set; }

        public string TECH_CODE { get; set; } = string.Empty;

        public string? COMPANY_ID { get; set; }

        public string FCM_TOKEN { get; set; } = string.Empty;

        public string? DEVICE_NAME { get; set; }

        public string? DEVICE_TYPE { get; set; }

        public bool IS_ACTIVE { get; set; }

        public DateTime CREATED_AT { get; set; }

        public DateTime? UPDATED_AT { get; set; }

        public DateTime? LAST_USED_AT { get; set; }
    }
}