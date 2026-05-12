namespace ServvistaWebAppAPI.Models
{
    public class AppUsers
    {
        public int COM_ID { get; set; }

        public string USER_CODE { get; set; } = string.Empty;

        public string USER_NAME { get; set; } = string.Empty;

        public string? MOBILE_NO { get; set; }

        public string? EMAIL { get; set; }

        public string? AREA { get; set; }

        public string? CITY { get; set; }

        public bool IS_ACTIVE { get; set; }

        public string? PASSWORD_HASH { get; set; }

        public string? REFRESH_TOKEN { get; set; }

        public string USER_ROLE { get; set; }
    }
}