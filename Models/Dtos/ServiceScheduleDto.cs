namespace ServvistaWebAppAPI.Models.Dtos
{
    public class ServiceScheduleDto
    {
        public int T_ID { get; set; }

        public string COM_ID { get; set; }

        public string SERIAL_NO { get; set; }

        public DateTime? SERVICE_DATE { get; set; }

        public string TECH_CODE { get; set; }

        public string TECH_NAME { get; set; }

        public string MACHINE_REF { get; set; }

        public string CUS_ID { get; set; }

        public string CUS_NAME { get; set; }

        public string SV1_STATUS { get; set; }

        public string SV2_STATUS { get; set; }

        public string SV3_STATUS { get; set; }

        public string SV4_STATUS { get; set; }

        public string SV5_STATUS { get; set; }

        public string SV6_STATUS { get; set; }

        public DateTime? EXPT_SV1 { get; set; }

        public DateTime? EXPT_SV2 { get; set; }

        public DateTime? EXPT_SV3 { get; set; }

        public DateTime? EXPT_SV4 { get; set; }

        public DateTime? EXPT_SV5 { get; set; }

        public DateTime? EXPT_SV6 { get; set; }
    }
}
