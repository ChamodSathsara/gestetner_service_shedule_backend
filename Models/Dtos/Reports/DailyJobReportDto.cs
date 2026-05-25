public class DailyJobReportDto
{
    public string? COM_ID { get; set; }      // DB returns "001" (string)
    public int DJ_ID { get; set; }
    public DateTime DJ_DATE { get; set; }
    public string? SERIAL_NO { get; set; }
    public string? MACHINE_REF_NO { get; set; }
    public string? TEAM_ID { get; set; }
    public string? TEAM_NAME { get; set; }
    public string? NOTE { get; set; }
    public string? INFORM_BY { get; set; }
    public string? CONTACT_TYPE { get; set; }
    public string? CUS_ADD1 { get; set; }
    public string? CUS_ADD2 { get; set; }
    public string? CUS_ADD3 { get; set; }
    public string? CUS_CONTACT { get; set; }
    public string? CUS_TEL_NO { get; set; }
    public string? CUS_SMS_NO { get; set; }
    public string? MACHINE_MODEL_ID { get; set; }
    public string? MACHINE_MODEL_NAME { get; set; }
    public string? CUS_STATUS { get; set; }
    public string? CUS_TYPE { get; set; }
    public string? JOB_STATUS { get; set; }
    public string? CUS_ID { get; set; }      // DB returns "BS01063", "1656" etc → string
    public string? CUS_NAME { get; set; }
    public string? TECH_CODE { get; set; }
    public string? TECH_MOBILE { get; set; }
    public string? CR_BY { get; set; }
    public DateTime? CR_DATE { get; set; }
    public string? COMPLETE_BY { get; set; }
    public DateTime? COMPLETED_DATE { get; set; }
    public string? CANCELLED_BY { get; set; }
    public DateTime? CANCELLED_DATE { get; set; }
    public int? SMS_DEL_REPORT_ID { get; set; }
    public bool? SMS_DELIVERED { get; set; }
    public DateTime? SMS_DELIVERED_DATETIME { get; set; }
    public bool? IS_TECH_NOTIFIED { get; set; }
    public string? STARTED_BY { get; set; }
    public DateTime? STARTED_DATE { get; set; }
    public string? CANCELLED_REASON { get; set; }
    public string? SOLUTION_CATEGORY { get; set; }
    public string? COMPLETE_SOLUTION { get; set; }
    public int? RECALL_ID { get; set; }
    public string? START_WORK_NOTE { get; set; }
}