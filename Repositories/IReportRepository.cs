using ServvistaWebAppAPI.Models.Dtos.Reports;

namespace ServvistaWebAppAPI.Repositories
{
    public interface IReportRepository
    {
        // NS Report methods
        List<NsReportDto> GetNsReportData(string startDate, string lastDate);
        byte[] DownloadNsReportExcel(string startDate, string lastDate);
        byte[] DownloadNsReportPdf(string startDate, string lastDate);

        // FS Report methods
        List<FsReportDto> GetFsReportData(string startDate, string lastDate);
        byte[] DownloadFsReportExcel(string startDate, string lastDate);
        byte[] DownloadFsReportPdf(string startDate, string lastDate);

        // MA Report methods
        List<MaReportDto> GetMaReportData(string startDate, string lastDate);
        byte[] DownloadMaReportExcel(string startDate, string lastDate);
        byte[] DownloadMaReportPdf(string startDate, string lastDate);

        // EX Report methods
        List<ExReportDto> GetExReportData(string startDate, string lastDate);
        byte[] DownloadExReportExcel(string startDate, string lastDate);
        byte[] DownloadExReportPdf(string startDate, string lastDate);

        //DailyJob Report
        List<DailyJobReportDto> GetDailyJobReportData(string startDate, string lastDate);
        byte[] DownloadDailyJobReportExcel(string startDate, string lastDate);
        byte[] DownloadDailyJobReportPdf(string startDate, string lastDate);

        //customer Base
        public List<CustomerReportDto> GetCustomerReportData();
        public byte[] DownloadCustomerReportExcel();
        public byte[] DownloadCustomerReportPdf();
        //customer Machine List
        public List<FsReportDto> GetMachineList(string customerCode);

        //Monthly Service Visits
        public List<MonthlyServiceReportDto> GetMonthlyServiceReportData(string startDate, string lastDate);
        public byte[] DownloadMonthlyServiceReportExcel(string startDate, string lastDate);
        public byte[] DownloadMonthlyServiceReportPdf(string startDate, string lastDate);

        //Feedback 
        public List<FeedbackReportDto> GetFeedbackReportData(string startDate, string lastDate);
        public byte[] DownloadFeedbackReportExcel(string startDate, string lastDate);
        public byte[] DownloadFeedbackReportPdf(string startDate, string lastDate);

        //Monthly Service Schedule
        public List<MonthlyServiceScheduleDto> GetMonthlyServiceScheduleData(string startDate, string lastDate, string techCode , string machineRefStartAt);
        public byte[] DownloadMonthlyServiceSchedulePdf(string startDate, string lastDate, string techCode , string machineRefStartAt);
    }
}