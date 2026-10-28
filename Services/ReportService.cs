using ClosedXML.Excel;
using Dapper;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using Org.BouncyCastle.Utilities.Zlib;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using ServvistaWebAppAPI.Models.Dtos.Reports;
using ServvistaWebAppAPI.Repositories;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

namespace ServvistaWebAppAPI.Services
{
    public class ReportService : IReportRepository
    {
        private readonly string _connectionString;
        private readonly ITenantService _tenantService;

        public ReportService(IConfiguration config, ITenantService tenantService)
        {
            _tenantService = tenantService;
            _connectionString = _tenantService.GetConnectionString();

            // Required once per app lifetime — Community license is free
            QuestPDF.Settings.License = LicenseType.Community;
        }

        // ──────────────────────────────────────────────────────────────────────
        // SHARED: fetch data from DB
        // ──────────────────────────────────────────────────────────────────────
        public List<NsReportDto> GetNsReportData(string startDate, string lastDate)
        {
            DateTime fromDate = DateTime.Parse(startDate);
            DateTime toDate = DateTime.Parse(lastDate);

            const string query = @"
                SELECT
                     [COM_ID]               AS ComId
                    ,[SERIAL_NO]            AS SerialNo
                    ,[DIVISION_ID]          AS DivisionId
                    ,[AGREEMENT_ID]         AS AgreementId
                    ,[CUS_CODE]             AS CusCode
                    ,[CUS_NAME]             AS CusName
                    ,[INV_ADD1]             AS InvAdd1
                    ,[INV_ADD2]             AS InvAdd2
                    ,[INV_ADD3]             AS InvAdd3
                    ,[TEL_NO]               AS TelNo
                    ,[FAX_NO]               AS FaxNo
                    ,[CUS_GRADE]            AS CusGrade
                    ,[MACHINE_REF_CODE]     AS MachineRefCode
                    ,[MACHINE_CODE]         AS MachineCode
                    ,[MACHINE_DESC]         AS MachineDesc
                    ,[M_LOC1]               AS MLoc1
                    ,[M_LOC2]               AS MLoc2
                    ,[M_LOC3]               AS MLoc3
                    ,[MACHINE_MODE]         AS MachineMode
                    ,[M_LOC_CONTACT_NAME]   AS MLocContactName
                    ,[M_LOC_CONTACT_NO]     AS MLocContactNo
                    ,[MFP_CITY]             AS MfpCity
                    ,[T_OFFICER_CODE]       AS TOfficerCode
                    ,[T_OFFICER_NAME]       AS TOfficerName
                    ,[M_INV_NO]             AS MInvNo
                    ,[M_INV_DATE]           AS MInvDate
                    ,[M_INS_DATE]           AS MInsDate
                    ,[M_WARRANTY_END_DATE]  AS MWarrantyEndDate
                    ,[LABOR_WARRANTY]       AS LaborWarranty
                    ,[PARTS_WARRANTY]       AS PartsWarranty
                    ,[IS_TRANSPORT]         AS IsTransport
                    ,[PROVISION]            AS Provision
                    ,[CUS_STATUS]           AS CusStatus
                    ,[VISITS_PER_YEAR]      AS VisitsPerYear
                    ,[INV_NO_TC]            AS InvNoTc
                    ,[INV_DATE_TC]          AS InvDateTc
                    ,[TEAM]                 AS Team
                    ,[MA_PERIOD_START]      AS MaPeriodStart
                    ,[MA_PERIOD_END]        AS MaPeriodEnd
                    ,[SPECIAL]              AS Special
                    ,[DATA_CHECKED]         AS DataChecked
                    ,[REP_CODE]             AS RepCode
                    ,[REP_NAME]             AS RepName
                    ,[TO_CODE]              AS ToCode
                    ,[TO_NAME]              AS ToName
                    ,[TO_AREA_CODE]         AS ToAreaCode
                    ,[TO_AREA_NAME]         AS ToAreaName
                    ,[DEALER_ID]            AS DealerId
                    ,[DEALER_NAME]          AS DealerName
                    ,[CONS_REP_CODE]        AS ConsRepCode
                    ,[CONS_REP_NAME]        AS ConsRepName
                    ,[SERVICE_DATE]         AS ServiceDate
                    ,[VISIT_CODE]           AS VisitCode
                    ,[WARRANTY_TYPE_ID]     AS WarrantyTypeId
                    ,[WARRANTY_TYPE_DESC]   AS WarrantyTypeDesc
                    ,[MACHINE_TYPE]         AS MachineType
                    ,[AD_REMORT]            AS AdRemort
                    ,[CUS_TYPE]             AS CusType
                    ,[SAGE_CODE]            AS SageCode
                    ,[SUB_CUS_TYPE]         AS SubCusType
                    ,[CONS_COMMENT]         AS ConsComment
                    ,[CUS_EMAIL]            AS CusEmail
                    ,[WARRANTY_YEARS]       AS WarrantyYears
                    ,[CN_NO]                AS CnNo
                    ,[M_LOC_TELEPHONE]      AS MLocTelephone
                FROM TBL_MACHINE_TRANSACTION
                WHERE
                    [CUS_STATUS] = 'NS'
                    AND [MACHINE_TYPE] in ('MFP','BCP','MPC' )
                    AND (
                        (
                            [MA_PERIOD_END] IS NOT NULL
                            AND [MA_PERIOD_END] > @fromDate
                            AND [MA_PERIOD_END] < @toDate
                        )
                        OR
                        (
                            [MA_PERIOD_END] IS NULL
                            AND [M_WARRANTY_END_DATE] > @fromDate
                            AND [M_WARRANTY_END_DATE] < @toDate
                        )
                    )
                ORDER BY [CUS_CODE]";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<NsReportDto>(query, new { fromDate, toDate }).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetNsReportData: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetNsReportData: {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // EXCEL  (NuGet: ClosedXML)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadNsReportExcel(string startDate, string lastDate)
        {
            var data = GetNsReportData(startDate, lastDate);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("NS Report");

            // ── Header row ──────────────────────────────────────────────────
            var headers = new[]
            {
                "COM ID","Serial No","Division","Agreement","Cus Code","Customer Name",
                "Inv Add1","Inv Add2","Inv Add3","Tel No","Fax No","Grade",
                "Machine Ref","Machine Code","Machine Desc",
                "M Loc1","M Loc2","M Loc3","Machine Mode",
                "Loc Contact","Loc Contact No","City",
                "Officer Code","Officer Name",
                "Inv No","Inv Date","Ins Date","Warranty End",
                "Labor Warranty","Parts Warranty","Transport","Provision",
                "Status","Visits/Yr","Inv No TC","Inv Date TC",
                "Team","MA Start","MA End","Special","Data Checked",
                "Rep Code","Rep Name","TO Code","TO Name","Area Code","Area Name",
                "Dealer ID","Dealer Name","Cons Rep Code","Cons Rep Name",
                "Service Date","Visit Code","Warranty Type ID","Warranty Type",
                "Machine Type","Ad Remort","Cus Type","Sage Code","Sub Cus Type",
                "Comment","Email","Warranty Yrs","CN No","Loc Telephone"
            };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192); // Dark blue
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // ── Data rows ───────────────────────────────────────────────────
            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.ComId ?? "";
                ws.Cell(row, c++).Value = r.SerialNo ?? "";
                ws.Cell(row, c++).Value = r.DivisionId ?? "";
                ws.Cell(row, c++).Value = r.AgreementId ?? "";
                ws.Cell(row, c++).Value = r.CusCode ?? "";
                ws.Cell(row, c++).Value = r.CusName ?? "";
                ws.Cell(row, c++).Value = r.InvAdd1 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd2 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd3 ?? "";
                ws.Cell(row, c++).Value = r.TelNo ?? "";
                ws.Cell(row, c++).Value = r.FaxNo ?? "";
                ws.Cell(row, c++).Value = r.CusGrade ?? "";
                ws.Cell(row, c++).Value = r.MachineRefCode ?? "";
                ws.Cell(row, c++).Value = r.MachineCode ?? "";
                ws.Cell(row, c++).Value = r.MachineDesc ?? "";
                ws.Cell(row, c++).Value = r.MLoc1 ?? "";
                ws.Cell(row, c++).Value = r.MLoc2 ?? "";
                ws.Cell(row, c++).Value = r.MLoc3 ?? "";
                ws.Cell(row, c++).Value = r.MachineMode ?? "";
                ws.Cell(row, c++).Value = r.MLocContactName ?? "";
                ws.Cell(row, c++).Value = r.MLocContactNo ?? "";
                ws.Cell(row, c++).Value = r.MfpCity ?? "";
                ws.Cell(row, c++).Value = r.TOfficerCode ?? "";
                ws.Cell(row, c++).Value = r.TOfficerName ?? "";
                ws.Cell(row, c++).Value = r.MInvNo ?? "";
                ws.Cell(row, c++).Value = r.MInvDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MInsDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MWarrantyEndDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.LaborWarranty ?? "";
                ws.Cell(row, c++).Value = r.PartsWarranty ?? "";
                ws.Cell(row, c++).Value = r.IsTransport ?? "";
                ws.Cell(row, c++).Value = r.Provision ?? "";
                ws.Cell(row, c++).Value = r.CusStatus ?? "";
                ws.Cell(row, c++).Value = r.VisitsPerYear?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.InvNoTc ?? "";
                ws.Cell(row, c++).Value = r.InvDateTc?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Team ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodStart?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodEnd?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Special ?? "";
                ws.Cell(row, c++).Value = r.DataChecked ?? "";
                ws.Cell(row, c++).Value = r.RepCode ?? "";
                ws.Cell(row, c++).Value = r.RepName ?? "";
                ws.Cell(row, c++).Value = r.ToCode ?? "";
                ws.Cell(row, c++).Value = r.ToName ?? "";
                ws.Cell(row, c++).Value = r.ToAreaCode ?? "";
                ws.Cell(row, c++).Value = r.ToAreaName ?? "";
                ws.Cell(row, c++).Value = r.DealerId ?? "";
                ws.Cell(row, c++).Value = r.DealerName ?? "";
                ws.Cell(row, c++).Value = r.ConsRepCode ?? "";
                ws.Cell(row, c++).Value = r.ConsRepName ?? "";
                ws.Cell(row, c++).Value = r.ServiceDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.VisitCode ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeId ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeDesc ?? "";
                ws.Cell(row, c++).Value = r.MachineType ?? "";
                ws.Cell(row, c++).Value = r.AdRemort ?? "";
                ws.Cell(row, c++).Value = r.CusType ?? "";
                ws.Cell(row, c++).Value = r.SageCode ?? "";
                ws.Cell(row, c++).Value = r.SubCusType ?? "";
                ws.Cell(row, c++).Value = r.ConsComment ?? "";
                ws.Cell(row, c++).Value = r.CusEmail ?? "";
                ws.Cell(row, c++).Value = r.WarrantyYears?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.CnNo ?? "";
                ws.Cell(row, c++).Value = r.MLocTelephone ?? "";

                // Zebra striping
                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 241, 250);

                row++;
            }

            // Auto-fit columns, freeze header row
            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // PDF  (NuGet: QuestPDF)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadNsReportPdf(string startDate, string lastDate)
        {
            var data = GetNsReportData(startDate, lastDate);

            // Only the most useful columns in landscape PDF (65 columns won't fit)
            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    // ── Header ──────────────────────────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text($"NS Customer Report")
                                .Bold().FontSize(14).FontColor(Colors.Blue.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Blue.Darken3);
                    });

                    // ── Content ─────────────────────────────────────────────
                    page.Content().PaddingTop(8).Table(table =>
                    {
                        // Column definitions (12 key columns in landscape)
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(1.2f); // Cus Code
                            cols.RelativeColumn(2.5f); // Customer Name
                            cols.RelativeColumn(1.0f); // Tel No
                            cols.RelativeColumn(1.2f); // Machine Ref
                            cols.RelativeColumn(1.5f); // Machine Desc
                            cols.RelativeColumn(1.0f); // Machine Type
                            cols.RelativeColumn(1.0f); // Status
                            cols.RelativeColumn(1.0f); // Team
                            cols.RelativeColumn(1.2f); // Warranty End
                            cols.RelativeColumn(1.2f); // MA Period End
                            cols.RelativeColumn(1.0f); // Visits/Yr
                            cols.RelativeColumn(1.5f); // Email
                        });

                        // ── Column headers ──────────────────────────────────
                        var colHeaders = new[]
                        {
                            "Cus Code", "Customer Name", "Tel No",
                            "Machine Ref", "Machine Desc", "Type",
                            "Status", "Team",
                            "Warranty End", "MA End", "Visits/Yr", "Email"
                        };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Blue.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        // ── Data rows ───────────────────────────────────────
                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Blue.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(r.CusCode);
                            Cell(r.CusName);
                            Cell(r.TelNo);
                            Cell(r.MachineRefCode);
                            Cell(r.MachineDesc);
                            Cell(r.MachineType);
                            Cell(r.CusStatus);
                            Cell(r.Team);
                            Cell(r.MWarrantyEndDate?.ToString("yyyy-MM-dd"));
                            Cell(r.MaPeriodEnd?.ToString("yyyy-MM-dd"));
                            Cell(r.VisitsPerYear?.ToString());
                            Cell(r.CusEmail);
                        }
                    });

                    // ── Footer ───────────────────────────────────────────────
                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // FS REPORT: fetch data from DB
        // ──────────────────────────────────────────────────────────────────────
        public List<FsReportDto> GetFsReportData(string startDate, string lastDate)
        {
            DateTime fromDate = DateTime.Parse(startDate);
            DateTime toDate = DateTime.Parse(lastDate);

            const string query = @"
        SELECT
             [COM_ID]               AS ComId
            ,[SERIAL_NO]            AS SerialNo
            ,[DIVISION_ID]          AS DivisionId
            ,[AGREEMENT_ID]         AS AgreementId
            ,[CUS_CODE]             AS CusCode
            ,[CUS_NAME]             AS CusName
            ,[INV_ADD1]             AS InvAdd1
            ,[INV_ADD2]             AS InvAdd2
            ,[INV_ADD3]             AS InvAdd3
            ,[TEL_NO]               AS TelNo
            ,[FAX_NO]               AS FaxNo
            ,[CUS_GRADE]            AS CusGrade
            ,[MACHINE_REF_CODE]     AS MachineRefCode
            ,[MACHINE_CODE]         AS MachineCode
            ,[MACHINE_DESC]         AS MachineDesc
            ,[M_LOC1]               AS MLoc1
            ,[M_LOC2]               AS MLoc2
            ,[M_LOC3]               AS MLoc3
            ,[MACHINE_MODE]         AS MachineMode
            ,[M_LOC_CONTACT_NAME]   AS MLocContactName
            ,[M_LOC_CONTACT_NO]     AS MLocContactNo
            ,[MFP_CITY]             AS MfpCity
            ,[T_OFFICER_CODE]       AS TOfficerCode
            ,[T_OFFICER_NAME]       AS TOfficerName
            ,[M_INV_NO]             AS MInvNo
            ,[M_INV_DATE]           AS MInvDate
            ,[M_INS_DATE]           AS MInsDate
            ,[M_WARRANTY_END_DATE]  AS MWarrantyEndDate
            ,[LABOR_WARRANTY]       AS LaborWarranty
            ,[PARTS_WARRANTY]       AS PartsWarranty
            ,[IS_TRANSPORT]         AS IsTransport
            ,[PROVISION]            AS Provision
            ,[CUS_STATUS]           AS CusStatus
            ,[VISITS_PER_YEAR]      AS VisitsPerYear
            ,[INV_NO_TC]            AS InvNoTc
            ,[INV_DATE_TC]          AS InvDateTc
            ,[TEAM]                 AS Team
            ,[MA_PERIOD_START]      AS MaPeriodStart
            ,[MA_PERIOD_END]        AS MaPeriodEnd
            ,[SPECIAL]              AS Special
            ,[DATA_CHECKED]         AS DataChecked
            ,[REP_CODE]             AS RepCode
            ,[REP_NAME]             AS RepName
            ,[TO_CODE]              AS ToCode
            ,[TO_NAME]              AS ToName
            ,[TO_AREA_CODE]         AS ToAreaCode
            ,[TO_AREA_NAME]         AS ToAreaName
            ,[DEALER_ID]            AS DealerId
            ,[DEALER_NAME]          AS DealerName
            ,[CONS_REP_CODE]        AS ConsRepCode
            ,[CONS_REP_NAME]        AS ConsRepName
            ,[SERVICE_DATE]         AS ServiceDate
            ,[VISIT_CODE]           AS VisitCode
            ,[WARRANTY_TYPE_ID]     AS WarrantyTypeId
            ,[WARRANTY_TYPE_DESC]   AS WarrantyTypeDesc
            ,[MACHINE_TYPE]         AS MachineType
            ,[AD_REMORT]            AS AdRemort
            ,[CUS_TYPE]             AS CusType
            ,[SAGE_CODE]            AS SageCode
            ,[SUB_CUS_TYPE]         AS SubCusType
            ,[CONS_COMMENT]         AS ConsComment
            ,[CUS_EMAIL]            AS CusEmail
            ,[WARRANTY_YEARS]       AS WarrantyYears
            ,[CN_NO]                AS CnNo
            ,[M_LOC_TELEPHONE]      AS MLocTelephone
        FROM TBL_MACHINE_TRANSACTION
        WHERE
            [CUS_STATUS] = 'FS'
            AND [MACHINE_TYPE] in ('MFP','BCP','MPC')
            AND [M_WARRANTY_END_DATE] > @fromDate
            AND [M_WARRANTY_END_DATE] < @toDate
        ORDER BY [CUS_CODE]";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<FsReportDto>(query, new { fromDate, toDate }).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetFsReportData: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetFsReportData: {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // FS REPORT: EXCEL (NuGet: ClosedXML)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadFsReportExcel(string startDate, string lastDate)
        {
            var data = GetFsReportData(startDate, lastDate);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("FS Report");

            // ── Header row ──────────────────────────────────────────────────
            var headers = new[]
            {
        "COM ID","Serial No","Division","Agreement","Cus Code","Customer Name",
        "Inv Add1","Inv Add2","Inv Add3","Tel No","Fax No","Grade",
        "Machine Ref","Machine Code","Machine Desc",
        "M Loc1","M Loc2","M Loc3","Machine Mode",
        "Loc Contact","Loc Contact No","City",
        "Officer Code","Officer Name",
        "Inv No","Inv Date","Ins Date","Warranty End",
        "Labor Warranty","Parts Warranty","Transport","Provision",
        "Status","Visits/Yr","Inv No TC","Inv Date TC",
        "Team","MA Start","MA End","Special","Data Checked",
        "Rep Code","Rep Name","TO Code","TO Name","Area Code","Area Name",
        "Dealer ID","Dealer Name","Cons Rep Code","Cons Rep Name",
        "Service Date","Visit Code","Warranty Type ID","Warranty Type",
        "Machine Type","Ad Remort","Cus Type","Sage Code","Sub Cus Type",
        "Comment","Email","Warranty Yrs","CN No","Loc Telephone"
    };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 0); // Dark green for FS
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // ── Data rows ───────────────────────────────────────────────────
            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.ComId ?? "";
                ws.Cell(row, c++).Value = r.SerialNo ?? "";
                ws.Cell(row, c++).Value = r.DivisionId ?? "";
                ws.Cell(row, c++).Value = r.AgreementId ?? "";
                ws.Cell(row, c++).Value = r.CusCode ?? "";
                ws.Cell(row, c++).Value = r.CusName ?? "";
                ws.Cell(row, c++).Value = r.InvAdd1 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd2 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd3 ?? "";
                ws.Cell(row, c++).Value = r.TelNo ?? "";
                ws.Cell(row, c++).Value = r.FaxNo ?? "";
                ws.Cell(row, c++).Value = r.CusGrade ?? "";
                ws.Cell(row, c++).Value = r.MachineRefCode ?? "";
                ws.Cell(row, c++).Value = r.MachineCode ?? "";
                ws.Cell(row, c++).Value = r.MachineDesc ?? "";
                ws.Cell(row, c++).Value = r.MLoc1 ?? "";
                ws.Cell(row, c++).Value = r.MLoc2 ?? "";
                ws.Cell(row, c++).Value = r.MLoc3 ?? "";
                ws.Cell(row, c++).Value = r.MachineMode ?? "";
                ws.Cell(row, c++).Value = r.MLocContactName ?? "";
                ws.Cell(row, c++).Value = r.MLocContactNo ?? "";
                ws.Cell(row, c++).Value = r.MfpCity ?? "";
                ws.Cell(row, c++).Value = r.TOfficerCode ?? "";
                ws.Cell(row, c++).Value = r.TOfficerName ?? "";
                ws.Cell(row, c++).Value = r.MInvNo ?? "";
                ws.Cell(row, c++).Value = r.MInvDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MInsDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MWarrantyEndDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.LaborWarranty ?? "";
                ws.Cell(row, c++).Value = r.PartsWarranty ?? "";
                ws.Cell(row, c++).Value = r.IsTransport ?? "";
                ws.Cell(row, c++).Value = r.Provision ?? "";
                ws.Cell(row, c++).Value = r.CusStatus ?? "";
                ws.Cell(row, c++).Value = r.VisitsPerYear?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.InvNoTc ?? "";
                ws.Cell(row, c++).Value = r.InvDateTc?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Team ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodStart?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodEnd?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Special ?? "";
                ws.Cell(row, c++).Value = r.DataChecked ?? "";
                ws.Cell(row, c++).Value = r.RepCode ?? "";
                ws.Cell(row, c++).Value = r.RepName ?? "";
                ws.Cell(row, c++).Value = r.ToCode ?? "";
                ws.Cell(row, c++).Value = r.ToName ?? "";
                ws.Cell(row, c++).Value = r.ToAreaCode ?? "";
                ws.Cell(row, c++).Value = r.ToAreaName ?? "";
                ws.Cell(row, c++).Value = r.DealerId ?? "";
                ws.Cell(row, c++).Value = r.DealerName ?? "";
                ws.Cell(row, c++).Value = r.ConsRepCode ?? "";
                ws.Cell(row, c++).Value = r.ConsRepName ?? "";
                ws.Cell(row, c++).Value = r.ServiceDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.VisitCode ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeId ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeDesc ?? "";
                ws.Cell(row, c++).Value = r.MachineType ?? "";
                ws.Cell(row, c++).Value = r.AdRemort ?? "";
                ws.Cell(row, c++).Value = r.CusType ?? "";
                ws.Cell(row, c++).Value = r.SageCode ?? "";
                ws.Cell(row, c++).Value = r.SubCusType ?? "";
                ws.Cell(row, c++).Value = r.ConsComment ?? "";
                ws.Cell(row, c++).Value = r.CusEmail ?? "";
                ws.Cell(row, c++).Value = r.WarrantyYears?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.CnNo ?? "";
                ws.Cell(row, c++).Value = r.MLocTelephone ?? "";

                // Zebra striping
                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 250, 235);

                row++;
            }

            // Auto-fit columns, freeze header row
            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // FS REPORT: PDF (NuGet: QuestPDF)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadFsReportPdf(string startDate, string lastDate)
        {
            var data = GetFsReportData(startDate, lastDate);

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    // ── Header ──────────────────────────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text($"FS Customer Report")
                                .Bold().FontSize(14).FontColor(Colors.Green.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Green.Darken3);
                    });

                    // ── Content ─────────────────────────────────────────────
                    page.Content().PaddingTop(8).Table(table =>
                    {
                        // Column definitions (12 key columns in landscape)
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(1.2f); // Cus Code
                            cols.RelativeColumn(2.5f); // Customer Name
                            cols.RelativeColumn(1.0f); // Tel No
                            cols.RelativeColumn(1.2f); // Machine Ref
                            cols.RelativeColumn(1.5f); // Machine Desc
                            cols.RelativeColumn(1.0f); // Machine Type
                            cols.RelativeColumn(1.0f); // Status
                            cols.RelativeColumn(1.0f); // Team
                            cols.RelativeColumn(1.2f); // Warranty End
                            cols.RelativeColumn(1.2f); // MA Period End
                            cols.RelativeColumn(1.0f); // Visits/Yr
                            cols.RelativeColumn(1.5f); // Email
                        });

                        // ── Column headers ──────────────────────────────────
                        var colHeaders = new[]
                        {
                    "Cus Code", "Customer Name", "Tel No",
                    "Machine Ref", "Machine Desc", "Type",
                    "Status", "Team",
                    "Warranty End", "MA End", "Visits/Yr", "Email"
                };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Green.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        // ── Data rows ───────────────────────────────────────
                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Green.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(r.CusCode);
                            Cell(r.CusName);
                            Cell(r.TelNo);
                            Cell(r.MachineRefCode);
                            Cell(r.MachineDesc);
                            Cell(r.MachineType);
                            Cell(r.CusStatus);
                            Cell(r.Team);
                            Cell(r.MWarrantyEndDate?.ToString("yyyy-MM-dd"));
                            Cell(r.MaPeriodEnd?.ToString("yyyy-MM-dd"));
                            Cell(r.VisitsPerYear?.ToString());
                            Cell(r.CusEmail);
                        }
                    });

                    // ── Footer ───────────────────────────────────────────────
                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // MA REPORT: fetch data from DB
        // ──────────────────────────────────────────────────────────────────────
        public List<MaReportDto> GetMaReportData(string startDate, string lastDate)
        {
            DateTime fromDate = DateTime.Parse(startDate);
            DateTime toDate = DateTime.Parse(lastDate);

            const string query = @"
        SELECT
             [COM_ID]               AS ComId
            ,[SERIAL_NO]            AS SerialNo
            ,[DIVISION_ID]          AS DivisionId
            ,[AGREEMENT_ID]         AS AgreementId
            ,[CUS_CODE]             AS CusCode
            ,[CUS_NAME]             AS CusName
            ,[INV_ADD1]             AS InvAdd1
            ,[INV_ADD2]             AS InvAdd2
            ,[INV_ADD3]             AS InvAdd3
            ,[TEL_NO]               AS TelNo
            ,[FAX_NO]               AS FaxNo
            ,[CUS_GRADE]            AS CusGrade
            ,[MACHINE_REF_CODE]     AS MachineRefCode
            ,[MACHINE_CODE]         AS MachineCode
            ,[MACHINE_DESC]         AS MachineDesc
            ,[M_LOC1]               AS MLoc1
            ,[M_LOC2]               AS MLoc2
            ,[M_LOC3]               AS MLoc3
            ,[MACHINE_MODE]         AS MachineMode
            ,[M_LOC_CONTACT_NAME]   AS MLocContactName
            ,[M_LOC_CONTACT_NO]     AS MLocContactNo
            ,[MFP_CITY]             AS MfpCity
            ,[T_OFFICER_CODE]       AS TOfficerCode
            ,[T_OFFICER_NAME]       AS TOfficerName
            ,[M_INV_NO]             AS MInvNo
            ,[M_INV_DATE]           AS MInvDate
            ,[M_INS_DATE]           AS MInsDate
            ,[M_WARRANTY_END_DATE]  AS MWarrantyEndDate
            ,[LABOR_WARRANTY]       AS LaborWarranty
            ,[PARTS_WARRANTY]       AS PartsWarranty
            ,[IS_TRANSPORT]         AS IsTransport
            ,[PROVISION]            AS Provision
            ,[CUS_STATUS]           AS CusStatus
            ,[VISITS_PER_YEAR]      AS VisitsPerYear
            ,[INV_NO_TC]            AS InvNoTc
            ,[INV_DATE_TC]          AS InvDateTc
            ,[TEAM]                 AS Team
            ,[MA_PERIOD_START]      AS MaPeriodStart
            ,[MA_PERIOD_END]        AS MaPeriodEnd
            ,[SPECIAL]              AS Special
            ,[DATA_CHECKED]         AS DataChecked
            ,[REP_CODE]             AS RepCode
            ,[REP_NAME]             AS RepName
            ,[TO_CODE]              AS ToCode
            ,[TO_NAME]              AS ToName
            ,[TO_AREA_CODE]         AS ToAreaCode
            ,[TO_AREA_NAME]         AS ToAreaName
            ,[DEALER_ID]            AS DealerId
            ,[DEALER_NAME]          AS DealerName
            ,[CONS_REP_CODE]        AS ConsRepCode
            ,[CONS_REP_NAME]        AS ConsRepName
            ,[SERVICE_DATE]         AS ServiceDate
            ,[VISIT_CODE]           AS VisitCode
            ,[WARRANTY_TYPE_ID]     AS WarrantyTypeId
            ,[WARRANTY_TYPE_DESC]   AS WarrantyTypeDesc
            ,[MACHINE_TYPE]         AS MachineType
            ,[AD_REMORT]            AS AdRemort
            ,[CUS_TYPE]             AS CusType
            ,[SAGE_CODE]            AS SageCode
            ,[SUB_CUS_TYPE]         AS SubCusType
            ,[CONS_COMMENT]         AS ConsComment
            ,[CUS_EMAIL]            AS CusEmail
            ,[WARRANTY_YEARS]       AS WarrantyYears
            ,[CN_NO]                AS CnNo
            ,[M_LOC_TELEPHONE]      AS MLocTelephone
        FROM TBL_MACHINE_TRANSACTION
        WHERE
            [CUS_STATUS] = 'MA'
            AND [MACHINE_TYPE] IN ('MFP','BCP','MPC','DCP','MF')
            AND [MA_PERIOD_END] > @fromDate
            AND [MA_PERIOD_END] < @toDate
        ORDER BY [CUS_CODE]";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<MaReportDto>(query, new { fromDate, toDate }).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetMaReportData: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetMaReportData: {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // MA REPORT: EXCEL
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadMaReportExcel(string startDate, string lastDate)
        {
            var data = GetMaReportData(startDate, lastDate);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("MA Report");

            var headers = new[]
            {
        "COM ID","Serial No","Division","Agreement","Cus Code","Customer Name",
        "Inv Add1","Inv Add2","Inv Add3","Tel No","Fax No","Grade",
        "Machine Ref","Machine Code","Machine Desc",
        "M Loc1","M Loc2","M Loc3","Machine Mode",
        "Loc Contact","Loc Contact No","City",
        "Officer Code","Officer Name",
        "Inv No","Inv Date","Ins Date","Warranty End",
        "Labor Warranty","Parts Warranty","Transport","Provision",
        "Status","Visits/Yr","Inv No TC","Inv Date TC",
        "Team","MA Start","MA End","Special","Data Checked",
        "Rep Code","Rep Name","TO Code","TO Name","Area Code","Area Name",
        "Dealer ID","Dealer Name","Cons Rep Code","Cons Rep Name",
        "Service Date","Visit Code","Warranty Type ID","Warranty Type",
        "Machine Type","Ad Remort","Cus Type","Sage Code","Sub Cus Type",
        "Comment","Email","Warranty Yrs","CN No","Loc Telephone"
    };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(128, 0, 128); // Purple for MA
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.ComId ?? "";
                ws.Cell(row, c++).Value = r.SerialNo ?? "";
                ws.Cell(row, c++).Value = r.DivisionId ?? "";
                ws.Cell(row, c++).Value = r.AgreementId ?? "";
                ws.Cell(row, c++).Value = r.CusCode ?? "";
                ws.Cell(row, c++).Value = r.CusName ?? "";
                ws.Cell(row, c++).Value = r.InvAdd1 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd2 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd3 ?? "";
                ws.Cell(row, c++).Value = r.TelNo ?? "";
                ws.Cell(row, c++).Value = r.FaxNo ?? "";
                ws.Cell(row, c++).Value = r.CusGrade ?? "";
                ws.Cell(row, c++).Value = r.MachineRefCode ?? "";
                ws.Cell(row, c++).Value = r.MachineCode ?? "";
                ws.Cell(row, c++).Value = r.MachineDesc ?? "";
                ws.Cell(row, c++).Value = r.MLoc1 ?? "";
                ws.Cell(row, c++).Value = r.MLoc2 ?? "";
                ws.Cell(row, c++).Value = r.MLoc3 ?? "";
                ws.Cell(row, c++).Value = r.MachineMode ?? "";
                ws.Cell(row, c++).Value = r.MLocContactName ?? "";
                ws.Cell(row, c++).Value = r.MLocContactNo ?? "";
                ws.Cell(row, c++).Value = r.MfpCity ?? "";
                ws.Cell(row, c++).Value = r.TOfficerCode ?? "";
                ws.Cell(row, c++).Value = r.TOfficerName ?? "";
                ws.Cell(row, c++).Value = r.MInvNo ?? "";
                ws.Cell(row, c++).Value = r.MInvDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MInsDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MWarrantyEndDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.LaborWarranty ?? "";
                ws.Cell(row, c++).Value = r.PartsWarranty ?? "";
                ws.Cell(row, c++).Value = r.IsTransport ?? "";
                ws.Cell(row, c++).Value = r.Provision ?? "";
                ws.Cell(row, c++).Value = r.CusStatus ?? "";
                ws.Cell(row, c++).Value = r.VisitsPerYear?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.InvNoTc ?? "";
                ws.Cell(row, c++).Value = r.InvDateTc?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Team ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodStart?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodEnd?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Special ?? "";
                ws.Cell(row, c++).Value = r.DataChecked ?? "";
                ws.Cell(row, c++).Value = r.RepCode ?? "";
                ws.Cell(row, c++).Value = r.RepName ?? "";
                ws.Cell(row, c++).Value = r.ToCode ?? "";
                ws.Cell(row, c++).Value = r.ToName ?? "";
                ws.Cell(row, c++).Value = r.ToAreaCode ?? "";
                ws.Cell(row, c++).Value = r.ToAreaName ?? "";
                ws.Cell(row, c++).Value = r.DealerId ?? "";
                ws.Cell(row, c++).Value = r.DealerName ?? "";
                ws.Cell(row, c++).Value = r.ConsRepCode ?? "";
                ws.Cell(row, c++).Value = r.ConsRepName ?? "";
                ws.Cell(row, c++).Value = r.ServiceDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.VisitCode ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeId ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeDesc ?? "";
                ws.Cell(row, c++).Value = r.MachineType ?? "";
                ws.Cell(row, c++).Value = r.AdRemort ?? "";
                ws.Cell(row, c++).Value = r.CusType ?? "";
                ws.Cell(row, c++).Value = r.SageCode ?? "";
                ws.Cell(row, c++).Value = r.SubCusType ?? "";
                ws.Cell(row, c++).Value = r.ConsComment ?? "";
                ws.Cell(row, c++).Value = r.CusEmail ?? "";
                ws.Cell(row, c++).Value = r.WarrantyYears?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.CnNo ?? "";
                ws.Cell(row, c++).Value = r.MLocTelephone ?? "";

                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(243, 235, 250);

                row++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // MA REPORT: PDF
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadMaReportPdf(string startDate, string lastDate)
        {
            var data = GetMaReportData(startDate, lastDate);

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("MA Customer Report")
                                .Bold().FontSize(14).FontColor(Colors.Purple.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Purple.Darken3);
                    });

                    page.Content().PaddingTop(8).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(1.2f); // Cus Code
                            cols.RelativeColumn(2.5f); // Customer Name
                            cols.RelativeColumn(1.0f); // Tel No
                            cols.RelativeColumn(1.2f); // Machine Ref
                            cols.RelativeColumn(1.5f); // Machine Desc
                            cols.RelativeColumn(1.0f); // Machine Type
                            cols.RelativeColumn(1.0f); // Status
                            cols.RelativeColumn(1.0f); // Team
                            cols.RelativeColumn(1.2f); // Warranty End
                            cols.RelativeColumn(1.2f); // MA Period End
                            cols.RelativeColumn(1.0f); // Visits/Yr
                            cols.RelativeColumn(1.5f); // Email
                        });

                        var colHeaders = new[]
                        {
                    "Cus Code", "Customer Name", "Tel No",
                    "Machine Ref", "Machine Desc", "Type",
                    "Status", "Team",
                    "Warranty End", "MA End", "Visits/Yr", "Email"
                };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Purple.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Purple.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(r.CusCode);
                            Cell(r.CusName);
                            Cell(r.TelNo);
                            Cell(r.MachineRefCode);
                            Cell(r.MachineDesc);
                            Cell(r.MachineType);
                            Cell(r.CusStatus);
                            Cell(r.Team);
                            Cell(r.MWarrantyEndDate?.ToString("yyyy-MM-dd"));
                            Cell(r.MaPeriodEnd?.ToString("yyyy-MM-dd"));
                            Cell(r.VisitsPerYear?.ToString());
                            Cell(r.CusEmail);
                        }
                    });

                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // EX REPORT: fetch data from DB
        // ──────────────────────────────────────────────────────────────────────
        public List<ExReportDto> GetExReportData(string startDate, string lastDate)
        {
            DateTime fromDate = DateTime.Parse(startDate);
            DateTime toDate = DateTime.Parse(lastDate);

            const string query = @"
        SELECT
             [COM_ID]               AS ComId
            ,[SERIAL_NO]            AS SerialNo
            ,[DIVISION_ID]          AS DivisionId
            ,[AGREEMENT_ID]         AS AgreementId
            ,[CUS_CODE]             AS CusCode
            ,[CUS_NAME]             AS CusName
            ,[INV_ADD1]             AS InvAdd1
            ,[INV_ADD2]             AS InvAdd2
            ,[INV_ADD3]             AS InvAdd3
            ,[TEL_NO]               AS TelNo
            ,[FAX_NO]               AS FaxNo
            ,[CUS_GRADE]            AS CusGrade
            ,[MACHINE_REF_CODE]     AS MachineRefCode
            ,[MACHINE_CODE]         AS MachineCode
            ,[MACHINE_DESC]         AS MachineDesc
            ,[M_LOC1]               AS MLoc1
            ,[M_LOC2]               AS MLoc2
            ,[M_LOC3]               AS MLoc3
            ,[MACHINE_MODE]         AS MachineMode
            ,[M_LOC_CONTACT_NAME]   AS MLocContactName
            ,[M_LOC_CONTACT_NO]     AS MLocContactNo
            ,[MFP_CITY]             AS MfpCity
            ,[T_OFFICER_CODE]       AS TOfficerCode
            ,[T_OFFICER_NAME]       AS TOfficerName
            ,[M_INV_NO]             AS MInvNo
            ,[M_INV_DATE]           AS MInvDate
            ,[M_INS_DATE]           AS MInsDate
            ,[M_WARRANTY_END_DATE]  AS MWarrantyEndDate
            ,[LABOR_WARRANTY]       AS LaborWarranty
            ,[PARTS_WARRANTY]       AS PartsWarranty
            ,[IS_TRANSPORT]         AS IsTransport
            ,[PROVISION]            AS Provision
            ,[CUS_STATUS]           AS CusStatus
            ,[VISITS_PER_YEAR]      AS VisitsPerYear
            ,[INV_NO_TC]            AS InvNoTc
            ,[INV_DATE_TC]          AS InvDateTc
            ,[TEAM]                 AS Team
            ,[MA_PERIOD_START]      AS MaPeriodStart
            ,[MA_PERIOD_END]        AS MaPeriodEnd
            ,[SPECIAL]              AS Special
            ,[DATA_CHECKED]         AS DataChecked
            ,[REP_CODE]             AS RepCode
            ,[REP_NAME]             AS RepName
            ,[TO_CODE]              AS ToCode
            ,[TO_NAME]              AS ToName
            ,[TO_AREA_CODE]         AS ToAreaCode
            ,[TO_AREA_NAME]         AS ToAreaName
            ,[DEALER_ID]            AS DealerId
            ,[DEALER_NAME]          AS DealerName
            ,[CONS_REP_CODE]        AS ConsRepCode
            ,[CONS_REP_NAME]        AS ConsRepName
            ,[SERVICE_DATE]         AS ServiceDate
            ,[VISIT_CODE]           AS VisitCode
            ,[WARRANTY_TYPE_ID]     AS WarrantyTypeId
            ,[WARRANTY_TYPE_DESC]   AS WarrantyTypeDesc
            ,[MACHINE_TYPE]         AS MachineType
            ,[AD_REMORT]            AS AdRemort
            ,[CUS_TYPE]             AS CusType
            ,[SAGE_CODE]            AS SageCode
            ,[SUB_CUS_TYPE]         AS SubCusType
            ,[CONS_COMMENT]         AS ConsComment
            ,[CUS_EMAIL]            AS CusEmail
            ,[WARRANTY_YEARS]       AS WarrantyYears
            ,[CN_NO]                AS CnNo
            ,[M_LOC_TELEPHONE]      AS MLocTelephone
        FROM TBL_MACHINE_TRANSACTION
        WHERE
            [CUS_STATUS] = 'EX'
            AND [MACHINE_TYPE] in ('MFP','BCP','MPC')
            AND (
                (
                    [MA_PERIOD_END] IS NOT NULL
                    AND [MA_PERIOD_END] > @fromDate
                    AND [MA_PERIOD_END] < @toDate
                )
                OR
                (
                    [MA_PERIOD_END] IS NULL
                    AND [M_WARRANTY_END_DATE] > @fromDate
                    AND [M_WARRANTY_END_DATE] < @toDate
                )
            )
        ORDER BY [CUS_CODE]";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<ExReportDto>(query, new { fromDate, toDate }).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetExReportData: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetExReportData: {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // EX REPORT: EXCEL (NuGet: ClosedXML)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadExReportExcel(string startDate, string lastDate)
        {
            var data = GetExReportData(startDate, lastDate);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("EX Report");

            // ── Header row ──────────────────────────────────────────────────
            var headers = new[]
            {
        "COM ID","Serial No","Division","Agreement","Cus Code","Customer Name",
        "Inv Add1","Inv Add2","Inv Add3","Tel No","Fax No","Grade",
        "Machine Ref","Machine Code","Machine Desc",
        "M Loc1","M Loc2","M Loc3","Machine Mode",
        "Loc Contact","Loc Contact No","City",
        "Officer Code","Officer Name",
        "Inv No","Inv Date","Ins Date","Warranty End",
        "Labor Warranty","Parts Warranty","Transport","Provision",
        "Status","Visits/Yr","Inv No TC","Inv Date TC",
        "Team","MA Start","MA End","Special","Data Checked",
        "Rep Code","Rep Name","TO Code","TO Name","Area Code","Area Name",
        "Dealer ID","Dealer Name","Cons Rep Code","Cons Rep Name",
        "Service Date","Visit Code","Warranty Type ID","Warranty Type",
        "Machine Type","Ad Remort","Cus Type","Sage Code","Sub Cus Type",
        "Comment","Email","Warranty Yrs","CN No","Loc Telephone"
    };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(192, 0, 0); // Dark red for EX
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // ── Data rows ───────────────────────────────────────────────────
            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.ComId ?? "";
                ws.Cell(row, c++).Value = r.SerialNo ?? "";
                ws.Cell(row, c++).Value = r.DivisionId ?? "";
                ws.Cell(row, c++).Value = r.AgreementId ?? "";
                ws.Cell(row, c++).Value = r.CusCode ?? "";
                ws.Cell(row, c++).Value = r.CusName ?? "";
                ws.Cell(row, c++).Value = r.InvAdd1 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd2 ?? "";
                ws.Cell(row, c++).Value = r.InvAdd3 ?? "";
                ws.Cell(row, c++).Value = r.TelNo ?? "";
                ws.Cell(row, c++).Value = r.FaxNo ?? "";
                ws.Cell(row, c++).Value = r.CusGrade ?? "";
                ws.Cell(row, c++).Value = r.MachineRefCode ?? "";
                ws.Cell(row, c++).Value = r.MachineCode ?? "";
                ws.Cell(row, c++).Value = r.MachineDesc ?? "";
                ws.Cell(row, c++).Value = r.MLoc1 ?? "";
                ws.Cell(row, c++).Value = r.MLoc2 ?? "";
                ws.Cell(row, c++).Value = r.MLoc3 ?? "";
                ws.Cell(row, c++).Value = r.MachineMode ?? "";
                ws.Cell(row, c++).Value = r.MLocContactName ?? "";
                ws.Cell(row, c++).Value = r.MLocContactNo ?? "";
                ws.Cell(row, c++).Value = r.MfpCity ?? "";
                ws.Cell(row, c++).Value = r.TOfficerCode ?? "";
                ws.Cell(row, c++).Value = r.TOfficerName ?? "";
                ws.Cell(row, c++).Value = r.MInvNo ?? "";
                ws.Cell(row, c++).Value = r.MInvDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MInsDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MWarrantyEndDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.LaborWarranty ?? "";
                ws.Cell(row, c++).Value = r.PartsWarranty ?? "";
                ws.Cell(row, c++).Value = r.IsTransport ?? "";
                ws.Cell(row, c++).Value = r.Provision ?? "";
                ws.Cell(row, c++).Value = r.CusStatus ?? "";
                ws.Cell(row, c++).Value = r.VisitsPerYear?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.InvNoTc ?? "";
                ws.Cell(row, c++).Value = r.InvDateTc?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Team ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodStart?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MaPeriodEnd?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.Special ?? "";
                ws.Cell(row, c++).Value = r.DataChecked ?? "";
                ws.Cell(row, c++).Value = r.RepCode ?? "";
                ws.Cell(row, c++).Value = r.RepName ?? "";
                ws.Cell(row, c++).Value = r.ToCode ?? "";
                ws.Cell(row, c++).Value = r.ToName ?? "";
                ws.Cell(row, c++).Value = r.ToAreaCode ?? "";
                ws.Cell(row, c++).Value = r.ToAreaName ?? "";
                ws.Cell(row, c++).Value = r.DealerId ?? "";
                ws.Cell(row, c++).Value = r.DealerName ?? "";
                ws.Cell(row, c++).Value = r.ConsRepCode ?? "";
                ws.Cell(row, c++).Value = r.ConsRepName ?? "";
                ws.Cell(row, c++).Value = r.ServiceDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.VisitCode ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeId ?? "";
                ws.Cell(row, c++).Value = r.WarrantyTypeDesc ?? "";
                ws.Cell(row, c++).Value = r.MachineType ?? "";
                ws.Cell(row, c++).Value = r.AdRemort ?? "";
                ws.Cell(row, c++).Value = r.CusType ?? "";
                ws.Cell(row, c++).Value = r.SageCode ?? "";
                ws.Cell(row, c++).Value = r.SubCusType ?? "";
                ws.Cell(row, c++).Value = r.ConsComment ?? "";
                ws.Cell(row, c++).Value = r.CusEmail ?? "";
                ws.Cell(row, c++).Value = r.WarrantyYears?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.CnNo ?? "";
                ws.Cell(row, c++).Value = r.MLocTelephone ?? "";

                // Zebra striping
                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 235, 235);

                row++;
            }

            // Auto-fit columns, freeze header row
            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // EX REPORT: PDF (NuGet: QuestPDF)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadExReportPdf(string startDate, string lastDate)
        {
            var data = GetExReportData(startDate, lastDate);

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    // ── Header ──────────────────────────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text($"EX Customer Report")
                                .Bold().FontSize(14).FontColor(Colors.Red.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Red.Darken3);
                    });

                    // ── Content ─────────────────────────────────────────────
                    page.Content().PaddingTop(8).Table(table =>
                    {
                        // Column definitions (12 key columns in landscape)
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(1.2f); // Cus Code
                            cols.RelativeColumn(2.5f); // Customer Name
                            cols.RelativeColumn(1.0f); // Tel No
                            cols.RelativeColumn(1.2f); // Machine Ref
                            cols.RelativeColumn(1.5f); // Machine Desc
                            cols.RelativeColumn(1.0f); // Machine Type
                            cols.RelativeColumn(1.0f); // Status
                            cols.RelativeColumn(1.0f); // Team
                            cols.RelativeColumn(1.2f); // Warranty End
                            cols.RelativeColumn(1.2f); // MA Period End
                            cols.RelativeColumn(1.0f); // Visits/Yr
                            cols.RelativeColumn(1.5f); // Email
                        });

                        // ── Column headers ──────────────────────────────────
                        var colHeaders = new[]
                        {
                    "Cus Code", "Customer Name", "Tel No",
                    "Machine Ref", "Machine Desc", "Type",
                    "Status", "Team",
                    "Warranty End", "MA End", "Visits/Yr", "Email"
                };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Red.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        // ── Data rows ───────────────────────────────────────
                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Red.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(r.CusCode);
                            Cell(r.CusName);
                            Cell(r.TelNo);
                            Cell(r.MachineRefCode);
                            Cell(r.MachineDesc);
                            Cell(r.MachineType);
                            Cell(r.CusStatus);
                            Cell(r.Team);
                            Cell(r.MWarrantyEndDate?.ToString("yyyy-MM-dd"));
                            Cell(r.MaPeriodEnd?.ToString("yyyy-MM-dd"));
                            Cell(r.VisitsPerYear?.ToString());
                            Cell(r.CusEmail);
                        }
                    });

                    // ── Footer ───────────────────────────────────────────────
                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }


        // ─── Shared: fetch raw data (Dapper) ─────────────────────────────────────────
        private List<DailyJobReportDto> FetchJobs(string startDate, string lastDate)
        {
            DateTime fromDate = DateTime.Parse(startDate);
            DateTime toDate = DateTime.Parse(lastDate);

            string sql = @"
        SELECT
             [COM_ID]                  AS COM_ID
            ,[DJ_ID]                   AS DJ_ID
            ,[DJ_DATE]                 AS DJ_DATE
            ,[SERIAL_NO]               AS SERIAL_NO
            ,[MACHINE_REF_NO]          AS MACHINE_REF_NO
            ,[TEAM_ID]                 AS TEAM_ID
            ,[TEAM_NAME]               AS TEAM_NAME
            ,[NOTE]                    AS NOTE
            ,[INFORM_BY]               AS INFORM_BY
            ,[CONTACT_TYPE]            AS CONTACT_TYPE
            ,[CUS_ADD1]                AS CUS_ADD1
            ,[CUS_ADD2]                AS CUS_ADD2
            ,[CUS_ADD3]                AS CUS_ADD3
            ,[CUS_CONTACT]             AS CUS_CONTACT
            ,[CUS_TEL_NO]              AS CUS_TEL_NO
            ,[CUS_SMS_NO]              AS CUS_SMS_NO
            ,[MACHINE_MODEL_ID]        AS MACHINE_MODEL_ID
            ,[MACHINE_MODEL_NAME]      AS MACHINE_MODEL_NAME
            ,[CUS_STATUS]              AS CUS_STATUS
            ,[CUS_TYPE]                AS CUS_TYPE
            ,[JOB_STATUS]              AS JOB_STATUS
            ,[CUS_ID]                  AS CUS_ID
            ,[CUS_NAME]                AS CUS_NAME
            ,[TECH_CODE]               AS TECH_CODE
            ,[TECH_MOBILE]             AS TECH_MOBILE
            ,[CR_BY]                   AS CR_BY
            ,[CR_DATE]                 AS CR_DATE
            ,[COMPLETE_BY]             AS COMPLETE_BY
            ,[COMPLETED_DATE]          AS COMPLETED_DATE
            ,[CANCELLED_BY]            AS CANCELLED_BY
            ,[CANCELLED_DATE]          AS CANCELLED_DATE
            ,[SMS_DEL_REPORT_ID]       AS SMS_DEL_REPORT_ID
            ,[SMS_DELIVERED]           AS SMS_DELIVERED
            ,[SMS_DELIVERED_DATETIME]  AS SMS_DELIVERED_DATETIME
            ,[IS_TECH_NOTIFIED]        AS IS_TECH_NOTIFIED
            ,[STARTED_BY]              AS STARTED_BY
            ,[STARTED_DATE]            AS STARTED_DATE
            ,[CANCELLED_REASON]        AS CANCELLED_REASON
            ,[SOLUTION_CATEGORY]       AS SOLUTION_CATEGORY
            ,[COMPLETE_SOLUTION]       AS COMPLETE_SOLUTION
            ,[RECALL_ID]               AS RECALL_ID
            ,[START_WORK_NOTE]         AS START_WORK_NOTE
        FROM TBL_DAILY_JOBS
        WHERE DJ_DATE BETWEEN @fromDate AND @toDate
        ORDER BY DJ_DATE DESC";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                // Pass buffered: false to disable Dapper's caching mechanism
                return connection.Query<DailyJobReportDto>(sql, new { fromDate, toDate }, buffered: false).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] FetchJobs: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] FetchJobs: {ex.Message}");
                throw;
            }
        }

        // ─── 1. Get data ──────────────────────────────────────────────────────────────
        public List<DailyJobReportDto> GetDailyJobReportData(string startDate, string lastDate)
            => FetchJobs(startDate, lastDate);

        // ─── 2. Excel ─────────────────────────────────────────────────────────────────
        public byte[] DownloadDailyJobReportExcel(string startDate, string lastDate)
        {
            var jobs = FetchJobs(startDate, lastDate);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Daily Job Report");

            string[] headers =
            {
        "DJ ID","Date","Serial No","Machine Ref","Team","Note",
        "Inform By","Contact Type","Address 1","Address 2","Address 3",
        "Contact","Tel No","SMS No","Model","Customer",
        "Cus Status","Cus Type","Job Status","Tech Code","Tech Mobile",
        "Created By","Created Date","Completed By","Completed Date",
        "Cancelled By","Cancelled Date","SMS Delivered","SMS Datetime",
        "Tech Notified","Started By","Started Date","Cancel Reason",
        "Solution Cat","Solution","Recall ID","Work Note"
    };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(31, 78, 121);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int row = 2;
            foreach (var j in jobs)
            {
                int c = 1;
                ws.Cell(row, c++).Value = j.DJ_ID;
                ws.Cell(row, c++).Value = j.DJ_DATE.ToString("yyyy-MM-dd");
                ws.Cell(row, c++).Value = j.SERIAL_NO ?? "";
                ws.Cell(row, c++).Value = j.MACHINE_REF_NO ?? "";
                ws.Cell(row, c++).Value = j.TEAM_NAME ?? "";
                ws.Cell(row, c++).Value = j.NOTE ?? "";
                ws.Cell(row, c++).Value = j.INFORM_BY ?? "";
                ws.Cell(row, c++).Value = j.CONTACT_TYPE ?? "";
                ws.Cell(row, c++).Value = j.CUS_ADD1 ?? "";
                ws.Cell(row, c++).Value = j.CUS_ADD2 ?? "";
                ws.Cell(row, c++).Value = j.CUS_ADD3 ?? "";
                ws.Cell(row, c++).Value = j.CUS_CONTACT ?? "";
                ws.Cell(row, c++).Value = j.CUS_TEL_NO ?? "";
                ws.Cell(row, c++).Value = j.CUS_SMS_NO ?? "";
                ws.Cell(row, c++).Value = j.MACHINE_MODEL_NAME ?? "";
                ws.Cell(row, c++).Value = j.CUS_NAME ?? "";
                ws.Cell(row, c++).Value = j.CUS_STATUS ?? "";
                ws.Cell(row, c++).Value = j.CUS_TYPE ?? "";
                ws.Cell(row, c++).Value = j.JOB_STATUS ?? "";
                ws.Cell(row, c++).Value = j.TECH_CODE ?? "";
                ws.Cell(row, c++).Value = j.TECH_MOBILE ?? "";
                ws.Cell(row, c++).Value = j.CR_BY ?? "";
                ws.Cell(row, c++).Value = j.CR_DATE?.ToString("yyyy-MM-dd HH:mm") ?? "";
                ws.Cell(row, c++).Value = j.COMPLETE_BY ?? "";
                ws.Cell(row, c++).Value = j.COMPLETED_DATE?.ToString("yyyy-MM-dd HH:mm") ?? "";
                ws.Cell(row, c++).Value = j.CANCELLED_BY ?? "";
                ws.Cell(row, c++).Value = j.CANCELLED_DATE?.ToString("yyyy-MM-dd HH:mm") ?? "";
                ws.Cell(row, c++).Value = j.SMS_DELIVERED?.ToString() ?? "";
                ws.Cell(row, c++).Value = j.SMS_DELIVERED_DATETIME?.ToString("yyyy-MM-dd HH:mm") ?? "";
                ws.Cell(row, c++).Value = j.IS_TECH_NOTIFIED?.ToString() ?? "";
                ws.Cell(row, c++).Value = j.STARTED_BY ?? "";
                ws.Cell(row, c++).Value = j.STARTED_DATE?.ToString("yyyy-MM-dd HH:mm") ?? "";
                ws.Cell(row, c++).Value = j.CANCELLED_REASON ?? "";
                ws.Cell(row, c++).Value = j.SOLUTION_CATEGORY ?? "";
                ws.Cell(row, c++).Value = j.COMPLETE_SOLUTION ?? "";
                ws.Cell(row, c++).Value = j.RECALL_ID?.ToString() ?? "";
                ws.Cell(row, c++).Value = j.START_WORK_NOTE ?? "";

                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(217, 225, 242);

                row++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            wb.SaveAs(ms);
            return ms.ToArray();
        }

        // ─── 3. PDF (QuestPDF — matches your existing pattern) ───────────────────────
        public byte[] DownloadDailyJobReportPdf(string startDate, string lastDate)
        {
            var jobs = FetchJobs(startDate, lastDate);

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    // ── Header ────────────────────────────────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Daily Job Report")
                                .Bold().FontSize(14).FontColor(Colors.Blue.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Blue.Darken3);
                    });

                    // ── Content ───────────────────────────────────────────────────
                    page.Content().PaddingTop(8).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(0.8f); // DJ ID
                            cols.RelativeColumn(1.2f); // Date
                            cols.RelativeColumn(2.0f); // Customer
                            cols.RelativeColumn(1.5f); // Machine
                            cols.RelativeColumn(1.2f); // Job Status
                            cols.RelativeColumn(1.2f); // Team
                            cols.RelativeColumn(1.0f); // Tech
                            cols.RelativeColumn(1.0f); // Created By
                            cols.RelativeColumn(1.2f); // Created Date
                            cols.RelativeColumn(1.2f); // Completed Date
                            cols.RelativeColumn(1.2f); // Tel No
                            cols.RelativeColumn(1.2f); // Cus Type
                        });

                        var colHeaders = new[]
                        {
                    "DJ ID", "Date", "Customer", "Machine",
                    "Job Status", "Team", "Tech", "Created By",
                    "Created Date", "Completed Date", "Tel No", "Cus Type"
                };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Blue.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        bool isEven = false;
                        foreach (var j in jobs)
                        {
                            var bg = isEven ? Colors.Blue.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(j.DJ_ID.ToString());
                            Cell(j.DJ_DATE.ToString("yyyy-MM-dd"));
                            Cell(j.CUS_NAME);
                            Cell(j.MACHINE_MODEL_NAME);
                            Cell(j.JOB_STATUS);
                            Cell(j.TEAM_NAME);
                            Cell(j.TECH_CODE);
                            Cell(j.CR_BY);
                            Cell(j.CR_DATE?.ToString("yyyy-MM-dd HH:mm"));
                            Cell(j.COMPLETED_DATE?.ToString("yyyy-MM-dd HH:mm"));
                            Cell(j.CUS_TEL_NO);
                            Cell(j.CUS_TYPE);
                        }
                    });

                    // ── Footer ────────────────────────────────────────────────────
                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // CUSTOMER REPORT: fetch data
        // ──────────────────────────────────────────────────────────────────────
        public List<CustomerReportDto> GetCustomerReportData()
        {
            const string sql = @"
        SELECT
             CM.[COM_ID]          AS ComId
            ,CM.[CUS_CODE]        AS CusCode
            ,CM.[SAGE_CODE]       AS SageCode
            ,CM.[CUS_NAME]        AS CusName
            ,CM.[ADD1]            AS Add1
            ,CM.[ADD2]            AS Add2
            ,CM.[ADD3]            AS Add3
            ,CM.[CONTACT_PERSON]  AS ContactPerson
            ,CM.[MOBILE_NO]       AS MobileNo
            ,CM.[TEL_NO]          AS TelNo
            ,CM.[FAX]             AS Fax
            ,CM.[EMAIL]           AS Email
            ,CM.[CITY]            AS City
            ,CM.[POSTAL_CODE]     AS PostalCode
            ,CM.[CUS_STATUS]      AS CusStatus
            ,CM.[GPS_LOC1]        AS GpsLoc1
            ,CM.[GPS_LOC2]        AS GpsLoc2
            ,CM.[CUS_GRADE_ID]    AS CusGradeId
            ,CM.[CUS_GRADE_NAME]  AS CusGradeName
            ,CM.[REP_CODE]        AS RepCode
            ,CM.[REP_NAME]        AS RepName
            ,CM.[TECH_CODE]       AS TechCode
            ,CM.[TECH_NAME]       AS TechName
            ,CM.[DEBT_CODE]       AS DebtCode
            ,CM.[CONS_REP_CODE]   AS ConsRepCode
            ,COUNT(MT.[SERIAL_NO]) AS MachineCount
        FROM MTBL_CUSTOMER_MASTER CM
        LEFT JOIN TBL_MACHINE_TRANSACTION MT
            ON CM.[CUS_CODE] = MT.[CUS_CODE]
        GROUP BY
             CM.[COM_ID], CM.[CUS_CODE], CM.[SAGE_CODE], CM.[CUS_NAME]
            ,CM.[ADD1], CM.[ADD2], CM.[ADD3], CM.[CONTACT_PERSON]
            ,CM.[MOBILE_NO], CM.[TEL_NO], CM.[FAX], CM.[EMAIL]
            ,CM.[CITY], CM.[POSTAL_CODE], CM.[CUS_STATUS]
            ,CM.[GPS_LOC1], CM.[GPS_LOC2], CM.[CUS_GRADE_ID], CM.[CUS_GRADE_NAME]
            ,CM.[REP_CODE], CM.[REP_NAME], CM.[TECH_CODE], CM.[TECH_NAME]
            ,CM.[DEBT_CODE], CM.[CONS_REP_CODE]
        ORDER BY CM.[CUS_CODE]";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<CustomerReportDto>(sql, buffered: false).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetCustomerReportData: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetCustomerReportData: {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // CUSTOMER REPORT: EXCEL
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadCustomerReportExcel()
        {
            var data = GetCustomerReportData();

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Customer Report");

            var headers = new[]
            {
        "COM ID", "Cus Code", "Sage Code", "Customer Name",
        "Add 1", "Add 2", "Add 3", "Contact Person",
        "Mobile No", "Tel No", "Fax", "Email",
        "City", "Postal Code", "Status",
        "GPS Loc 1", "GPS Loc 2",
        "Grade ID", "Grade Name",
        "Rep Code", "Rep Name",
        "Tech Code", "Tech Name",
        "Debt Code", "Cons Rep Code",
        "Machine Count"
    };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 112, 192);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.ComId ?? "";
                ws.Cell(row, c++).Value = r.CusCode ?? "";
                ws.Cell(row, c++).Value = r.SageCode ?? "";
                ws.Cell(row, c++).Value = r.CusName ?? "";
                ws.Cell(row, c++).Value = r.Add1 ?? "";
                ws.Cell(row, c++).Value = r.Add2 ?? "";
                ws.Cell(row, c++).Value = r.Add3 ?? "";
                ws.Cell(row, c++).Value = r.ContactPerson ?? "";
                ws.Cell(row, c++).Value = r.MobileNo ?? "";
                ws.Cell(row, c++).Value = r.TelNo ?? "";
                ws.Cell(row, c++).Value = r.Fax ?? "";
                ws.Cell(row, c++).Value = r.Email ?? "";
                ws.Cell(row, c++).Value = r.City ?? "";
                ws.Cell(row, c++).Value = r.PostalCode ?? "";
                ws.Cell(row, c++).Value = r.CusStatus ?? "";
                ws.Cell(row, c++).Value = r.GpsLoc1 ?? "";
                ws.Cell(row, c++).Value = r.GpsLoc2 ?? "";
                ws.Cell(row, c++).Value = r.CusGradeId ?? "";
                ws.Cell(row, c++).Value = r.CusGradeName ?? "";
                ws.Cell(row, c++).Value = r.RepCode ?? "";
                ws.Cell(row, c++).Value = r.RepName ?? "";
                ws.Cell(row, c++).Value = r.TechCode ?? "";
                ws.Cell(row, c++).Value = r.TechName ?? "";
                ws.Cell(row, c++).Value = r.DebtCode ?? "";
                ws.Cell(row, c++).Value = r.ConsRepCode ?? "";
                ws.Cell(row, c++).Value = r.MachineCount;

                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(235, 241, 250);

                row++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // CUSTOMER REPORT: PDF
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadCustomerReportPdf()
        {
            var data = GetCustomerReportData();

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Customer Report")
                                .Bold().FontSize(14).FontColor(Colors.Blue.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().LineHorizontal(1).LineColor(Colors.Blue.Darken3);
                    });

                    page.Content().PaddingTop(8).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(1.0f); // Cus Code
                            cols.RelativeColumn(2.5f); // Customer Name
                            cols.RelativeColumn(1.0f); // Tel No
                            cols.RelativeColumn(1.5f); // Email
                            cols.RelativeColumn(1.2f); // City
                            cols.RelativeColumn(0.8f); // Status
                            cols.RelativeColumn(1.0f); // Grade
                            cols.RelativeColumn(1.2f); // Rep Name
                            cols.RelativeColumn(1.2f); // Tech Name
                            cols.RelativeColumn(1.0f); // Machine Count
                        });

                        var colHeaders = new[]
                        {
                    "Cus Code", "Customer Name", "Tel No", "Email",
                    "City", "Status", "Grade", "Rep Name", "Tech Name", "Machines"
                };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Blue.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Blue.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3).Text(val ?? "").FontSize(7);

                            Cell(r.CusCode);
                            Cell(r.CusName);
                            Cell(r.TelNo);
                            Cell(r.Email);
                            Cell(r.City);
                            Cell(r.CusStatus);
                            Cell(r.CusGradeName);
                            Cell(r.RepName);
                            Cell(r.TechName);
                            Cell(r.MachineCount.ToString());
                        }
                    });

                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // CUSTOMER MACHINE LIST
        // ──────────────────────────────────────────────────────────────────────
        public List<FsReportDto> GetMachineList(string customerCode)
        {
            const string sql = @"
        SELECT
             [COM_ID]               AS ComId
            ,[SERIAL_NO]            AS SerialNo
            ,[DIVISION_ID]          AS DivisionId
            ,[AGREEMENT_ID]         AS AgreementId
            ,[CUS_CODE]             AS CusCode
            ,[CUS_NAME]             AS CusName
            ,[INV_ADD1]             AS InvAdd1
            ,[INV_ADD2]             AS InvAdd2
            ,[INV_ADD3]             AS InvAdd3
            ,[TEL_NO]               AS TelNo
            ,[FAX_NO]               AS FaxNo
            ,[CUS_GRADE]            AS CusGrade
            ,[MACHINE_REF_CODE]     AS MachineRefCode
            ,[MACHINE_CODE]         AS MachineCode
            ,[MACHINE_DESC]         AS MachineDesc
            ,[M_LOC1]               AS MLoc1
            ,[M_LOC2]               AS MLoc2
            ,[M_LOC3]               AS MLoc3
            ,[MACHINE_MODE]         AS MachineMode
            ,[M_LOC_CONTACT_NAME]   AS MLocContactName
            ,[M_LOC_CONTACT_NO]     AS MLocContactNo
            ,[MFP_CITY]             AS MfpCity
            ,[T_OFFICER_CODE]       AS TOfficerCode
            ,[T_OFFICER_NAME]       AS TOfficerName
            ,[M_INV_NO]             AS MInvNo
            ,[M_INV_DATE]           AS MInvDate
            ,[M_INS_DATE]           AS MInsDate
            ,[M_WARRANTY_END_DATE]  AS MWarrantyEndDate
            ,[LABOR_WARRANTY]       AS LaborWarranty
            ,[PARTS_WARRANTY]       AS PartsWarranty
            ,[IS_TRANSPORT]         AS IsTransport
            ,[PROVISION]            AS Provision
            ,[CUS_STATUS]           AS CusStatus
            ,[VISITS_PER_YEAR]      AS VisitsPerYear
            ,[INV_NO_TC]            AS InvNoTc
            ,[INV_DATE_TC]          AS InvDateTc
            ,[TEAM]                 AS Team
            ,[MA_PERIOD_START]      AS MaPeriodStart
            ,[MA_PERIOD_END]        AS MaPeriodEnd
            ,[SPECIAL]              AS Special
            ,[DATA_CHECKED]         AS DataChecked
            ,[REP_CODE]             AS RepCode
            ,[REP_NAME]             AS RepName
            ,[TO_CODE]              AS ToCode
            ,[TO_NAME]              AS ToName
            ,[TO_AREA_CODE]         AS ToAreaCode
            ,[TO_AREA_NAME]         AS ToAreaName
            ,[DEALER_ID]            AS DealerId
            ,[DEALER_NAME]          AS DealerName
            ,[CONS_REP_CODE]        AS ConsRepCode
            ,[CONS_REP_NAME]        AS ConsRepName
            ,[SERVICE_DATE]         AS ServiceDate
            ,[VISIT_CODE]           AS VisitCode
            ,[WARRANTY_TYPE_ID]     AS WarrantyTypeId
            ,[WARRANTY_TYPE_DESC]   AS WarrantyTypeDesc
            ,[MACHINE_TYPE]         AS MachineType
            ,[AD_REMORT]            AS AdRemort
            ,[CUS_TYPE]             AS CusType
            ,[SAGE_CODE]            AS SageCode
            ,[SUB_CUS_TYPE]         AS SubCusType
            ,[CONS_COMMENT]         AS ConsComment
            ,[CUS_EMAIL]            AS CusEmail
            ,[WARRANTY_YEARS]       AS WarrantyYears
            ,[CN_NO]                AS CnNo
            ,[M_LOC_TELEPHONE]      AS MLocTelephone
        FROM TBL_MACHINE_TRANSACTION
        WHERE [CUS_CODE] = @customerCode
        ORDER BY [SERIAL_NO]";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<FsReportDto>(sql, new { customerCode }, buffered: false).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetMachineList: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetMachineList: {ex.Message}");
                throw;
            }
        }



        // ──────────────────────────────────────────────────────────────────────
        // Monthly Service Visit REPORT: fetch data
        // ──────────────────────────────────────────────────────────────────────
        public List<MonthlyServiceReportDto> GetMonthlyServiceReportData(string startDate, string lastDate)
        {
            string sql = @"
                SELECT 
                T.T_ID,
                T.COM_ID,
                T.SERIAL_NO,
                T.CUS_ID,
                T.CUS_NAME,
                T.TECH_CODE,
                T.TECH_NAME,
                TC.AREA,                    -- Added from MTBL_TECH_OFFICERS
                T.MACHINE_REF,
                T.M_MODEL,
                V.VisitNo,
                V.ExpectedVisitDate,
                V.ServiceDate,
                V.MeterReading,
                CASE 
                    WHEN V.Status IS NULL THEN 'Pending'
                    WHEN LOWER(V.Status) = 'pending' THEN 'Pending'
                    WHEN LOWER(V.Status) = 'completed' THEN 'Completed'
                    WHEN LOWER(V.Status) = 'started' THEN 'Started'
                    ELSE V.Status
                END AS VisitStatus
            FROM TBL_SERVICE_SCEDULE_UPDATE T
            LEFT JOIN MTBL_TECH_OFFICERS TC   -- Joined on TECH_CODE
                ON T.TECH_CODE = TC.TECH_CODE
            CROSS APPLY
            (
                VALUES
                    (1, T.EXPT_SV1, T.SV1, T.SV1_STATUS, T.SV1_MR),
                    (2, T.EXPT_SV2, T.SV2, T.SV2_STATUS, T.SV2_MR),
                    (3, T.EXPT_SV3, T.SV3, T.SV3_STATUS, T.SV3_MR),
                    (4, T.EXPT_SV4, T.SV4, T.SV4_STATUS, T.SV4_MR),
                    (5, T.EXPT_SV5, T.SV5, T.SV5_STATUS, T.SV5_MR),
                    (6, T.EXPT_SV6, T.SV6, T.SV6_STATUS, T.SV6_MR)
            ) V (VisitNo, ExpectedVisitDate, ServiceDate, Status, MeterReading)
            WHERE 
                V.ExpectedVisitDate IS NOT NULL
                AND V.ExpectedVisitDate BETWEEN @startDate AND @lastDate
            ORDER BY 
                T.T_ID,
                V.VisitNo";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<MonthlyServiceReportDto>(sql, new { startDate, lastDate }, buffered: false).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetMonthly Service Visits Report: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetMonthly Service Visits Report : {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // Monthly Service Visit REPORT: EXCEL (NuGet: ClosedXML)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadMonthlyServiceReportExcel(string startDate, string lastDate)
        {
            var data = GetMonthlyServiceReportData(startDate, lastDate);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Service Visit Report");

            // ── Header row ──────────────────────────────────────────────────
            var headers = new[]
            {
        "T ID", "COM ID", "Serial No", "Cus ID", "Customer Name",
        "Tech Code", "Tech Name", "Machine Ref", "Model",
        "Visit No", "Expected Visit Date", "Service Date", "Meter Reading", "Visit Status"
    };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(31, 78, 121); // Dark teal/navy
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // ── Data rows ───────────────────────────────────────────────────
            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.T_ID;
                ws.Cell(row, c++).Value = r.COM_ID ?? "";
                ws.Cell(row, c++).Value = r.SERIAL_NO ?? "";
                ws.Cell(row, c++).Value = r.CUS_ID ?? "";
                ws.Cell(row, c++).Value = r.CUS_NAME ?? "";
                ws.Cell(row, c++).Value = r.TECH_CODE ?? "";
                ws.Cell(row, c++).Value = r.TECH_NAME ?? "";
                ws.Cell(row, c++).Value = r.MACHINE_REF ?? "";
                ws.Cell(row, c++).Value = r.M_MODEL ?? "";
                ws.Cell(row, c++).Value = r.VisitNo;
                ws.Cell(row, c++).Value = r.ExpectedVisitDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.ServiceDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, c++).Value = r.MeterReading ?? "";

                // Colour-code the status cell
                var statusCell = ws.Cell(row, c++);
                statusCell.Value = r.VisitStatus ?? "";
                statusCell.Style.Font.Bold = true;
                statusCell.Style.Font.FontColor = r.VisitStatus switch
                {
                    "Completed" => XLColor.FromArgb(0, 128, 0),   // Green
                    "Started" => XLColor.FromArgb(0, 70, 127),  // Blue
                    "Pending" => XLColor.FromArgb(192, 0, 0),   // Red
                    _ => XLColor.Black
                };

                // Zebra striping
                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(217, 225, 242);

                row++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // Monthly Service Visit REPORT: PDF (NuGet: QuestPDF)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadMonthlyServiceReportPdf(string startDate, string lastDate)
        {
            var data = GetMonthlyServiceReportData(startDate, lastDate);

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    // ── Header ──────────────────────────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Monthly Service Visit Report")
                                .Bold().FontSize(14).FontColor(Colors.Blue.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Blue.Darken3);
                    });

                    // ── Content ─────────────────────────────────────────────
                    page.Content().PaddingTop(8).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(0.7f); // Visit No
                            cols.RelativeColumn(1.0f); // Serial No
                            cols.RelativeColumn(2.2f); // Customer Name
                            cols.RelativeColumn(1.0f); // Tech Code
                            cols.RelativeColumn(1.5f); // Tech Name
                            cols.RelativeColumn(1.2f); // Machine Ref
                            cols.RelativeColumn(1.5f); // Model
                            cols.RelativeColumn(1.2f); // Expected Date
                            cols.RelativeColumn(1.2f); // Service Date
                            cols.RelativeColumn(1.0f); // Meter Reading
                            cols.RelativeColumn(1.0f); // Status
                        });

                        var colHeaders = new[]
                        {
                    "Visit No", "Serial No", "Customer", "Tech Code", "Tech Name",
                    "Machine Ref", "Model", "Expected Date", "Service Date",
                    "Meter Reading", "Status"
                };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Blue.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        // ── Data rows ───────────────────────────────────────
                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Blue.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(r.VisitNo.ToString());
                            Cell(r.SERIAL_NO);
                            Cell(r.CUS_NAME);
                            Cell(r.TECH_CODE);
                            Cell(r.TECH_NAME);
                            Cell(r.MACHINE_REF);
                            Cell(r.M_MODEL);
                            Cell(r.ExpectedVisitDate?.ToString("yyyy-MM-dd"));
                            Cell(r.ServiceDate?.ToString("yyyy-MM-dd"));
                            Cell(r.MeterReading);

                            // Status cell with colour
                            var statusColor = r.VisitStatus switch
                            {
                                "Completed" => Colors.Green.Darken2,
                                "Started" => Colors.Blue.Darken2,
                                "Pending" => Colors.Red.Darken2,
                                _ => Colors.Grey.Darken1
                            };
                            table.Cell().Background(bg).Padding(3)
                                .Text(r.VisitStatus ?? "").FontSize(7).FontColor(statusColor).Bold();
                        }
                    });

                    // ── Footer ───────────────────────────────────────────────
                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // Feedback REPORT: fetch data
        // ──────────────────────────────────────────────────────────────────────
        public List<FeedbackReportDto> GetFeedbackReportData(string startDate, string lastDate)
        {
            string sql = @"
            SELECT
                  [TID]
                , [COM_ID]
                , [FB_DATE]
                , [MOBILE_NO]
                , [CUS_CODE]
                , [CUS_NAME]
                , [RATING]
                , [FULL_MSG]
                , [JOB_ID]
                , [SERIAL_NO]
                , [MACHINE_REF_NO]
                , [CUSTOMER_REVIEW]
                , [TYPE]
                , [VISIT_NO]
                , [TECH_CODE]

                    FROM TBL_SV_CUSTOMER_JOB_FEEDBACKS

                    WHERE
                        CAST([FB_DATE] AS DATE)
                BETWEEN @StartDate AND @lastDate
            ORDER BY[FB_DATE] DESC";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();
                return connection.Query<FeedbackReportDto>(sql, new { startDate, lastDate }, buffered: false).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] Get Feedback Visits Report: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Get Feedback Visits Report : {ex.Message}");
                throw;
            }
        }

        // ──────────────────────────────────────────────────────────────────────
        // Feedback REPORT: EXCEL (NuGet: ClosedXML)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadFeedbackReportExcel(string startDate, string lastDate)
        {
            var data = GetFeedbackReportData(startDate, lastDate);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Feedback Report");

            // ── Header row ──────────────────────────────────────────────────
            var headers = new[]
            {
                "TID", "COM ID", "Feedback Date", "Mobile No",
                "Cus Code", "Customer Name", "Rating", "Full Message",
                "Job ID", "Serial No", "Machine Ref No", "Customer Review",
                "Type", "Visit No", "Tech Code"
            };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = ws.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(31, 78, 121); // Dark navy — matches DailyJob
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // ── Data rows ───────────────────────────────────────────────────
            int row = 2;
            foreach (var r in data)
            {
                int c = 1;
                ws.Cell(row, c++).Value = r.TID;
                ws.Cell(row, c++).Value = r.COM_ID ?? "";
                ws.Cell(row, c++).Value = r.FB_DATE.ToString("yyyy-MM-dd HH:mm");
                ws.Cell(row, c++).Value = r.MOBILE_NO ?? "";
                ws.Cell(row, c++).Value = r.CUS_CODE ?? "";
                ws.Cell(row, c++).Value = r.CUS_NAME ?? "";

                // Rating cell — colour-coded 1–5
                var ratingCell = ws.Cell(row, c++);
                ratingCell.Value = r.RATING;
                ratingCell.Style.Font.Bold = true;
                ratingCell.Style.Font.FontColor = r.RATING switch
                {
                    5 => XLColor.FromArgb(0, 128, 0),    // Green
                    4 => XLColor.FromArgb(0, 176, 80),   // Light green
                    3 => XLColor.FromArgb(255, 165, 0),  // Orange
                    2 => XLColor.FromArgb(192, 80, 77),  // Salmon red
                    1 => XLColor.FromArgb(192, 0, 0),    // Dark red
                    _ => XLColor.Black
                };

                ws.Cell(row, c++).Value = r.FULL_MSG ?? "";
                ws.Cell(row, c++).Value = r.JOB_ID?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.SERIAL_NO ?? "";
                ws.Cell(row, c++).Value = r.MACHINE_REF_NO ?? "";
                ws.Cell(row, c++).Value = r.CUSTOMER_REVIEW ?? "";
                ws.Cell(row, c++).Value = r.TYPE ?? "";
                ws.Cell(row, c++).Value = r.VISIT_NO?.ToString() ?? "";
                ws.Cell(row, c++).Value = r.TECH_CODE ?? "";

                // Zebra striping
                if (row % 2 == 0)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.FromArgb(217, 225, 242);

                row++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        // ──────────────────────────────────────────────────────────────────────
        // Feedback REPORT: PDF (NuGet: QuestPDF)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadFeedbackReportPdf(string startDate, string lastDate)
        {
            var data = GetFeedbackReportData(startDate, lastDate);

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(7).FontFamily("Arial"));

                    // ── Header ──────────────────────────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Customer Feedback Report")
                                .Bold().FontSize(14).FontColor(Colors.Blue.Darken3);
                            row.ConstantItem(200).AlignRight()
                                .Text($"Period: {startDate}  →  {lastDate}")
                                .FontSize(8).FontColor(Colors.Grey.Darken1);
                        });
                        col.Item().PaddingBottom(4)
                            .Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}")
                            .FontSize(7).FontColor(Colors.Grey.Medium);
                        col.Item().LineHorizontal(1).LineColor(Colors.Blue.Darken3);
                    });

                    // ── Content ─────────────────────────────────────────────
                    page.Content().PaddingTop(8).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(0.6f); // TID
                            cols.RelativeColumn(1.2f); // FB Date
                            cols.RelativeColumn(1.0f); // Mobile No
                            cols.RelativeColumn(2.0f); // Customer Name
                            cols.RelativeColumn(0.5f); // Rating
                            cols.RelativeColumn(2.5f); // Full Message
                            cols.RelativeColumn(1.2f); // Serial No
                            cols.RelativeColumn(1.2f); // Machine Ref
                            cols.RelativeColumn(0.8f); // Type
                            cols.RelativeColumn(0.8f); // Tech Code
                        });

                        var colHeaders = new[]
                        {
                            "TID", "Feedback Date", "Mobile No", "Customer",
                            "Rating", "Message", "Serial No", "Machine Ref",
                            "Type", "Tech Code"
                        };

                        table.Header(header =>
                        {
                            foreach (var h in colHeaders)
                            {
                                header.Cell().Background(Colors.Blue.Darken3)
                                    .Padding(3)
                                    .Text(h)
                                    .Bold().FontColor(Colors.White).FontSize(7);
                            }
                        });

                        // ── Data rows ───────────────────────────────────────
                        bool isEven = false;
                        foreach (var r in data)
                        {
                            var bg = isEven ? Colors.Blue.Lighten5 : Colors.White;
                            isEven = !isEven;

                            void Cell(string? val) =>
                                table.Cell().Background(bg).Padding(3)
                                    .Text(val ?? "").FontSize(7);

                            Cell(r.TID.ToString());
                            Cell(r.FB_DATE.ToString("yyyy-MM-dd HH:mm"));
                            Cell(r.MOBILE_NO);
                            Cell(r.CUS_NAME);

                            // Rating — colour-coded
                            var ratingColor = r.RATING switch
                            {
                                5 => Colors.Green.Darken2,
                                4 => Colors.Green.Lighten1,
                                3 => Colors.Orange.Darken1,
                                2 => Colors.Red.Lighten1,
                                1 => Colors.Red.Darken3,
                                _ => Colors.Grey.Darken1
                            };
                            table.Cell().Background(bg).Padding(3)
                                .Text(r.RATING.ToString())
                                .FontSize(7).FontColor(ratingColor).Bold();

                            Cell(r.FULL_MSG);
                            Cell(r.SERIAL_NO);
                            Cell(r.MACHINE_REF_NO);
                            Cell(r.TYPE);
                            Cell(r.TECH_CODE);
                        }
                    });

                    // ── Footer ───────────────────────────────────────────────
                    page.Footer().AlignCenter()
                        .Text(x =>
                        {
                            x.Span("Page ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.CurrentPageNumber().FontSize(7).FontColor(Colors.Grey.Medium);
                            x.Span(" of ").FontSize(7).FontColor(Colors.Grey.Medium);
                            x.TotalPages().FontSize(7).FontColor(Colors.Grey.Medium);
                        });
                });
            });

            return document.GeneratePdf();
        }

        // ──────────────────────────────────────────────────────────────────────
        // MONTHLY SERVICE SCHEDULE: fetch data
        // ──────────────────────────────────────────────────────────────────────
        public List<MonthlyServiceScheduleDto> GetMonthlyServiceScheduleData(
            string startDate, string lastDate, string techCode, string machineRefStartAt)
        {
            DateTime fromDate = DateTime.Parse(startDate);
            DateTime toDate = DateTime.Parse(lastDate);

            // Normalize: "all" (any case) or blank means no prefix filter
            bool filterByPrefix = !string.IsNullOrWhiteSpace(machineRefStartAt)
                                   && !machineRefStartAt.Equals("all", StringComparison.OrdinalIgnoreCase);

            string sql = @"
                SELECT 
                    T.T_ID,
                    T.COM_ID,
                    T.SERIAL_NO,
                    T.CUS_ID AS Customer_Code,
                    T.CUS_NAME AS Customer_Name,
                    T.TECH_CODE,
                    T.TECH_NAME,
                    TC.AREA,
                    T.MACHINE_REF,
                    T.M_MODEL,

                    M.CUS_STATUS AS agrementType,
                    M.TEL_NO,

                    -- Machine Location from TBL_MACHINE_TRANSACTION
                    M.INV_ADD1 AS M_LOC1,
                    M.INV_ADD2 AS M_LOC2,
                    M.INV_ADD3 AS M_LOC3,

                    COALESCE(
                        M.MA_PERIOD_START,
                        M.M_INS_DATE
                    ) AS machine_instalation_or_period_start_date,

                    V.VisitNo,
                    V.ExpectedVisitDate,
                    V.ServiceDate,
                    V.MeterReading,

                    CASE 
                        WHEN V.Status IS NULL THEN 'Pending'
                        WHEN LOWER(V.Status) = 'pending' THEN 'Pending'
                        WHEN LOWER(V.Status) = 'completed' THEN 'Completed'
                        WHEN LOWER(V.Status) = 'started' THEN 'Started'
                        ELSE V.Status
                    END AS VisitStatus

                FROM [DB_A67CC4_Servvistagcp].[dbo].[TBL_SERVICE_SCEDULE_UPDATE] T

                LEFT JOIN [DB_A67CC4_Servvistagcp].[dbo].[MTBL_TECH_OFFICERS] TC 
                    ON T.TECH_CODE = TC.TECH_CODE

                LEFT JOIN [DB_A67CC4_Servvistagcp].[dbo].[TBL_MACHINE_TRANSACTION] M 
                    ON T.SERIAL_NO = M.SERIAL_NO

                CROSS APPLY
                (
                    VALUES
                        (1, T.EXPT_SV1, T.SV1, T.SV1_STATUS, T.SV1_MR),
                        (2, T.EXPT_SV2, T.SV2, T.SV2_STATUS, T.SV2_MR),
                        (3, T.EXPT_SV3, T.SV3, T.SV3_STATUS, T.SV3_MR),
                        (4, T.EXPT_SV4, T.SV4, T.SV4_STATUS, T.SV4_MR),
                        (5, T.EXPT_SV5, T.SV5, T.SV5_STATUS, T.SV5_MR),
                        (6, T.EXPT_SV6, T.SV6, T.SV6_STATUS, T.SV6_MR)
                ) V (
                    VisitNo,
                    ExpectedVisitDate,
                    ServiceDate,
                    Status,
                    MeterReading
                )

                WHERE 
                    V.ExpectedVisitDate IS NOT NULL
                    AND V.ExpectedVisitDate >= @fromDate
                    AND V.ExpectedVisitDate < @toDate
                    AND T.TECH_CODE = @techCode
                    AND M.CUS_STATUS <> 'NS'";

            if (filterByPrefix)
            {
                sql += " AND T.MACHINE_REF LIKE @machineRefPattern";
            }

            sql += @"
            ORDER BY
                V.ExpectedVisitDate ASC,
                T.T_ID,
                V.VisitNo";

            try
            {
                using var connection = new SqlConnection(_connectionString);
                connection.Open();

                var parameters = new DynamicParameters();
                parameters.Add("fromDate", fromDate);
                parameters.Add("toDate", toDate);
                parameters.Add("techCode", techCode);

                if (filterByPrefix)
                {
                    // Escape LIKE wildcards in the character just in case, then anchor prefix match
                    string escaped = machineRefStartAt.Replace("[", "[[]").Replace("%", "[%]").Replace("_", "[_]");
                    parameters.Add("machineRefPattern", escaped + "%");
                }

                return connection.Query<MonthlyServiceScheduleDto>(
                    sql, parameters, buffered: false).ToList();
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"[SQL ERROR] GetMonthlyServiceScheduleData: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] GetMonthlyServiceScheduleData: {ex.Message}");
                throw;
            }
        }


        // ──────────────────────────────────────────────────────────────────────
        // MONTHLY SERVICE SCHEDULE: PDF (NuGet: QuestPDF)
        // ──────────────────────────────────────────────────────────────────────
        public byte[] DownloadMonthlyServiceSchedulePdf(
    string startDate, string lastDate, string techCode, string machineRefStartAt)
        {
            var data = GetMonthlyServiceScheduleData(startDate, lastDate, techCode, machineRefStartAt);

            DateTime fromDate = DateTime.Parse(startDate);
            string periodLabel = fromDate.ToString("MMMM yyyy");
            string techName = data.FirstOrDefault()?.TECH_NAME ?? "";

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(20);
                    page.DefaultTextStyle(x => x.FontSize(9).FontFamily("Arial"));

                    // ── Header (repeats on every page) ─────────────────────────
                    page.Header().Column(col =>
                    {
                        col.Item().Row(row =>
                        {
                            row.RelativeItem().AlignCenter()
                                .Text($"Service Schedule Report - {periodLabel}")
                                .Bold().FontSize(18);
                        });

                        col.Item().Row(row =>
                        {
                            row.RelativeItem().Text("");
                            row.ConstantItem(160).AlignRight()
                                .Text($"{DateTime.Now:dd-MMMM-yyyy}").FontSize(9);
                        });

                        col.Item().PaddingTop(5).Border(1).BorderColor(Colors.Black)
                            .Padding(8).Row(row =>
                            {
                                row.RelativeItem().Column(c =>
                                {
                                    c.Item().Text("Technician Code").Bold().FontSize(9);
                                    c.Item().Text(techCode).FontSize(10);
                                });
                                row.RelativeItem().Column(c =>
                                {
                                    c.Item().AlignCenter().Text("Technician Name").Bold().FontSize(9);
                                    c.Item().AlignCenter().Text(techName).FontSize(10);
                                });
                                row.RelativeItem().Column(c =>
                                {
                                    c.Item().AlignRight().Text("Time Period").Bold().FontSize(9);
                                    c.Item().AlignRight().Text(periodLabel).FontSize(10);
                                });
                            });
                    });

                    // ── Content: one bordered block per visit ──────────────────
                    page.Content().PaddingTop(10).Column(col =>
                    {
                        // Summary: group all scheduled machines by agreement type.
                        col.Item()
                            .BorderBottom(2).BorderColor(Colors.Blue.Darken2)
                            .PaddingBottom(6).Row(row =>
                            {
                                row.RelativeItem().Text("SERVICE SCHEDULE SUMMARY")
                                    .Bold().FontSize(13).FontColor(Colors.Blue.Darken2);
                                row.ConstantItem(110).AlignRight().Text($"Total: {data.Count}")
                                    .SemiBold().FontSize(8.5f).FontColor(Colors.Grey.Darken2);
                            });

                        col.Item().PaddingTop(6).PaddingBottom(18).Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.ConstantColumn(32);       // Index
                                columns.RelativeColumn(2.2f);    // Customer name
                                columns.RelativeColumn(0.9f);    // Q number
                                columns.RelativeColumn(1.2f);    // Serial number
                                columns.RelativeColumn(2.7f);    // Address
                            });

                            table.Header(header =>
                            {
                                void HeaderCell(string text, bool center = false)
                                {
                                    var cell = header.Cell()
                                        .Background(Colors.Blue.Darken2)
                                        .BorderRight(0.5f).BorderColor(Colors.Blue.Lighten1)
                                        .PaddingVertical(6).PaddingHorizontal(5);

                                    if (center)
                                        cell.AlignCenter().Text(text).Bold().FontSize(7.5f).FontColor(Colors.White);
                                    else
                                        cell.Text(text).Bold().FontSize(7.5f).FontColor(Colors.White);
                                }

                                HeaderCell("NO.", true);
                                HeaderCell("Customer Name");
                                HeaderCell("Q Number", true);
                                HeaderCell("Serial Number", true);
                                HeaderCell("Address");
                            });

                            int summaryIndex = 1;
                            var agreementGroups = data
                                .GroupBy(record => string.IsNullOrWhiteSpace(record.agrementType)
                                    ? "Not Specified"
                                    : record.agrementType.Trim());

                            foreach (var agreementGroup in agreementGroups)
                            {
                                table.Cell().ColumnSpan(5)
                                    .Background(Colors.Blue.Lighten4)
                                    .BorderBottom(1).BorderColor(Colors.Blue.Lighten2)
                                    .PaddingVertical(5).PaddingHorizontal(7)
                                    .Row(row =>
                                    {
                                        row.RelativeItem().Text(agreementGroup.Key.ToUpperInvariant())
                                            .Bold().FontSize(8.5f).FontColor(Colors.Blue.Darken3);
                                        row.ConstantItem(75).AlignRight()
                                            .Text($"{agreementGroup.Count()} record(s)")
                                            .SemiBold().FontSize(7).FontColor(Colors.Blue.Darken2);
                                    });

                                int rowIndex = 0;
                                foreach (var record in agreementGroup)
                                {
                                    var address = string.Join(", ", new[]
                                    {
                                        record.M_LOC1,
                                        record.M_LOC2,
                                        record.M_LOC3
                                    }.Where(s => !string.IsNullOrWhiteSpace(s)));

                                    var rowBackground = rowIndex++ % 2 == 0
                                        ? Colors.White
                                        : Colors.Grey.Lighten4;

                                    IContainer SummaryCell() => table.Cell()
                                        .Background(rowBackground)
                                        .BorderBottom(0.5f).BorderColor(Colors.Grey.Lighten2)
                                        .PaddingVertical(5).PaddingHorizontal(5)
                                        .MinHeight(22).AlignMiddle();

                                    SummaryCell().AlignCenter()
                                        .Text(summaryIndex++.ToString()).SemiBold().FontSize(7.5f)
                                        .FontColor(Colors.Grey.Darken2);
                                    SummaryCell().Text(record.Customer_Name ?? "").SemiBold().FontSize(7.5f);
                                    SummaryCell().AlignCenter().Text(record.MACHINE_REF ?? "").FontSize(7.5f);
                                    SummaryCell().AlignCenter().Text(record.SERIAL_NO ?? "").FontSize(7.5f);
                                    SummaryCell().Text(address).FontSize(7.5f).FontColor(Colors.Grey.Darken3);
                                }
                            }
                        });

                        int counter = 1;
                        foreach (var r in data)
                        {
                            var record = r;
                            var idx = counter++;

                            col.Item().PaddingBottom(8).Border(1).BorderColor(Colors.Black)
                                .Padding(8).Row(row =>
                                {
                                    row.ConstantItem(22).AlignTop().Text(idx.ToString())
                                .Bold().FontSize(11);

                                    // Left: auto-filled fields
                                    row.RelativeItem(1.5f).PaddingRight(10).Column(left =>
                                    {
                                        void F(string label, string? val) => left.Item().PaddingBottom(2).Row(fr =>
                                        {
                                            fr.ConstantItem(160).Text(label).FontSize(8.5f).SemiBold();
                                            fr.ConstantItem(10).Text(":").FontSize(8.5f);
                                            fr.RelativeItem().Text(val ?? "").FontSize(8.5f);
                                        });

                                        var address = string.Join(", ", new[] { record.M_LOC1, record.M_LOC2, record.M_LOC3 }
                                    .Where(s => !string.IsNullOrWhiteSpace(s)));

                                        F("Customer Code", record.Customer_Code);
                                        F("Customer / Company Name", record.Customer_Name);
                                        F("Address", address);
                                        F("Telephone", record.TEL_NO);
                                        F("Machine Model", record.M_MODEL);
                                        F("Serial No.", record.SERIAL_NO);
                                        F("Q.No.", record.MACHINE_REF);
                                        F("Machine Install or Agreement Date",
                                    record.machine_instalation_or_period_start_date?.ToString("yyyy-MM-dd"));
                                        F("Maintenance Service Date (Expected Visit Date)",
                                    record.ExpectedVisitDate?.ToString("yyyy-MM-dd"));
                                        F("Visit No", record.VisitNo.ToString());
                                        F("Service Type / Status", record.agrementType);
                                    });

                                    row.ConstantItem(1).Background(Colors.Grey.Lighten1);

                                    // Right: manual fields (blank for filling in by hand)
                                    row.RelativeItem(1.2f).PaddingLeft(10).Column(right =>
                                    {
                                        void M(string label) => right.Item().PaddingBottom(6).Row(fr =>
                                        {
                                            fr.ConstantItem(120).Text(label).FontSize(8.5f).SemiBold();
                                            fr.ConstantItem(10).Text(":").FontSize(8.5f);
                                            fr.RelativeItem().Text("________________").FontSize(8.5f);
                                        });

                                        M("Service Date");
                                        M("Counter Reading");
                                        M("Meter Reading");
                                        right.Item().PaddingBottom(4).LineHorizontal(0.5f).LineColor(Colors.Grey.Lighten1);
                                        M("Customer Name");
                                        M("Customer Mobile No.");
                                        M("Rubber Stamp");
                                        right.Item().PaddingBottom(4).LineHorizontal(0.5f).LineColor(Colors.Grey.Lighten1);
                                        M("Customer Signature");
                                        M("Customer Designation");
                                    });
                                });
                        }

                    });

                    page.Footer().AlignCenter().Text(x =>
                    {
                        x.Span("Page ").FontSize(8);
                        x.CurrentPageNumber().FontSize(8);
                        x.Span(" of ").FontSize(8);
                        x.TotalPages().FontSize(8);
                    });
                });
            });

            return document.GeneratePdf();
        }
    }
}
