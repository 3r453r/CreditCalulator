using ClosedXML.Excel;
using CreditTool.Models;

namespace CreditTool.Services;

public class ExcelService
{
    public (CreditParameters Parameters, List<InterestRatePeriod> Rates) Import(Stream fileStream)
    {
        using var workbook = new XLWorkbook(fileStream);
        var parameters = ReadParameters(workbook.Worksheet(1));
        var rates = ReadRates(workbook.Worksheet(2));
        return (parameters, rates);
    }

    public byte[] Export(CreditParameters parameters, IEnumerable<InterestRatePeriod> rates, IEnumerable<ScheduleItem> schedule)
    {
        using var workbook = new XLWorkbook();
        var parameterSheet = workbook.AddWorksheet("Parameters");
        WriteParameters(parameterSheet, parameters);

        var rateSheet = workbook.AddWorksheet("Rates");
        WriteRates(rateSheet, rates);

        var scheduleSheet = workbook.AddWorksheet("Schedule");
        WriteSchedule(scheduleSheet, schedule);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static CreditParameters ReadParameters(IXLWorksheet worksheet)
    {
        var parameterMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var currentRow = 2;

        while (!worksheet.Cell(currentRow, 1).IsEmpty())
        {
            var key = worksheet.Cell(currentRow, 1).GetString();
            var value = worksheet.Cell(currentRow, 2).GetString();
            if (!string.IsNullOrWhiteSpace(key))
            {
                parameterMap[key.Trim()] = value;
            }

            currentRow++;
        }

        return new CreditParameters
        {
            NetValue = ParseDecimal(parameterMap, "NetValue"),
            MarginRate = ParseDecimal(parameterMap, "MarginRate"),
            PaymentFrequency = ParseEnum(parameterMap, "PaymentFrequency", PaymentFrequency.Monthly),
            PaymentDay = ParseEnum(parameterMap, "PaymentDay", PaymentDayOption.LastOfMonth),
            CreditStartDate = ParseDate(parameterMap, "CreditStartDate"),
            CreditEndDate = ParseDate(parameterMap, "CreditEndDate"),
            DayCountBasis = ParseEnum(parameterMap, "DayCountBasis", DayCountBasis.Actual365),
            RoundingMode = ParseEnum(parameterMap, "RoundingMode", RoundingModeOption.Bankers),
            RoundingDecimals = ParseInt(parameterMap, "RoundingDecimals", 2),
            ProcessingFeeRate = ParseDecimal(parameterMap, "ProcessingFeeRate"),
            BulletRepayment = ParseBool(parameterMap, "BulletRepayment")
        };
    }

    private static void WriteParameters(IXLWorksheet sheet, CreditParameters parameters)
    {
        sheet.Cell(1, 1).Value = "Parameter";
        sheet.Cell(1, 2).Value = "Value";

        var values = new (string, object?)[]
        {
            ("NetValue", parameters.NetValue),
            ("MarginRate", parameters.MarginRate),
            ("PaymentFrequency", parameters.PaymentFrequency),
            ("PaymentDay", parameters.PaymentDay),
            ("CreditStartDate", parameters.CreditStartDate),
            ("CreditEndDate", parameters.CreditEndDate),
            ("DayCountBasis", parameters.DayCountBasis),
            ("RoundingMode", parameters.RoundingMode),
            ("RoundingDecimals", parameters.RoundingDecimals),
            ("ProcessingFeeRate", parameters.ProcessingFeeRate),
            ("BulletRepayment", parameters.BulletRepayment)
        };

        var row = 2;
        foreach (var (key, value) in values)
        {
            sheet.Cell(row, 1).Value = key;
            sheet.Cell(row, 2).SetValue(value?.ToString() ?? string.Empty);
            row++;
        }

        sheet.Columns().AdjustToContents();
    }

    private static List<InterestRatePeriod> ReadRates(IXLWorksheet worksheet)
    {
        var rows = new List<InterestRatePeriod>();
        var currentRow = 2;

        while (!worksheet.Cell(currentRow, 1).IsEmpty())
        {
            rows.Add(new InterestRatePeriod
            {
                DateFrom = worksheet.Cell(currentRow, 1).GetDateTime(),
                DateTo = worksheet.Cell(currentRow, 2).GetDateTime(),
                Rate = worksheet.Cell(currentRow, 3).GetValue<decimal>()
            });

            currentRow++;
        }

        return rows;
    }

    private static void WriteRates(IXLWorksheet worksheet, IEnumerable<InterestRatePeriod> rates)
    {
        worksheet.Cell(1, 1).Value = "DateFrom";
        worksheet.Cell(1, 2).Value = "DateTo";
        worksheet.Cell(1, 3).Value = "Rate";

        var row = 2;
        foreach (var rate in rates)
        {
            worksheet.Cell(row, 1).Value = rate.DateFrom;
            worksheet.Cell(row, 2).Value = rate.DateTo;
            worksheet.Cell(row, 3).Value = rate.Rate;
            row++;
        }

        worksheet.Columns().AdjustToContents();
    }

    private static void WriteSchedule(IXLWorksheet worksheet, IEnumerable<ScheduleItem> schedule)
    {
        worksheet.Cell(1, 1).Value = "PaymentDate";
        worksheet.Cell(1, 2).Value = "DaysInPeriod";
        worksheet.Cell(1, 3).Value = "InterestRate";
        worksheet.Cell(1, 4).Value = "InterestAmount";
        worksheet.Cell(1, 5).Value = "PrincipalPayment";
        worksheet.Cell(1, 6).Value = "TotalPayment";
        worksheet.Cell(1, 7).Value = "RemainingPrincipal";

        var row = 2;
        foreach (var item in schedule)
        {
            worksheet.Cell(row, 1).SetValue(item.PaymentDate);
            worksheet.Cell(row, 2).SetValue(item.DaysInPeriod);
            worksheet.Cell(row, 3).SetValue(item.InterestRate);
            worksheet.Cell(row, 4).SetValue(item.InterestAmount);
            worksheet.Cell(row, 5).SetValue(item.PrincipalPayment);
            worksheet.Cell(row, 6).SetValue(item.TotalPayment);
            worksheet.Cell(row, 7).SetValue(item.RemainingPrincipal);
            row++;
        }

        worksheet.Columns().AdjustToContents();
    }

    private static decimal ParseDecimal(IDictionary<string, string> parameters, string key)
    {
        return parameters.TryGetValue(key, out var value) && decimal.TryParse(value, out var result) ? result : 0m;
    }

    private static int ParseInt(IDictionary<string, string> parameters, string key, int defaultValue)
    {
        return parameters.TryGetValue(key, out var value) && int.TryParse(value, out var result) ? result : defaultValue;
    }

    private static DateTime ParseDate(IDictionary<string, string> parameters, string key)
    {
        return parameters.TryGetValue(key, out var value) && DateTime.TryParse(value, out var result)
            ? result
            : DateTime.Today;
    }

    private static TEnum ParseEnum<TEnum>(IDictionary<string, string> parameters, string key, TEnum defaultValue) where TEnum : struct
    {
        return parameters.TryGetValue(key, out var value) && Enum.TryParse<TEnum>(value, out var parsed)
            ? parsed
            : defaultValue;
    }

    private static bool ParseBool(IDictionary<string, string> parameters, string key)
    {
        return parameters.TryGetValue(key, out var value) && bool.TryParse(value, out var result) && result;
    }
}
