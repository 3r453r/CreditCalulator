using ClosedXML.Excel;
using CreditTool.Models;
using System.Globalization;

namespace CreditTool.Services;

public class ExcelService
{
    private static readonly string[] DayFirstDateFormats =
    {
        "dd/MM/yyyy",
        "d/M/yyyy",
        "dd.MM.yyyy",
        "d.M.yyyy",
        "dd-MM-yyyy",
        "d-M-yyyy"
    };

    private static readonly CultureInfo[] DayFirstCultures =
    {
        CultureInfo.InvariantCulture,
        CultureInfo.GetCultureInfo("pl-PL"),
        CultureInfo.GetCultureInfo("en-GB")
    };

    internal static readonly Dictionary<string, string> ParameterKeyAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Kwota netto"] = "NetValue",
        ["Marża"] = "MarginRate",
        ["Częstotliwość płatności"] = "PaymentFrequency",
        ["Dzień płatności"] = "PaymentDay",
        ["Data początkowa"] = "CreditStartDate",
        ["Data końcowa"] = "CreditEndDate",
        ["Konwencja dni"] = "DayCountBasis",
        ["Zaokrąglanie"] = "RoundingMode",
        ["Miejsca po przecinku"] = "RoundingDecimals",
        ["Prowizja przygotowawcza"] = "ProcessingFeeRate",
        ["Prowizja przygotowawcza (kwota)"] = "ProcessingFeeAmount",
        ["Spłata balonowa"] = "BulletRepayment",
        ["Typ spłaty"] = "PaymentType"
    };

    public (CreditParameters Parameters, List<InterestRatePeriod> Rates) Import(Stream fileStream)
    {
        using var workbook = new XLWorkbook(fileStream);
        var parameters = ReadParameters(workbook.Worksheet(1));
        var rates = ReadRates(workbook.Worksheet(2));

        if (rates.Count == 0)
        {
            throw new InvalidOperationException("Brak stóp procentowych w arkuszu importu.");
        }

        return (parameters, rates);
    }

    public byte[] Export(CreditParameters parameters, IEnumerable<InterestRatePeriod> rates, IEnumerable<ScheduleItem> schedule, decimal totalInterest)
    {
        using var workbook = new XLWorkbook();
        var parameterSheet = workbook.AddWorksheet("Parametry");
        WriteParameters(parameterSheet, parameters);

        var rateSheet = workbook.AddWorksheet("Stopy procentowe");
        WriteRates(rateSheet, rates);

        var scheduleSheet = workbook.AddWorksheet("Harmonogram");
        var apr = AprCalculator.CalculateAnnualPercentageRate(parameters, schedule);
        WriteSchedule(scheduleSheet, schedule, totalInterest, apr);

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
                var normalizedKey = NormalizeParameterKey(key);
                parameterMap[normalizedKey] = value;
            }

            currentRow++;
        }

        ValidateRequiredParameters(parameterMap);

        return new CreditParameters
        {
            NetValue = ParseDecimal(parameterMap, "NetValue", required: true),
            MarginRate = ParseDecimal(parameterMap, "MarginRate", required: true),
            PaymentFrequency = ParseEnum(parameterMap, "PaymentFrequency", PaymentFrequency.Monthly, required: true),
            PaymentDay = ParseEnum(parameterMap, "PaymentDay", PaymentDayOption.LastOfMonth, required: true),
            CreditStartDate = ParseDate(parameterMap, "CreditStartDate", required: true),
            CreditEndDate = ParseDate(parameterMap, "CreditEndDate", required: true),
            DayCountBasis = ParseEnum(parameterMap, "DayCountBasis", DayCountBasis.Actual365, required: true),
            RoundingMode = ParseEnum(parameterMap, "RoundingMode", RoundingModeOption.Bankers, required: true),
            RoundingDecimals = ParseInt(parameterMap, "RoundingDecimals", 4, required: true),
            ProcessingFeeRate = ParseDecimal(parameterMap, "ProcessingFeeRate"),
            ProcessingFeeAmount = ParseDecimal(parameterMap, "ProcessingFeeAmount"),
            BulletRepayment = ParseBool(parameterMap, "BulletRepayment"),
            PaymentType = ParseEnum(parameterMap, "PaymentType", PaymentType.DecreasingInstallments)
        };
    }

    private static void WriteParameters(IXLWorksheet sheet, CreditParameters parameters)
    {
        sheet.Cell(1, 1).Value = "Parametr";
        sheet.Cell(1, 2).Value = "Wartość";

        var values = new (string Key, string Label, object? Value)[]
        {
            ("NetValue", "Kwota netto", parameters.NetValue),
            ("MarginRate", "Marża", parameters.MarginRate),
            ("PaymentFrequency", "Częstotliwość płatności", parameters.PaymentFrequency),
            ("PaymentDay", "Dzień płatności", parameters.PaymentDay),
            ("CreditStartDate", "Data początkowa", parameters.CreditStartDate),
            ("CreditEndDate", "Data końcowa", parameters.CreditEndDate),
            ("DayCountBasis", "Konwencja dni", parameters.DayCountBasis),
            ("RoundingMode", "Zaokrąglanie", parameters.RoundingMode),
            ("RoundingDecimals", "Miejsca po przecinku", parameters.RoundingDecimals),
            ("ProcessingFeeRate", "Prowizja przygotowawcza", parameters.ProcessingFeeRate),
            ("ProcessingFeeAmount", "Prowizja przygotowawcza (kwota)", parameters.ProcessingFeeAmount),
            ("PaymentType", "Typ spłaty", parameters.PaymentType),
            ("BulletRepayment", "Spłata balonowa", parameters.BulletRepayment)
        };

        var choiceOptions = new Dictionary<string, string[]>
        {
            ["PaymentFrequency"] = Enum.GetNames<PaymentFrequency>(),
            ["PaymentDay"] = Enum.GetNames<PaymentDayOption>(),
            ["DayCountBasis"] = Enum.GetNames<DayCountBasis>(),
            ["RoundingMode"] = Enum.GetNames<RoundingModeOption>(),
            ["PaymentType"] = Enum.GetNames<PaymentType>(),
            ["BulletRepayment"] = new[] { bool.TrueString, bool.FalseString }
        };

        var row = 2;
        foreach (var (key, label, value) in values)
        {
            sheet.Cell(row, 1).Value = label;
            var valueCell = sheet.Cell(row, 2);
            valueCell.SetValue(value?.ToString() ?? string.Empty);

            if (choiceOptions.TryGetValue(key, out var options))
            {
                var validation = valueCell.CreateDataValidation();
                validation.List(string.Join(',', options));
                validation.InCellDropdown = true;
            }

            row++;
        }

        var legendStartRow = row + 1;
        sheet.Cell(legendStartRow, 1).Value = "Legenda pól wyboru";
        sheet.Cell(legendStartRow, 2).Value = "Dostępne wartości";

        var legendRow = legendStartRow + 1;
        foreach (var (key, options) in choiceOptions)
        {
            var label = values.First(v => v.Key == key).Label;
            sheet.Cell(legendRow, 1).Value = label;
            sheet.Cell(legendRow, 2).Value = string.Join(", ", options);
            legendRow++;
        }

        sheet.Columns().AdjustToContents();
    }

    private static List<InterestRatePeriod> ReadRates(IXLWorksheet worksheet)
    {
        var rows = new List<InterestRatePeriod>();
        var currentRow = 2;

        while (!worksheet.Cell(currentRow, 1).IsEmpty())
        {
            var fromCell = worksheet.Cell(currentRow, 1);
            var toCell = worksheet.Cell(currentRow, 2);
            var rateCell = worksheet.Cell(currentRow, 3);

            if (fromCell.IsEmpty() || toCell.IsEmpty() ||
                !TryGetDate(fromCell, out var from) ||
                !TryGetDate(toCell, out var to))
            {
                throw new InvalidOperationException($"Nieprawidłowa lub brakująca data w wierszu {currentRow} tabeli stóp procentowych.");
            }

            if (TryParseDecimal(rateCell.GetString(), out var rate))
            {
                rows.Add(new InterestRatePeriod
                {
                    DateFrom = from,
                    DateTo = to,
                    Rate = rate
                });
            }

            currentRow++;
        }

        return rows;
    }

    private static bool TryParseDecimal(string value, out decimal result)
    {
        foreach (var culture in new[] { CultureInfo.InvariantCulture, CultureInfo.GetCultureInfo("pl-PL"), CultureInfo.CurrentCulture })
        {
            if (decimal.TryParse(value, NumberStyles.Any, culture, out result))
            {
                return true;
            }
        }

        result = 0m;
        return false;
    }

    private static bool TryGetDate(IXLCell cell, out DateTime date)
    {
        if (cell.DataType == XLDataType.DateTime)
        {
            date = cell.GetDateTime();
            return true;
        }

        var text = cell.GetString();
        if (!string.IsNullOrWhiteSpace(text))
        {
            foreach (var culture in DayFirstCultures)
            {
                if (DateTime.TryParseExact(text, DayFirstDateFormats, culture, DateTimeStyles.None, out date))
                {
                    return true;
                }
            }

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
                DateTime.TryParse(text, out date))
            {
                return true;
            }
        }

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var serialDate))
        {
            date = DateTime.FromOADate(serialDate);
            return true;
        }

        date = default;
        return false;
    }

    private static void WriteRates(IXLWorksheet worksheet, IEnumerable<InterestRatePeriod> rates)
    {
        worksheet.Cell(1, 1).Value = "Od";
        worksheet.Cell(1, 2).Value = "Do";
        worksheet.Cell(1, 3).Value = "Stopa (%)";

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

    private static void WriteSchedule(IXLWorksheet worksheet, IEnumerable<ScheduleItem> schedule, decimal totalInterest, decimal apr)
    {
        worksheet.Cell(1, 1).Value = "Data płatności";
        worksheet.Cell(1, 2).Value = "Dni w okresie";
        worksheet.Cell(1, 3).Value = "Stopa %";
        worksheet.Cell(1, 4).Value = "Odsetki";
        worksheet.Cell(1, 5).Value = "Spłata kapitału";
        worksheet.Cell(1, 6).Value = "Łączna płatność";
        worksheet.Cell(1, 7).Value = "Pozostały kapitał";

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

        var summaryRow = row + 1;
        worksheet.Cell(summaryRow, 1).Value = "Łączne odsetki";
        worksheet.Cell(summaryRow, 2).SetValue(totalInterest);
        worksheet.Cell(summaryRow, 1).Style.Font.Bold = true;
        worksheet.Cell(summaryRow, 2).Style.Font.Bold = true;

        worksheet.Cell(summaryRow + 1, 1).Value = "RRSO (APR)";
        worksheet.Cell(summaryRow + 1, 2).SetValue(apr);
        worksheet.Cell(summaryRow + 1, 1).Style.Font.Bold = true;
        worksheet.Cell(summaryRow + 1, 2).Style.Font.Bold = true;

        worksheet.Columns().AdjustToContents();
    }

    private static decimal ParseDecimal(IDictionary<string, string> parameters, string key, bool required = false, decimal defaultValue = 0m)
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brak wymaganej wartości liczbowej dla parametru {key}.");
            }

            return defaultValue;
        }

        if (TryParseDecimal(value, out var parsed))
        {
            return parsed;
        }

        if (required)
        {
            throw new InvalidOperationException($"Nie można odczytać wartości liczbowej dla parametru {key}.");
        }

        return defaultValue;
    }

    private static int ParseInt(IDictionary<string, string> parameters, string key, int defaultValue, bool required = false)
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brak wymaganej wartości liczbowej dla parametru {key}.");
            }

            return defaultValue;
        }

        if (int.TryParse(value, out var result))
        {
            return result;
        }

        throw new InvalidOperationException($"Nie można odczytać wartości liczbowej dla parametru {key}.");
    }

    private static DateTime ParseDate(IDictionary<string, string> parameters, string key, bool required)
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brak wymaganej daty dla parametru {key}.");
            }

            return default;
        }

        if (TryParseDateString(value, out var result))
        {
            return result;
        }

        throw new InvalidOperationException($"Nie można odczytać daty dla parametru {key}.");
    }

    private static bool TryParseDateString(string value, out DateTime date)
    {
        foreach (var culture in DayFirstCultures)
        {
            if (DateTime.TryParseExact(value, DayFirstDateFormats, culture, DateTimeStyles.None, out date))
            {
                return true;
            }
        }

        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
            DateTime.TryParse(value, out date))
        {
            return true;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var serialDate))
        {
            date = DateTime.FromOADate(serialDate);
            return true;
        }

        date = default;
        return false;
    }

    private static TEnum ParseEnum<TEnum>(IDictionary<string, string> parameters, string key, TEnum defaultValue, bool required = false) where TEnum : struct
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brak wymaganej wartości wyboru dla parametru {key}.");
            }

            return defaultValue;
        }

        if (Enum.TryParse<TEnum>(value, out var parsed))
        {
            return parsed;
        }

        throw new InvalidOperationException($"Nie można odczytać wartości dla parametru {key}.");
    }

    private static bool ParseBool(IDictionary<string, string> parameters, string key)
    {
        return parameters.TryGetValue(key, out var value) && bool.TryParse(value, out var result) && result;
    }

    private static void ValidateRequiredParameters(IDictionary<string, string> parameters)
    {
        var requiredKeys = new[]
        {
            "NetValue",
            "MarginRate",
            "PaymentFrequency",
            "PaymentDay",
            "CreditStartDate",
            "CreditEndDate",
            "DayCountBasis",
            "RoundingMode",
            "RoundingDecimals"
        };

        var missing = requiredKeys
            .Where(key => !parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            .ToList();

        if (missing.Count > 0)
        {
            throw new InvalidOperationException($"Brakuje wymaganych parametrów importu: {string.Join(", ", missing)}.");
        }
    }

    internal static string NormalizeParameterKey(string key)
    {
        var trimmed = key.Trim();
        return ParameterKeyAliases.TryGetValue(trimmed, out var canonical)
            ? canonical
            : trimmed;
    }
}
