using CreditTool.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CreditTool.Services;

public class WordImportService
{
    public (CreditParameters Parameters, List<InterestRatePeriod> Rates) Import(Stream stream)
    {
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("Dokument jest pusty");
        var tables = body.Elements<Table>().ToList();
        if (tables.Count < 2)
        {
            throw new InvalidOperationException("Nie znaleziono tabel z parametrami i stopami.");
        }

        var parameters = ReadParameters(tables[0]);
        var rates = ReadRates(tables[1]);
        return (parameters, rates);
    }

    private static CreditParameters ReadParameters(Table parameterTable)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in parameterTable.Elements<TableRow>().Skip(1))
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count < 2)
            {
                continue;
            }

            var key = cells[0].InnerText.Trim();
            var value = cells[1].InnerText.Trim();
            if (!string.IsNullOrWhiteSpace(key))
            {
                var normalized = ExcelService.NormalizeParameterKey(key);
                map[normalized] = value;
            }
        }

        return new CreditParameters
        {
            NetValue = ParseDecimal(map, "NetValue"),
            MarginRate = ParseDecimal(map, "MarginRate"),
            PaymentFrequency = ParseEnum(map, "PaymentFrequency", PaymentFrequency.Monthly),
            PaymentDay = ParseEnum(map, "PaymentDay", PaymentDayOption.LastOfMonth),
            CreditStartDate = ParseDate(map, "CreditStartDate"),
            CreditEndDate = ParseDate(map, "CreditEndDate"),
            DayCountBasis = ParseEnum(map, "DayCountBasis", DayCountBasis.Actual365),
            RoundingMode = ParseEnum(map, "RoundingMode", RoundingModeOption.Bankers),
            RoundingDecimals = ParseInt(map, "RoundingDecimals", 4),
            ProcessingFeeRate = ParseDecimal(map, "ProcessingFeeRate"),
            PaymentType = ParseEnum(map, "PaymentType", PaymentType.DecreasingInstallments),
            BulletRepayment = ParseBool(map, "BulletRepayment")
        };
    }

    private static List<InterestRatePeriod> ReadRates(Table rateTable)
    {
        var rows = new List<InterestRatePeriod>();
        foreach (var row in rateTable.Elements<TableRow>().Skip(1))
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count < 3)
            {
                continue;
            }

            if (DateTime.TryParse(cells[0].InnerText, out var from) &&
                DateTime.TryParse(cells[1].InnerText, out var to) &&
                decimal.TryParse(cells[2].InnerText, out var rate))
            {
                rows.Add(new InterestRatePeriod
                {
                    DateFrom = from,
                    DateTo = to,
                    Rate = rate
                });
            }
        }

        return rows;
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
