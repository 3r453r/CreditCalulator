using CreditTool.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

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
            NetValue = ParseDecimal(map, "NetValue", required: true),
            MarginRate = ParseDecimal(map, "MarginRate", required: true),
            PaymentFrequency = ParseEnum(map, "PaymentFrequency", PaymentFrequency.Monthly, required: true),
            PaymentDay = ParseEnum(map, "PaymentDay", PaymentDayOption.LastOfMonth, required: true),
            CreditStartDate = ParseDate(map, "CreditStartDate", required: true),
            CreditEndDate = ParseDate(map, "CreditEndDate", required: true),
            DayCountBasis = ParseEnum(map, "DayCountBasis", DayCountBasis.Actual365, required: true),
            RoundingMode = ParseEnum(map, "RoundingMode", RoundingModeOption.Bankers, required: true),
            RoundingDecimals = ParseInt(map, "RoundingDecimals", 4, required: true),
            ProcessingFeeRate = ParseDecimal(map, "ProcessingFeeRate"),
            ProcessingFeeAmount = ParseDecimal(map, "ProcessingFeeAmount"),
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

    private static decimal ParseDecimal(IDictionary<string, string> parameters, string key, bool required = false)
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brakuje wymaganej wartości liczbowej dla parametru {key}.");
            }

            return 0m;
        }

        foreach (var culture in new[] { CultureInfo.InvariantCulture, CultureInfo.GetCultureInfo("pl-PL"), CultureInfo.CurrentCulture })
        {
            if (decimal.TryParse(value, NumberStyles.Any, culture, out var result))
            {
                return result;
            }
        }

        throw new InvalidOperationException($"Nie można odczytać wartości liczbowej dla parametru {key}.");
    }

    private static int ParseInt(IDictionary<string, string> parameters, string key, int defaultValue, bool required = false)
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brakuje wymaganej wartości liczbowej dla parametru {key}.");
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

        if (DateTime.TryParse(value, out var result))
        {
            return result;
        }

        throw new InvalidOperationException($"Nie można odczytać daty dla parametru {key}.");
    }

    private static TEnum ParseEnum<TEnum>(IDictionary<string, string> parameters, string key, TEnum defaultValue, bool required = false) where TEnum : struct
    {
        if (!parameters.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            if (required)
            {
                throw new InvalidOperationException($"Brakuje wymaganej wartości wyboru dla parametru {key}.");
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
}
