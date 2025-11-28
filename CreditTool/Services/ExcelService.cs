using ClosedXML.Excel;
using CreditTool.Models;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;

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

    public (CreditParameters Parameters, List<InterestRatePeriod> Rates) Import(Stream fileStream, string extension)
    {
        return extension.ToLowerInvariant() switch
        {
            ".xlsx" => ImportXlsx(fileStream),
            ".ods" => ImportOds(fileStream),
            _ => throw new InvalidOperationException("Nieobsługiwany format pliku. Użyj .xlsx lub .ods")
        };
    }

    public async Task<(CreditParameters Parameters, List<InterestRatePeriod> Rates)> ImportJsonAsync(Stream fileStream)
    {
        var request = await JsonSerializer.DeserializeAsync<CalculationRequest>(fileStream, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            Converters = { new JsonStringEnumConverter() }
        });

        if (request?.Parameters is null || request.Rates is null || request.Rates.Count == 0)
        {
            throw new InvalidOperationException("Nieprawidłowy plik JSON z parametrami lub stopami procentowymi.");
        }

        return (request.Parameters, request.Rates);
    }

    public byte[] ExportJson(CreditParameters parameters, IEnumerable<InterestRatePeriod> rates, IEnumerable<ScheduleItem> schedule, decimal totalInterest, decimal annualPercentageRate)
    {
        var payload = new
        {
            parameters,
            rates,
            schedule,
            totalInterest,
            annualPercentageRate
        };

        return JsonSerializer.SerializeToUtf8Bytes(payload, new JsonSerializerOptions
        {
            WriteIndented = true,
            Converters = { new JsonStringEnumConverter() }
        });
    }

    public byte[] ExportXlsx(CreditParameters parameters, IEnumerable<InterestRatePeriod> rates, IEnumerable<ScheduleItem> schedule, decimal totalInterest, decimal annualPercentageRate)
    {
        using var workbook = new XLWorkbook();
        var parameterSheet = workbook.AddWorksheet("Parametry");
        WriteParameters(parameterSheet, parameters);

        var rateSheet = workbook.AddWorksheet("Stopy procentowe");
        WriteRates(rateSheet, rates);

        var scheduleSheet = workbook.AddWorksheet("Harmonogram");
        WriteSchedule(scheduleSheet, schedule, totalInterest, annualPercentageRate);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportOds(CreditParameters parameters, IEnumerable<InterestRatePeriod> rates, IEnumerable<ScheduleItem> schedule, decimal totalInterest, decimal annualPercentageRate)
    {
        var parameterRows = BuildParameterRows(parameters);
        var rateRows = BuildRateRows(rates);
        var scheduleRows = BuildScheduleRows(schedule, totalInterest, annualPercentageRate);

        var contentDocument = BuildOdsContent(parameterRows, rateRows, scheduleRows);
        var manifest = BuildOdsManifest();

        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            var mimeEntry = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
            using (var writer = new StreamWriter(mimeEntry.Open(), new UTF8Encoding(false)))
            {
                writer.Write("application/vnd.oasis.opendocument.spreadsheet");
            }

            var contentEntry = archive.CreateEntry("content.xml", CompressionLevel.Optimal);
            using (var writer = new StreamWriter(contentEntry.Open(), new UTF8Encoding(false)))
            {
                writer.Write(contentDocument);
            }

            var manifestEntry = archive.CreateEntry("META-INF/manifest.xml", CompressionLevel.Optimal);
            using (var writer = new StreamWriter(manifestEntry.Open(), new UTF8Encoding(false)))
            {
                writer.Write(manifest);
            }
        }

        return stream.ToArray();
    }

    private static (CreditParameters Parameters, List<InterestRatePeriod> Rates) ImportXlsx(Stream fileStream)
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

    private (CreditParameters Parameters, List<InterestRatePeriod> Rates) ImportOds(Stream fileStream)
    {
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Read, leaveOpen: true);
        var contentEntry = archive.GetEntry("content.xml") ?? throw new InvalidOperationException("Nie znaleziono danych arkusza w pliku ODS.");
        using var contentStream = contentEntry.Open();
        var document = XDocument.Load(contentStream);

        XNamespace tableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        var tables = document.Descendants(tableNs + "table").ToList();
        if (tables.Count < 2)
        {
            throw new InvalidOperationException("Nie znaleziono wymaganych tabel w pliku ODS.");
        }

        var parameterTable = tables.First();
        var rateTable = tables.Skip(1).First();

        var parameterMap = ReadOdsParameterMap(parameterTable);
        ValidateRequiredParameters(parameterMap);

        var parameters = new CreditParameters
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

        var rates = ReadOdsRates(rateTable);
        if (rates.Count == 0)
        {
            throw new InvalidOperationException("Brak stóp procentowych w arkuszu importu.");
        }

        return (parameters, rates);
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

    private static List<List<OdsCell>> BuildParameterRows(CreditParameters parameters)
    {
        var rows = new List<List<OdsCell>>
        {
            new() { new OdsCell("Parametr"), new OdsCell("Wartość") }
        };

        var values = new (string Label, object? Value)[]
        {
            ("Kwota netto", parameters.NetValue),
            ("Marża", parameters.MarginRate),
            ("Częstotliwość płatności", parameters.PaymentFrequency),
            ("Dzień płatności", parameters.PaymentDay),
            ("Data początkowa", parameters.CreditStartDate),
            ("Data końcowa", parameters.CreditEndDate),
            ("Konwencja dni", parameters.DayCountBasis),
            ("Zaokrąglanie", parameters.RoundingMode),
            ("Miejsca po przecinku", parameters.RoundingDecimals),
            ("Prowizja przygotowawcza", parameters.ProcessingFeeRate),
            ("Prowizja przygotowawcza (kwota)", parameters.ProcessingFeeAmount),
            ("Typ spłaty", parameters.PaymentType),
            ("Spłata balonowa", parameters.BulletRepayment)
        };

        foreach (var (label, value) in values)
        {
            rows.Add(new List<OdsCell>
            {
                new(label),
                OdsCell.FromValue(value)
            });
        }

        return rows;
    }

    private static List<List<OdsCell>> BuildRateRows(IEnumerable<InterestRatePeriod> rates)
    {
        var rows = new List<List<OdsCell>>
        {
            new() { new OdsCell("Od"), new OdsCell("Do"), new OdsCell("Stopa (%)") }
        };

        foreach (var rate in rates)
        {
            rows.Add(new List<OdsCell>
            {
                new(rate.DateFrom),
                new(rate.DateTo),
                new OdsCell(rate.Rate)
            });
        }

        return rows;
    }

    private static List<List<OdsCell>> BuildScheduleRows(IEnumerable<ScheduleItem> schedule, decimal totalInterest, decimal apr)
    {
        var rows = new List<List<OdsCell>>
        {
            new()
            {
                new OdsCell("Data płatności"),
                new OdsCell("Dni w okresie"),
                new OdsCell("Stopa %"),
                new OdsCell("Odsetki"),
                new OdsCell("Spłata kapitału"),
                new OdsCell("Łączna płatność"),
                new OdsCell("Pozostały kapitał")
            }
        };

        foreach (var item in schedule)
        {
            rows.Add(new List<OdsCell>
            {
                new(item.PaymentDate),
                new OdsCell(item.DaysInPeriod),
                new OdsCell(item.InterestRate),
                new OdsCell(item.InterestAmount),
                new OdsCell(item.PrincipalPayment),
                new OdsCell(item.TotalPayment),
                new OdsCell(item.RemainingPrincipal)
            });
        }

        rows.Add(new List<OdsCell> { new("Łączne odsetki"), new(totalInterest) });
        rows.Add(new List<OdsCell> { new("RRSO (APR)"), new(apr) });

        return rows;
    }

    private static string BuildOdsContent(List<List<OdsCell>> parameterRows, List<List<OdsCell>> rateRows, List<List<OdsCell>> scheduleRows)
    {
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

        var document = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(office + "document-content",
                new XAttribute(XNamespace.Xmlns + "office", office),
                new XAttribute(XNamespace.Xmlns + "table", table),
                new XAttribute(XNamespace.Xmlns + "text", text),
                new XAttribute(office + "version", "1.2"),
                new XElement(office + "automatic-styles"),
                new XElement(office + "body",
                    new XElement(office + "spreadsheet",
                        BuildOdsTable("Parametry", parameterRows, table, text),
                        BuildOdsTable("Stopy procentowe", rateRows, table, text),
                        BuildOdsTable("Harmonogram", scheduleRows, table, text))
                )));

        return document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static string BuildOdsManifest()
    {
        XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        var manifestDoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(manifest + "manifest",
                new XAttribute(XNamespace.Xmlns + "manifest", manifest),
                new XElement(manifest + "file-entry",
                    new XAttribute(manifest + "full-path", "/"),
                    new XAttribute(manifest + "media-type", "application/vnd.oasis.opendocument.spreadsheet"),
                    new XAttribute(manifest + "version", "1.2")),
                new XElement(manifest + "file-entry",
                    new XAttribute(manifest + "full-path", "content.xml"),
                    new XAttribute(manifest + "media-type", "text/xml"))));

        return manifestDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static XElement BuildOdsTable(string name, IEnumerable<IEnumerable<OdsCell>> rows, XNamespace table, XNamespace text)
    {
        return new XElement(table + "table",
            new XAttribute(table + "name", name),
            rows.Select(row => new XElement(table + "table-row",
                row.Select(cell => cell.ToXElement(table, text)))));
    }

    private static Dictionary<string, string> ReadOdsParameterMap(XElement table)
    {
        XNamespace tableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var rows = ParseOdsTableRows(table, tableNs, text).Skip(1); // Skip header
        foreach (var row in rows)
        {
            if (row.Count < 2)
            {
                continue;
            }

            var key = row[0];
            var value = row[1];
            if (!string.IsNullOrWhiteSpace(key))
            {
                var normalized = NormalizeParameterKey(key);
                map[normalized] = value;
            }
        }

        return map;
    }

    private static List<InterestRatePeriod> ReadOdsRates(XElement table)
    {
        XNamespace tableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

        var rates = new List<InterestRatePeriod>();
        var rows = ParseOdsTableRows(table, tableNs, text).Skip(1); // Skip header
        var rowIndex = 2;
        foreach (var row in rows)
        {
            if (row.Count < 3)
            {
                rowIndex++;
                continue;
            }

            if (!TryParseDateString(row[0], out var from) || !TryParseDateString(row[1], out var to))
            {
                throw new InvalidOperationException($"Nieprawidłowa data w wierszu {rowIndex} tabeli stóp procentowych.");
            }

            if (TryParseDecimal(row[2], out var rate))
            {
                rates.Add(new InterestRatePeriod
                {
                    DateFrom = from,
                    DateTo = to,
                    Rate = rate
                });
            }
            else
            {
                throw new InvalidOperationException($"Nieprawidłowa stopa procentowa w wierszu {rowIndex} tabeli stóp procentowych.");
            }

            rowIndex++;
        }

        return rates;
    }

    private static List<List<string>> ParseOdsTableRows(XElement tableElement, XNamespace tableNs, XNamespace textNs)
    {
        var rows = new List<List<string>>();
        foreach (var row in tableElement.Elements(tableNs + "table-row"))
        {
            var cells = new List<string>();
            foreach (var cell in row.Elements(tableNs + "table-cell"))
            {
                var repeatAttr = (int?)cell.Attribute(tableNs + "number-columns-repeated") ?? 1;
                var text = string.Join("\n", cell.Elements(textNs + "p").Select(p => (string?)p ?? string.Empty));
                for (var i = 0; i < repeatAttr; i++)
                {
                    cells.Add(text);
                }
            }

            rows.Add(cells);
        }

        return rows;
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

internal readonly struct OdsCell
{
    public enum OdsCellType
    {
        String,
        Number,
        Date
    }

    public object? Value { get; }
    public OdsCellType Type { get; }

    public OdsCell(string text)
    {
        Value = text;
        Type = OdsCellType.String;
    }

    public OdsCell(decimal number)
    {
        Value = number;
        Type = OdsCellType.Number;
    }

    public OdsCell(int number)
    {
        Value = number;
        Type = OdsCellType.Number;
    }

    public OdsCell(DateTime date)
    {
        Value = date;
        Type = OdsCellType.Date;
    }

    public XElement ToXElement(XNamespace table, XNamespace text)
    {
        var cell = new XElement(table + "table-cell");
        switch (Type)
        {
            case OdsCellType.Number:
                var numeric = Convert.ToDecimal(Value, CultureInfo.InvariantCulture);
                cell.SetAttributeValue(XName.Get("value-type", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"), "float");
                cell.SetAttributeValue(XName.Get("value", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"), numeric.ToString(CultureInfo.InvariantCulture));
                cell.Add(new XElement(text + "p", numeric.ToString(CultureInfo.InvariantCulture)));
                break;
            case OdsCellType.Date:
                var date = (DateTime)Value!;
                cell.SetAttributeValue(XName.Get("value-type", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"), "date");
                cell.SetAttributeValue(XName.Get("date-value", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"), date.ToString("yyyy-MM-dd"));
                cell.Add(new XElement(text + "p", date.ToString("yyyy-MM-dd")));
                break;
            default:
                cell.SetAttributeValue(XName.Get("value-type", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"), "string");
                cell.Add(new XElement(text + "p", Value?.ToString() ?? string.Empty));
                break;
        }

        return cell;
    }

    public static OdsCell FromValue(object? value)
    {
        return value switch
        {
            DateTime date => new OdsCell(date),
            decimal number => new OdsCell(number),
            int number => new OdsCell(number),
            double number => new OdsCell(Convert.ToDecimal(number, CultureInfo.InvariantCulture)),
            bool flag => new OdsCell(flag ? bool.TrueString : bool.FalseString),
            Enum enumValue => new OdsCell(enumValue.ToString()),
            _ when value is null => new OdsCell(string.Empty),
            _ => new OdsCell(value.ToString() ?? string.Empty)
        };
    }
}
