using CreditTool.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CreditTool.Services;

public class WordExportService
{
    public byte[] Export(CreditParameters parameters, IEnumerable<InterestRatePeriod> rates, IEnumerable<ScheduleItem> schedule, decimal totalInterest)
    {
        using var memoryStream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true);
        var mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var body = mainPart.Document.Body!;
        body.Append(CreateHeading("Harmonogram kredytu"));
        body.Append(CreateParagraph($"Kwota netto: {parameters.NetValue:N2}"));
        body.Append(CreateParagraph($"Marża: {parameters.MarginRate:N2}%"));
        body.Append(CreateParagraph($"Okres od {parameters.CreditStartDate:yyyy-MM-dd} do {parameters.CreditEndDate:yyyy-MM-dd}"));
        var apr = AprCalculator.CalculateAnnualPercentageRate(parameters, schedule);
        body.Append(CreateParagraph($"Łączne odsetki: {totalInterest:N2}"));
        body.Append(CreateParagraph($"RRSO (APR): {apr:N4}%"));

        body.Append(CreateHeading("Parametry"));
        body.Append(CreateParameterTable(parameters));

        body.Append(CreateHeading("Stopy procentowe"));
        body.Append(CreateRateTable(rates));

        body.Append(CreateHeading("Harmonogram spłat"));
        body.Append(CreateScheduleTable(schedule));

        mainPart.Document.Save();
        return memoryStream.ToArray();
    }

    private static Paragraph CreateHeading(string text)
    {
        return new Paragraph(new Run(new Text(text)))
        {
            ParagraphProperties = new ParagraphProperties
            {
                ParagraphStyleId = new ParagraphStyleId { Val = "Heading1" },
                SpacingBetweenLines = new SpacingBetweenLines { After = "200" }
            }
        };
    }

    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    private static Table CreateParameterTable(CreditParameters parameters)
    {
        var rows = new (string Label, string Value)[]
        {
            ("Kwota netto", parameters.NetValue.ToString("N2")),
            ("Marża", $"{parameters.MarginRate:N2}%"),
            ("Częstotliwość płatności", parameters.PaymentFrequency.ToString()),
            ("Dzień płatności", parameters.PaymentDay.ToString()),
            ("Data początkowa", parameters.CreditStartDate.ToString("yyyy-MM-dd")),
            ("Data końcowa", parameters.CreditEndDate.ToString("yyyy-MM-dd")),
            ("Konwencja dni", parameters.DayCountBasis.ToString()),
            ("Zaokrąglanie", parameters.RoundingMode.ToString()),
            ("Miejsca po przecinku", parameters.RoundingDecimals.ToString()),
            ("Prowizja przygotowawcza", $"{parameters.ProcessingFeeRate:N2}%"),
            ("Prowizja przygotowawcza (kwota)", parameters.ProcessingFeeAmount.ToString("N2")),
            ("Typ spłaty", parameters.PaymentType.ToString()),
            ("Spłata balonowa", parameters.BulletRepayment ? "Tak" : "Nie")
        };

        return BuildTable(new[] { "Parametr", "Wartość" }, rows.Select(row => new[] { row.Label, row.Value }));
    }

    private static Table CreateRateTable(IEnumerable<InterestRatePeriod> rates)
    {
        var rateRows = rates.Select(rate => new[]
        {
            rate.DateFrom.ToString("yyyy-MM-dd"),
            rate.DateTo.ToString("yyyy-MM-dd"),
            rate.Rate.ToString("N4")
        });

        return BuildTable(new[] { "Od", "Do", "Stopa (%)" }, rateRows);
    }

    private static Table CreateScheduleTable(IEnumerable<ScheduleItem> schedule)
    {
        var scheduleRows = schedule.Select(item => new[]
        {
            item.PaymentDate.ToString("yyyy-MM-dd"),
            item.DaysInPeriod.ToString(),
            item.InterestRate.ToString("N4"),
            item.InterestAmount.ToString("N2"),
            item.PrincipalPayment.ToString("N2"),
            item.TotalPayment.ToString("N2"),
            item.RemainingPrincipal.ToString("N2")
        });

        return BuildTable(
            new[] { "Data płatności", "Dni", "Stopa %", "Odsetki", "Kapitał", "Płatność", "Pozostały kapitał" },
            scheduleRows);
    }

    private static Table BuildTable(IEnumerable<string> headers, IEnumerable<IEnumerable<string>> rows)
    {
        var table = new Table();
        table.Append(CreateTableProperties());

        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            headerRow.Append(CreateCell(header, true));
        }

        table.Append(headerRow);

        foreach (var rowValues in rows)
        {
            var row = new TableRow();
            foreach (var value in rowValues)
            {
                row.Append(CreateCell(value, false));
            }

            table.Append(row);
        }

        return table;
    }

    private static TableProperties CreateTableProperties()
    {
        return new TableProperties(
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }));
    }

    private static TableCell CreateCell(string text, bool isHeader)
    {
        var run = new Run(new Text(text));
        var paragraph = new Paragraph(run);
        if (isHeader)
        {
            run.RunProperties = new RunProperties(new Bold());
        }

        return new TableCell(paragraph);
    }
}
