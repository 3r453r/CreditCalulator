using ClosedXML.Excel;
using CreditTool.Services;
using CreditTool.Models;

namespace CreditTool.Tests;

public class ExcelServiceTests
{
    [Theory]
    [InlineData("15/02/2024", "16/02/2024", 2024, 2, 15, 2024, 2, 16)]
    [InlineData("1/3/2024", "5/3/2024", 2024, 3, 1, 2024, 3, 5)]
    [InlineData("31.03.2024", "01.04.2024", 2024, 3, 31, 2024, 4, 1)]
    public void ImportAcceptsDayFirstFormats(string fromText, string toText, int expectedFromYear, int expectedFromMonth, int expectedFromDay, int expectedToYear, int expectedToMonth, int expectedToDay)
    {
        using var workbook = new XLWorkbook();
        var parameterSheet = workbook.AddWorksheet("Parametry");
        WriteRequiredParameters(parameterSheet);
        var rateSheet = workbook.AddWorksheet("Stopy procentowe");
        rateSheet.Cell(1, 1).Value = "Od";
        rateSheet.Cell(1, 2).Value = "Do";
        rateSheet.Cell(1, 3).Value = "Stopa (%)";
        rateSheet.Cell(2, 1).Value = fromText;
        rateSheet.Cell(2, 2).Value = toText;
        rateSheet.Cell(2, 3).Value = "5";

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        var service = new ExcelService();
        var (_, rates) = service.Import(stream);

        Assert.Single(rates);
        var period = rates[0];
        Assert.Equal(new DateTime(expectedFromYear, expectedFromMonth, expectedFromDay), period.DateFrom);
        Assert.Equal(new DateTime(expectedToYear, expectedToMonth, expectedToDay), period.DateTo);
    }

    [Fact]
    public void ImportThrowsWhenRateDatesMissing()
    {
        using var workbook = new XLWorkbook();
        var parameterSheet = workbook.AddWorksheet("Parametry");
        WriteRequiredParameters(parameterSheet);
        var rateSheet = workbook.AddWorksheet("Stopy procentowe");
        rateSheet.Cell(1, 1).Value = "Od";
        rateSheet.Cell(1, 2).Value = "Do";
        rateSheet.Cell(1, 3).Value = "Stopa (%)";
        rateSheet.Cell(2, 1).Value = "15/02/2024";
        rateSheet.Cell(2, 3).Value = "5";

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        var service = new ExcelService();

        Assert.Throws<InvalidOperationException>(() => service.Import(stream));
    }

    private static void WriteRequiredParameters(IXLWorksheet sheet)
    {
        sheet.Cell(1, 1).Value = "Parametr";
        sheet.Cell(1, 2).Value = "Wartość";

        var row = 2;
        sheet.Cell(row++, 1).Value = "Kwota netto";
        sheet.Cell(row - 1, 2).Value = 100000;

        sheet.Cell(row++, 1).Value = "Marża";
        sheet.Cell(row - 1, 2).Value = 2.5;

        sheet.Cell(row++, 1).Value = "Częstotliwość płatności";
        sheet.Cell(row - 1, 2).Value = PaymentFrequency.Monthly.ToString();

        sheet.Cell(row++, 1).Value = "Dzień płatności";
        sheet.Cell(row - 1, 2).Value = PaymentDayOption.LastOfMonth.ToString();

        sheet.Cell(row++, 1).Value = "Data początkowa";
        sheet.Cell(row - 1, 2).Value = new DateTime(2024, 1, 1);

        sheet.Cell(row++, 1).Value = "Data końcowa";
        sheet.Cell(row - 1, 2).Value = new DateTime(2025, 1, 1);

        sheet.Cell(row++, 1).Value = "Konwencja dni";
        sheet.Cell(row - 1, 2).Value = DayCountBasis.Actual365.ToString();

        sheet.Cell(row++, 1).Value = "Zaokrąglanie";
        sheet.Cell(row - 1, 2).Value = RoundingModeOption.Bankers.ToString();

        sheet.Cell(row++, 1).Value = "Miejsca po przecinku";
        sheet.Cell(row - 1, 2).Value = 4;
    }
}
