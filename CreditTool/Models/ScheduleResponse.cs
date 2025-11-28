namespace CreditTool.Models;

public class ScheduleResponse
{
    public List<ScheduleItem> Schedule { get; set; } = new();

    public List<CalculationLogEntry> CalculationLog { get; set; } = new();

    public decimal TotalInterest { get; set; }

    public decimal AnnualPercentageRate { get; set; }
}
