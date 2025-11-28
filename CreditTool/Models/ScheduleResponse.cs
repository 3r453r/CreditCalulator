namespace CreditTool.Models;

public class ScheduleResponse
{
    public List<ScheduleItem> Schedule { get; set; } = new();
    public List<CalculationLogEntry> CalculationLog { get; set; } = new();
    public decimal TotalInterest { get; set; }
    public decimal AnnualPercentageRate { get; set; }
    public List<string> Warnings { get; set; } = new();
    public decimal? TargetLevelPayment { get; set; }
    public decimal? ActualFinalPayment { get; set; }
}
