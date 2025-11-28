namespace CreditTool.Models;

public class ScheduleCalculationResult
{
    public List<ScheduleItem> Schedule { get; set; } = new();
    public List<CalculationLogEntry> CalculationLog { get; set; } = new();

    /// <summary>
    /// Validation warnings encountered during calculation
    /// </summary>
    public List<string> Warnings { get; set; } = new();

    /// <summary>
    /// The target fixed payment for equal installments (before final adjustment)
    /// </summary>
    public decimal? TargetLevelPayment { get; set; }

    /// <summary>
    /// Actual final payment amount (if different from target)
    /// </summary>
    public decimal? ActualFinalPayment { get; set; }
}
