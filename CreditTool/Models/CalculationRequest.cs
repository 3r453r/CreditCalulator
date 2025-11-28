namespace CreditTool.Models;

public class CalculationRequest
{
    public CreditParameters Parameters { get; set; } = new();

    public List<InterestRatePeriod> Rates { get; set; } = new();

    /// <summary>
    /// Optional pre-calculated log. When provided to export-log endpoint,
    /// this log will be used instead of recalculating to ensure consistency.
    /// </summary>
    public List<CalculationLogEntry>? CalculationLog { get; set; }
}
