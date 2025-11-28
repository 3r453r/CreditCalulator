namespace CreditTool.Models;

public class ScheduleCalculationResult
{
    public IReadOnlyList<ScheduleItem> Schedule { get; set; } = Array.Empty<ScheduleItem>();

    public List<CalculationLogEntry> CalculationLog { get; set; } = new();
}
