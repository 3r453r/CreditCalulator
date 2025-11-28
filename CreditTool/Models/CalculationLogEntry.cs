namespace CreditTool.Models;

public class CalculationLogEntry
{
    public string ShortDescription { get; set; } = string.Empty;

    public string SymbolicFormula { get; set; } = string.Empty;

    public string SubstitutedFormula { get; set; } = string.Empty;

    public string Result { get; set; } = string.Empty;

    /// <summary>
    /// Optional context information for grouping and organization
    /// </summary>
    public LogEntryContext? Context { get; set; }
}

public class LogEntryContext
{
    /// <summary>
    /// Payment period number (1-based)
    /// </summary>
    public int? PaymentNumber { get; set; }

    /// <summary>
    /// Payment date
    /// </summary>
    public DateTime? PaymentDate { get; set; }

    /// <summary>
    /// Entry type for better organization
    /// </summary>
    public LogEntryType Type { get; set; }

    /// <summary>
    /// Additional metadata as key-value pairs
    /// </summary>
    public Dictionary<string, string>? Metadata { get; set; }
}

public enum LogEntryType
{
    Header,           // Section header (e.g., "Payment Period #1")
    RateChange,       // Interest rate change notification
    PeriodCalculation,// Days in period calculation
    InterestCalculation, // Interest amount calculation
    PrincipalCalculation, // Principal payment calculation
    BalanceUpdate,    // Remaining balance update
    Summary,          // Summary information
    Detail            // Detailed breakdown (e.g., day-by-day for variable rates)
}
