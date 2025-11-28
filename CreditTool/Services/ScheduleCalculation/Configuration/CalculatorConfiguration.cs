namespace CreditTool.Services.ScheduleCalculation.Configuration;

/// <summary>
/// Configuration options for the schedule calculator.
/// </summary>
public class CalculatorConfiguration
{
    /// <summary>
    /// Tolerance for level payment binary search (in currency units).
    /// </summary>
    public decimal LevelPaymentTolerance { get; set; } = 0.0001m;

    /// <summary>
    /// Maximum iterations for level payment calculation.
    /// </summary>
    public int MaxLevelPaymentIterations { get; set; } = 200;

    /// <summary>
    /// Maximum iterations for initial range finding.
    /// </summary>
    public int MaxRangeFindingIterations { get; set; } = 25;

    /// <summary>
    /// Enable validation checks (negative amortization, etc.).
    /// </summary>
    public bool EnableValidation { get; set; } = true;

    /// <summary>
    /// Throw exception on negative amortization detection.
    /// </summary>
    public bool ThrowOnNegativeAmortization { get; set; } = false;
}
