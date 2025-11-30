namespace CreditTool.Models;

public class ScheduleItem
{
    public DateTime PaymentDate { get; set; }
    public int DaysInPeriod { get; set; }

    /// <summary>
    /// The effective annual rate used for this period (including margin)
    /// </summary>
    public decimal InterestRate { get; set; }

    /// <summary>
    /// The nominal annual rate (for display purposes with compound interest)
    /// </summary>
    public decimal? NominalRate { get; set; }

    /// <summary>
    /// The effective period rate actually applied
    /// </summary>
    public decimal? EffectivePeriodRate { get; set; }

    public decimal InterestAmount { get; set; }
    public decimal PrincipalPayment { get; set; }
    public decimal TotalPayment { get; set; }
    public decimal RemainingPrincipal { get; set; }

    /// <summary>
    /// Indicates if this is an adjusted final payment
    /// </summary>
    public bool IsFinalPaymentAdjusted { get; set; }

    /// <summary>
    /// Indicates if this payment is within the grace period (interest-only)
    /// </summary>
    public bool IsInGracePeriod { get; set; }

    /// <summary>
    /// Warning flags for this period
    /// </summary>
    public ScheduleWarnings Warnings { get; set; } = ScheduleWarnings.None;
}

[Flags]
public enum ScheduleWarnings
{
    None = 0,
    NegativeAmortization = 1,
    InterestExceedsPayment = 2,
    FinalPaymentAdjusted = 4
}
