namespace CreditTool.Models;

public enum PaymentFrequency
{
    Daily,
    Monthly,
    Quarterly
}

public enum PaymentDayOption
{
    FirstOfMonth,
    TenthOfMonth,
    LastOfMonth
}

public enum DayCountBasis
{
    Actual365,
    Actual360
}

public enum RoundingModeOption
{
    Bankers,
    AwayFromZero
}

public class CreditParameters
{
    public decimal NetValue { get; set; }

    public decimal MarginRate { get; set; }

    public PaymentFrequency PaymentFrequency { get; set; } = PaymentFrequency.Monthly;

    public PaymentDayOption PaymentDay { get; set; } = PaymentDayOption.LastOfMonth;

    public DateTime CreditStartDate { get; set; }

    public DateTime CreditEndDate { get; set; }

    public DayCountBasis DayCountBasis { get; set; } = DayCountBasis.Actual365;

    public RoundingModeOption RoundingMode { get; set; } = RoundingModeOption.Bankers;

    public int RoundingDecimals { get; set; } = 2;

    /// <summary>
    /// Optional upfront processing fee expressed as percentage of the net value.
    /// </summary>
    public decimal ProcessingFeeRate { get; set; }

    /// <summary>
    /// Flag that indicates whether principal is repaid evenly (default) or bullet at maturity.
    /// </summary>
    public bool BulletRepayment { get; set; }
}
