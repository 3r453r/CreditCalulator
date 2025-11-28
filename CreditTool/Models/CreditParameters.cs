using System;

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

public enum PaymentType
{
    EqualInstallments,
    DecreasingInstallments,
    Bullet
}

public enum InterestRateApplication
{
    DailyAccrual,
    ApplyChangedRateNextPeriod
}

public class CreditParameters
{
    public const int MinRoundingDecimals = 4;
    public const int MaxRoundingDecimals = 10;

    private int roundingDecimals = 4;

    public decimal NetValue { get; set; }

    public decimal MarginRate { get; set; }

    public PaymentFrequency PaymentFrequency { get; set; } = PaymentFrequency.Monthly;

    public PaymentDayOption PaymentDay { get; set; } = PaymentDayOption.LastOfMonth;

    public DateTime CreditStartDate { get; set; }

    public DateTime CreditEndDate { get; set; }

    public DayCountBasis DayCountBasis { get; set; } = DayCountBasis.Actual365;

    public RoundingModeOption RoundingMode { get; set; } = RoundingModeOption.Bankers;

    public int RoundingDecimals
    {
        get => roundingDecimals;
        set => roundingDecimals = Math.Clamp(value, MinRoundingDecimals, MaxRoundingDecimals);
    }

    /// <summary>
    /// Optional upfront processing fee expressed as percentage of the net value.
    /// </summary>
    public decimal ProcessingFeeRate { get; set; }

    /// <summary>
    /// Optional upfront processing fee expressed as a fixed amount.
    /// </summary>
    public decimal ProcessingFeeAmount { get; set; }

    /// <summary>
    /// Defines how the principal is amortized over time.
    /// </summary>
    public PaymentType PaymentType { get; set; } = PaymentType.DecreasingInstallments;

    /// <summary>
    /// Backward-compatible flag that maps to the bullet repayment option.
    /// </summary>
    public bool BulletRepayment
    {
        get => PaymentType == PaymentType.Bullet;
        set => PaymentType = value ? PaymentType.Bullet : PaymentType.DecreasingInstallments;
    }

    /// <summary>
    /// Defines how changing base rates are applied to interest accrual.
    /// </summary>
    public InterestRateApplication InterestRateApplication { get; set; } = InterestRateApplication.DailyAccrual;
}
