using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.Interest;

/// <summary>
/// Strategy interface for calculating interest over a period.
/// </summary>
public interface IInterestCalculationStrategy
{
    /// <summary>
    /// Calculates interest for the given period.
    /// </summary>
    /// <param name="from">Start date of the interest period.</param>
    /// <param name="to">End date of the interest period.</param>
    /// <param name="principal">Principal amount on which to calculate interest.</param>
    /// <param name="marginRate">Margin rate to add to the base rate.</param>
    /// <param name="ratePeriods">Collection of interest rate periods.</param>
    /// <param name="basis">Day count basis for calculations.</param>
    /// <returns>Result containing calculated interest and effective rates.</returns>
    InterestCalculationResult Calculate(
        DateTime from,
        DateTime to,
        decimal principal,
        decimal marginRate,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        DayCountBasis basis);
}

/// <summary>
/// Result of an interest calculation including the interest amount and applicable rates.
/// </summary>
public record InterestCalculationResult(
    decimal Interest,
    decimal EffectiveRate,
    decimal? NominalRate,
    decimal? EffectivePeriodRate,
    IReadOnlyList<RateBreakdownEntry> RateBreakdown);

public record RateBreakdownEntry(
    int Days,
    decimal BaseRate,
    decimal MarginRate,
    decimal EffectiveRate,
    decimal InterestContribution);
