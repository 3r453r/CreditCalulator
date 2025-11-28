using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.Interest;

/// <summary>
/// Interest calculation strategy that applies rate changes only at the start of the next period.
/// Uses the rate at the beginning of the period for the entire period.
/// </summary>
public class ApplyRateNextPeriodStrategy : IInterestCalculationStrategy
{
    public InterestCalculationResult Calculate(
        DateTime from,
        DateTime to,
        decimal principal,
        decimal marginRate,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        DayCountBasis basis)
    {
        var daysInPeriod = Math.Max((to - from).Days, 0);
        var denominator = basis == DayCountBasis.Actual360 ? 360m : 365m;

        var baseRate = FindRateForDate(ratePeriods, from)?.Rate ?? 0m;
        var effectiveRate = baseRate + marginRate;
        var interest = principal * effectiveRate / 100m / denominator * daysInPeriod;

        return new InterestCalculationResult(interest, effectiveRate, null, null);
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period =>
            period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }
}
