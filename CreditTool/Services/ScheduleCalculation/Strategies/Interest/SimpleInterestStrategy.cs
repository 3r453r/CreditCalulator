using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.Interest;

/// <summary>
/// Simple interest calculation strategy using daily accrual.
/// Each day's interest is calculated independently and summed.
/// </summary>
public class SimpleInterestStrategy : IInterestCalculationStrategy
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

        decimal interest = 0m;
        decimal totalRate = 0m;

        for (var day = from; day < to; day = day.AddDays(1))
        {
            var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
            totalRate += rate;
            var effectiveRate = rate + marginRate;
            interest += principal * effectiveRate / 100m / denominator;
        }

        var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
        return new InterestCalculationResult(interest, averageRate + marginRate, null, null);
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period =>
            period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }
}
