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
        decimal totalEffectiveRate = 0m;
        var breakdown = new List<RateBreakdownEntry>();

        var currentDate = from;
        while (currentDate < to)
        {
            var period = FindRateForDate(ratePeriods, currentDate);
            var nextChangeDate = period?.DateTo.AddDays(1) ?? to;
            var chunkEnd = nextChangeDate > to ? to : nextChangeDate;

            var days = Math.Max((chunkEnd - currentDate).Days, 0);
            if (days == 0)
            {
                break;
            }

            var baseRate = period?.Rate ?? 0m;
            var effectiveRate = baseRate + marginRate;
            var contribution = principal * effectiveRate / 100m / denominator * days;

            interest += contribution;
            totalEffectiveRate += effectiveRate * days;

            breakdown.Add(new RateBreakdownEntry(
                Days: days,
                BaseRate: baseRate,
                MarginRate: marginRate,
                EffectiveRate: effectiveRate,
                InterestContribution: contribution));

            currentDate = chunkEnd;
        }

        var averageEffectiveRate = daysInPeriod > 0 ? totalEffectiveRate / daysInPeriod : 0m;
        return new InterestCalculationResult(interest, averageEffectiveRate, null, null, breakdown);
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period =>
            period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }
}
