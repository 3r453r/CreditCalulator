using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.Interest;

/// <summary>
/// Compound interest strategy with daily compounding.
/// Interest is compounded each day based on the daily rate.
/// </summary>
public class CompoundDailyStrategy : IInterestCalculationStrategy
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

        decimal factor = 1m;
        decimal totalRate = 0m;

        for (var day = from; day < to; day = day.AddDays(1))
        {
            var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
            totalRate += rate;
            var effectiveRateDaily = rate + marginRate;
            factor *= 1m + (effectiveRateDaily / 100m) / denominator;
        }

        var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
        var nominalRate = averageRate + marginRate;
        var periodRate = factor - 1m;
        var interest = principal * periodRate;

        return new InterestCalculationResult(
            Interest: interest,
            EffectiveRate: nominalRate,
            NominalRate: nominalRate,
            EffectivePeriodRate: periodRate,
            RateBreakdown: new List<RateBreakdownEntry>
            {
                new(
                    Days: daysInPeriod,
                    BaseRate: averageRate,
                    MarginRate: marginRate,
                    EffectiveRate: nominalRate,
                    InterestContribution: interest)
            });
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period =>
            period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }
}
