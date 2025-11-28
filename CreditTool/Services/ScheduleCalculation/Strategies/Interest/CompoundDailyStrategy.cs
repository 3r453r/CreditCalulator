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
            var previousFactor = factor;
            var dailyRate = 1m + (effectiveRate / 100m / denominator);
            var chunkFactor = DecimalMath.Power(dailyRate, days);

            factor *= chunkFactor;
            totalRate += baseRate * days;

            breakdown.Add(new RateBreakdownEntry(
                Days: days,
                BaseRate: baseRate,
                MarginRate: marginRate,
                EffectiveRate: effectiveRate,
                InterestContribution: principal * previousFactor * (chunkFactor - 1m)));

            currentDate = chunkEnd;
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
            RateBreakdown: breakdown);
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period =>
            period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }
}
