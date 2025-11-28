using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.Interest;

/// <summary>
/// Compound interest strategy with periodic compounding (monthly or quarterly).
/// </summary>
public class CompoundPeriodicStrategy : IInterestCalculationStrategy
{
    private readonly int _compoundingPeriodsPerYear;

    /// <summary>
    /// Creates a periodic compounding strategy.
    /// </summary>
    /// <param name="compoundingPeriodsPerYear">Number of compounding periods per year (12 for monthly, 4 for quarterly).</param>
    public CompoundPeriodicStrategy(int compoundingPeriodsPerYear)
    {
        _compoundingPeriodsPerYear = compoundingPeriodsPerYear;
    }

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

        decimal totalRate = 0m;
        for (var day = from; day < to; day = day.AddDays(1))
        {
            var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
            totalRate += rate;
        }

        var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
        var nominalRate = averageRate + marginRate;
        var n = (decimal)_compoundingPeriodsPerYear;
        var tYears = daysInPeriod / denominator;
        var rNom = nominalRate / 100m;

        var periodFactor = (decimal)Math.Pow(1.0 + (double)(rNom / n), (double)(n * tYears));
        var periodRate = periodFactor - 1m;
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
