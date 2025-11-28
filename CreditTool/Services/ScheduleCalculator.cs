using CreditTool.Models;

namespace CreditTool.Services;

public interface IScheduleCalculator
{
    IReadOnlyList<ScheduleItem> Calculate(CreditParameters parameters, IReadOnlyCollection<InterestRatePeriod> ratePeriods);
}

public class DayToDayScheduleCalculator : IScheduleCalculator
{
    public IReadOnlyList<ScheduleItem> Calculate(CreditParameters parameters, IReadOnlyCollection<InterestRatePeriod> ratePeriods)
    {
        if (parameters.CreditEndDate <= parameters.CreditStartDate)
        {
            throw new ArgumentException("Credit end date must be after start date");
        }

        var paymentDates = BuildPaymentDates(parameters.CreditStartDate, parameters.CreditEndDate, parameters.PaymentFrequency, parameters.PaymentDay);
        if (paymentDates.Count == 0 || paymentDates[^1] != parameters.CreditEndDate)
        {
            paymentDates.Add(parameters.CreditEndDate);
        }

        var principalRemaining = parameters.NetValue;
        var paymentCount = paymentDates.Count;
        var principalStep = parameters.PaymentType == PaymentType.Bullet
            ? 0m
            : RoundingService.Round(parameters.NetValue / paymentCount, parameters.RoundingMode, parameters.RoundingDecimals);

        decimal? fixedTotalPayment = null;
        if (parameters.PaymentType == PaymentType.EqualInstallments)
        {
            fixedTotalPayment = CalculateLevelPayment(parameters, ratePeriods, paymentDates);
        }

        var schedule = new List<ScheduleItem>();
        var previousDate = parameters.CreditStartDate;

        foreach (var paymentDate in paymentDates)
        {
            var daysInPeriod = (paymentDate - previousDate).Days;
            var interest = CalculateInterestDaily(previousDate, paymentDate, principalRemaining, parameters.MarginRate, ratePeriods, parameters.DayCountBasis);
            var interestRounded = RoundingService.Round(interest, parameters.RoundingMode, parameters.RoundingDecimals);

            decimal principalPayment;
            if (parameters.PaymentType == PaymentType.Bullet && paymentDate != paymentDates[^1])
            {
                principalPayment = 0m;
            }
            else if (parameters.PaymentType == PaymentType.Bullet || parameters.PaymentType == PaymentType.DecreasingInstallments)
            {
                principalPayment = paymentDate == paymentDates[^1] ? principalRemaining : principalStep;
            }
            else
            {
                var targetTotal = fixedTotalPayment!.Value;
                principalPayment = targetTotal - interestRounded;
                if (principalPayment < 0m)
                {
                    principalPayment = 0m;
                }

                if (paymentDate != paymentDates[^1])
                {
                    principalPayment = Math.Min(principalPayment, principalRemaining);
                }
                else
                {
                    principalPayment = principalRemaining;
                }
            }

            principalPayment = RoundingService.Round(principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            principalRemaining = RoundingService.Round(principalRemaining - principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            var totalPayment = fixedTotalPayment ?? (interestRounded + principalPayment);

            var periodRate = FindRateForDate(ratePeriods, previousDate)?.Rate ?? 0m;
            schedule.Add(new ScheduleItem
            {
                PaymentDate = paymentDate,
                DaysInPeriod = daysInPeriod,
                InterestRate = periodRate + parameters.MarginRate,
                InterestAmount = interestRounded,
                PrincipalPayment = principalPayment,
                TotalPayment = totalPayment,
                RemainingPrincipal = Math.Max(principalRemaining, 0m)
            });

            previousDate = paymentDate;
        }

        return schedule;
    }

    private decimal CalculateLevelPayment(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates)
    {
        var low = 0m;
        var high = Math.Max(parameters.NetValue, parameters.NetValue + 1000m);
        var tolerance = 0.01m;

        var remainingHigh = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, high);
        var guard = 0;
        while (remainingHigh > 0m && guard < 25)
        {
            low = high;
            high *= 2m;
            remainingHigh = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, high);
            guard++;
        }

        for (var i = 0; i < 200; i++)
        {
            var mid = (low + high) / 2m;
            var remaining = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, mid);

            if (Math.Abs(high - low) <= tolerance)
            {
                break;
            }

            if (remaining > 0)
            {
                low = mid;
            }
            else
            {
                high = mid;
            }
        }

        return RoundingService.Round((low + high) / 2m, parameters.RoundingMode, parameters.RoundingDecimals);
    }

    private decimal SimulateRemainingPrincipal(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates,
        decimal totalPayment)
    {
        var principalRemaining = parameters.NetValue;
        var previousDate = parameters.CreditStartDate;

        foreach (var paymentDate in paymentDates)
        {
            var interest = CalculateInterestDaily(previousDate, paymentDate, principalRemaining, parameters.MarginRate, ratePeriods, parameters.DayCountBasis);
            var interestRounded = RoundingService.Round(interest, parameters.RoundingMode, parameters.RoundingDecimals);

            var principalPayment = totalPayment - interestRounded;
            if (principalPayment < 0m)
            {
                principalPayment = 0m;
            }

            principalPayment = Math.Min(principalPayment, principalRemaining);

            principalPayment = RoundingService.Round(principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            principalRemaining = RoundingService.Round(principalRemaining - principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            previousDate = paymentDate;
        }

        return principalRemaining;
    }

    private static decimal CalculateInterestDaily(DateTime from, DateTime to, decimal principal, decimal marginRate, IReadOnlyCollection<InterestRatePeriod> ratePeriods, DayCountBasis basis)
    {
        decimal interest = 0m;
        var denominator = basis == DayCountBasis.Actual360 ? 360m : 365m;

        for (var day = from; day < to; day = day.AddDays(1))
        {
            var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
            var effectiveRate = rate + marginRate;
            interest += principal * effectiveRate / 100m / denominator;
        }

        return interest;
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period => period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }

    private static List<DateTime> BuildPaymentDates(DateTime start, DateTime end, PaymentFrequency frequency, PaymentDayOption paymentDay)
    {
        var dates = new List<DateTime>();
        var current = start;

        while (current < end)
        {
            DateTime next = frequency switch
            {
                PaymentFrequency.Daily => current.AddDays(1),
                PaymentFrequency.Monthly => NextMonthDate(current, 1, paymentDay),
                PaymentFrequency.Quarterly => NextMonthDate(current, 3, paymentDay),
                _ => current.AddMonths(1)
            };

            if (next > end)
            {
                next = end;
            }

            dates.Add(next);
            current = next;
        }

        return dates;
    }

    private static DateTime NextMonthDate(DateTime from, int monthsToAdd, PaymentDayOption paymentDay)
    {
        var tentative = from.AddMonths(monthsToAdd);
        var year = tentative.Year;
        var month = tentative.Month;

        return paymentDay switch
        {
            PaymentDayOption.FirstOfMonth => new DateTime(year, month, 1),
            PaymentDayOption.TenthOfMonth => new DateTime(year, month, 10),
            PaymentDayOption.LastOfMonth => new DateTime(year, month, DateTime.DaysInMonth(year, month)),
            _ => tentative
        };
    }
}
