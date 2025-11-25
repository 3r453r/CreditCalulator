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
        var principalStep = parameters.BulletRepayment
            ? 0m
            : RoundingService.Round(parameters.NetValue / paymentCount, parameters.RoundingMode, parameters.RoundingDecimals);

        var schedule = new List<ScheduleItem>();
        var previousDate = parameters.CreditStartDate;

        foreach (var paymentDate in paymentDates)
        {
            var daysInPeriod = (paymentDate - previousDate).Days;
            var interest = CalculateInterestDaily(previousDate, paymentDate, principalRemaining, parameters.MarginRate, ratePeriods, parameters.DayCountBasis);
            var interestRounded = RoundingService.Round(interest, parameters.RoundingMode, parameters.RoundingDecimals);

            decimal principalPayment;
            if (parameters.BulletRepayment && paymentDate == paymentDates[^1])
            {
                principalPayment = principalRemaining;
            }
            else if (parameters.BulletRepayment)
            {
                principalPayment = 0m;
            }
            else if (paymentDate == paymentDates[^1])
            {
                principalPayment = principalRemaining;
            }
            else
            {
                principalPayment = principalStep;
            }

            principalPayment = RoundingService.Round(principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            principalRemaining = RoundingService.Round(principalRemaining - principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            var totalPayment = interestRounded + principalPayment;

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
