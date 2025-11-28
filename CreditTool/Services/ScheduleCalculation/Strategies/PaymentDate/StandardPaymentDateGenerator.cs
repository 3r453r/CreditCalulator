using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.PaymentDate;

/// <summary>
/// Standard implementation of payment date generation.
/// Generates dates based on frequency and payment day options.
/// </summary>
public class StandardPaymentDateGenerator : IPaymentDateGenerator
{
    public List<DateTime> GeneratePaymentDates(
        DateTime startDate,
        DateTime endDate,
        PaymentFrequency frequency,
        PaymentDayOption paymentDay)
    {
        var dates = new List<DateTime>();
        var current = startDate;

        while (current < endDate)
        {
            DateTime next = frequency switch
            {
                PaymentFrequency.Daily => current.AddDays(1),
                PaymentFrequency.Monthly => NextMonthDate(current, 1, paymentDay),
                PaymentFrequency.Quarterly => NextMonthDate(current, 3, paymentDay),
                _ => current.AddMonths(1)
            };

            if (next > endDate)
            {
                next = endDate;
            }

            dates.Add(next);
            current = next;
        }

        // Ensure the end date is included if not already
        if (dates.Count == 0 || dates[^1] != endDate)
        {
            dates.Add(endDate);
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
