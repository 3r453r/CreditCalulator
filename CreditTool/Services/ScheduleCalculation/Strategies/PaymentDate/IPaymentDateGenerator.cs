using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation.Strategies.PaymentDate;

/// <summary>
/// Interface for generating payment dates based on loan parameters.
/// </summary>
public interface IPaymentDateGenerator
{
    /// <summary>
    /// Generates a list of payment dates between the start and end dates.
    /// </summary>
    /// <param name="startDate">The loan start date.</param>
    /// <param name="endDate">The loan end date.</param>
    /// <param name="frequency">The payment frequency.</param>
    /// <param name="paymentDay">The day of month option for payments.</param>
    /// <returns>A list of payment dates.</returns>
    List<DateTime> GeneratePaymentDates(
        DateTime startDate,
        DateTime endDate,
        PaymentFrequency frequency,
        PaymentDayOption paymentDay);
}
