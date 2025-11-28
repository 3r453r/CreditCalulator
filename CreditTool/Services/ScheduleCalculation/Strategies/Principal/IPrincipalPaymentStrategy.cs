namespace CreditTool.Services.ScheduleCalculation.Strategies.Principal;

/// <summary>
/// Strategy interface for determining principal payment amounts.
/// </summary>
public interface IPrincipalPaymentStrategy
{
    /// <summary>
    /// Calculates the principal payment for a given installment.
    /// </summary>
    /// <param name="context">Context containing all necessary information for the calculation.</param>
    /// <returns>The principal payment amount.</returns>
    decimal CalculatePrincipalPayment(PrincipalPaymentContext context);
}

/// <summary>
/// Context object containing all information needed to calculate principal payment.
/// </summary>
public record PrincipalPaymentContext(
    decimal RemainingPrincipal,
    decimal InterestAmount,
    int CurrentPaymentIndex,
    int TotalPayments,
    bool IsLastPayment,
    decimal? TargetTotalPayment = null);
