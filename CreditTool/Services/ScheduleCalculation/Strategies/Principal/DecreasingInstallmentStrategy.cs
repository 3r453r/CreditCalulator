namespace CreditTool.Services.ScheduleCalculation.Strategies.Principal;

/// <summary>
/// Decreasing installment strategy where equal principal amounts are paid each period.
/// Results in decreasing total payments as interest decreases over time.
/// </summary>
public class DecreasingInstallmentStrategy : IPrincipalPaymentStrategy
{
    private readonly decimal _initialPrincipal;
    private readonly decimal _principalStep;

    /// <summary>
    /// Creates a decreasing installment strategy.
    /// </summary>
    /// <param name="initialPrincipal">The initial loan principal.</param>
    /// <param name="principalStep">The fixed principal amount to pay each period.</param>
    public DecreasingInstallmentStrategy(decimal initialPrincipal, decimal principalStep)
    {
        _initialPrincipal = initialPrincipal;
        _principalStep = principalStep;
    }

    public decimal CalculatePrincipalPayment(PrincipalPaymentContext context)
    {
        // Grace period: no principal payment
        if (context.IsInGracePeriod)
        {
            return 0;
        }

        // Last payment: pay remaining balance to close the loan
        if (context.IsLastPayment)
        {
            return context.RemainingPrincipal;
        }

        // Regular payment: fixed principal step
        return _principalStep;
    }
}
