namespace CreditTool.Services.ScheduleCalculation.Strategies.Principal;

/// <summary>
/// Annuity (equal installment) strategy where total payments remain constant.
/// Principal payment equals the target payment minus interest.
/// </summary>
public class AnnuityStrategy : IPrincipalPaymentStrategy
{
    public decimal CalculatePrincipalPayment(PrincipalPaymentContext context)
    {
        if (context.TargetTotalPayment == null)
        {
            throw new InvalidOperationException("Target total payment must be provided for annuity strategy.");
        }

        // Last payment: pay remaining balance exactly to close the loan
        if (context.IsLastPayment)
        {
            return context.RemainingPrincipal;
        }

        // Regular payment: target total minus interest
        var principalPayment = context.TargetTotalPayment.Value - context.InterestAmount;

        // Ensure we don't pay more than the remaining balance
        principalPayment = Math.Min(principalPayment, context.RemainingPrincipal);

        // Ensure non-negative principal payment
        return Math.Max(principalPayment, 0m);
    }
}
