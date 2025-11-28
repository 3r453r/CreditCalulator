namespace CreditTool.Services.ScheduleCalculation.Strategies.Principal;

/// <summary>
/// Bullet payment strategy where the entire principal is paid at maturity.
/// No principal payments are made during the loan term except for the final payment.
/// </summary>
public class BulletPaymentStrategy : IPrincipalPaymentStrategy
{
    public decimal CalculatePrincipalPayment(PrincipalPaymentContext context)
    {
        // For bullet loans, only pay principal on the last payment
        return context.IsLastPayment ? context.RemainingPrincipal : 0m;
    }
}
