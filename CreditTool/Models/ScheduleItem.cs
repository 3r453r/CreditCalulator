namespace CreditTool.Models;

public class ScheduleItem
{
    public DateTime PaymentDate { get; set; }

    public int DaysInPeriod { get; set; }

    public decimal InterestRate { get; set; }

    public decimal InterestAmount { get; set; }

    public decimal PrincipalPayment { get; set; }

    public decimal TotalPayment { get; set; }

    public decimal RemainingPrincipal { get; set; }
}
