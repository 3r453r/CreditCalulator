namespace CreditTool.Models;

public class InterestRatePeriod
{
    public DateTime DateFrom { get; set; }

    public DateTime DateTo { get; set; }

    /// <summary>
    /// Percentage value of the base rate for the period.
    /// </summary>
    public decimal Rate { get; set; }
}
