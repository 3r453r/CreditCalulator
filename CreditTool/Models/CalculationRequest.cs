namespace CreditTool.Models;

public class CalculationRequest
{
    public CreditParameters Parameters { get; set; } = new();

    public List<InterestRatePeriod> Rates { get; set; } = new();
}
