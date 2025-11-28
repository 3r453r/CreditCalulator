namespace CreditTool.Models;

public class CalculationLogEntry
{
    public string ShortDescription { get; set; } = string.Empty;

    public string SymbolicFormula { get; set; } = string.Empty;

    public string SubstitutedFormula { get; set; } = string.Empty;

    public string Result { get; set; } = string.Empty;
}
