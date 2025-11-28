using System.Text;
using CreditTool.Models;

namespace CreditTool.Services;

public class LogExportService
{
    public byte[] Export(IEnumerable<CalculationLogEntry> logEntries)
    {
        var builder = new StringBuilder();

        builder.AppendLine("# Log obliczeń kredytu");
        builder.AppendLine();
        builder.AppendLine("## Szczegóły kroków");
        builder.AppendLine();

        var index = 1;
        foreach (var entry in logEntries)
        {
            builder.AppendLine($"### {index}. {entry.ShortDescription}");
            builder.AppendLine($"- Wzór: {entry.SymbolicFormula}");
            builder.AppendLine($"- Wzór z danymi: {entry.SubstitutedFormula}");
            builder.AppendLine($"- Wynik: {entry.Result}");
            builder.AppendLine();
            index++;
        }

        return Encoding.UTF8.GetBytes(builder.ToString());
    }
}
