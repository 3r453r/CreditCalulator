using System.Text;
using CreditTool.Models;

namespace CreditTool.Services;

public class LogExportService
{
    public byte[] Export(IEnumerable<CalculationLogEntry> logEntries)
    {
        var builder = new StringBuilder();

        builder.AppendLine("# 📊 Log obliczeń kredytu");
        builder.AppendLine();
        builder.AppendLine("Ten dokument zawiera szczegółowe obliczenia każdej raty kredytu.");
        builder.AppendLine("Każda rata jest osobno opisana z wszystkimi krokami obliczeń.");
        builder.AppendLine();

        var currentPayment = 0;
        var inPaymentSection = false;

        foreach (var entry in logEntries)
        {
            var entryType = entry.Context?.Type ?? LogEntryType.Detail;
            var paymentNumber = entry.Context?.PaymentNumber;

            // Handle section headers
            if (entryType == LogEntryType.Header && paymentNumber.HasValue)
            {
                if (inPaymentSection)
                {
                    builder.AppendLine();
                    builder.AppendLine("---");
                    builder.AppendLine();
                }

                currentPayment = paymentNumber.Value;
                inPaymentSection = true;

                builder.AppendLine($"## {entry.ShortDescription}");
                builder.AppendLine();
                builder.AppendLine($"**{entry.Result}**");
                builder.AppendLine();
                continue;
            }

            // Handle rate changes (global notifications)
            if (entryType == LogEntryType.RateChange)
            {
                builder.AppendLine($"### ℹ️ {entry.ShortDescription}");
                if (!string.IsNullOrEmpty(entry.SymbolicFormula))
                {
                    builder.AppendLine($"- **Formuła:** `{entry.SymbolicFormula}`");
                }
                if (!string.IsNullOrEmpty(entry.SubstitutedFormula))
                {
                    builder.AppendLine($"- **Obliczenie:** `{entry.SubstitutedFormula}`");
                }
                builder.AppendLine($"- **Wynik:** {entry.Result}");
                builder.AppendLine();
                continue;
            }

            // Handle summary entries (highlighted)
            if (entryType == LogEntryType.Summary)
            {
                builder.AppendLine($"### ✅ {entry.ShortDescription}");
                if (!string.IsNullOrEmpty(entry.SymbolicFormula))
                {
                    builder.AppendLine($"- **Formuła:** `{entry.SymbolicFormula}`");
                }
                if (!string.IsNullOrEmpty(entry.SubstitutedFormula))
                {
                    builder.AppendLine($"- **Obliczenie:** `{entry.SubstitutedFormula}`");
                }
                builder.AppendLine($"- **Wynik:** **{entry.Result}**");
                builder.AppendLine();
                continue;
            }

            // Handle regular calculation entries
            builder.AppendLine($"**{entry.ShortDescription}**");
            if (!string.IsNullOrEmpty(entry.SymbolicFormula))
            {
                builder.AppendLine($"- Formuła: `{entry.SymbolicFormula}`");
            }
            if (!string.IsNullOrEmpty(entry.SubstitutedFormula))
            {
                builder.AppendLine($"- Obliczenie: `{entry.SubstitutedFormula}`");
            }
            builder.AppendLine($"- Wynik: {entry.Result}");
            builder.AppendLine();
        }

        builder.AppendLine();
        builder.AppendLine("---");
        builder.AppendLine();
        builder.AppendLine("*Koniec logu obliczeń*");

        return Encoding.UTF8.GetBytes(builder.ToString());
    }
}
