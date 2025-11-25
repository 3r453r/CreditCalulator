using CreditTool.Models;

namespace CreditTool.Services;

public static class RoundingService
{
    public static decimal Round(decimal value, RoundingModeOption mode, int decimals)
    {
        var midpointRounding = mode == RoundingModeOption.AwayFromZero
            ? MidpointRounding.AwayFromZero
            : MidpointRounding.ToEven;

        return Math.Round(value, decimals, midpointRounding);
    }
}
