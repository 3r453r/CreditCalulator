namespace CreditTool.Services;

/// <summary>
/// Provides high-precision mathematical operations for decimal types.
/// Used to avoid precision loss from decimal-to-double conversions in financial calculations.
/// </summary>
public static class DecimalMath
{
    /// <summary>
    /// Calculates base raised to an integer power using decimal precision.
    /// Uses exponentiation by squaring for O(log n) efficiency.
    /// </summary>
    /// <param name="baseValue">The base value (must be decimal for precision)</param>
    /// <param name="exponent">The integer exponent (can be positive, negative, or zero)</param>
    /// <returns>The result of baseValue^exponent with full decimal precision</returns>
    /// <remarks>
    /// This method avoids the precision loss that occurs when converting decimal to double
    /// for Math.Pow operations. Critical for compound interest calculations where small
    /// errors accumulate over multiple periods (e.g., 31-day compounding periods).
    /// </remarks>
    public static decimal Power(decimal baseValue, int exponent)
    {
        if (exponent == 0)
            return 1m;

        if (exponent == 1)
            return baseValue;

        if (baseValue == 0m)
            return 0m;

        if (baseValue == 1m)
            return 1m;

        // Use exponentiation by squaring algorithm
        // For positive exponent: repeatedly square and multiply
        // For negative exponent: calculate positive then invert
        decimal result = 1m;
        decimal currentPower = baseValue;
        int exp = Math.Abs(exponent);

        while (exp > 0)
        {
            // If current bit is 1, multiply result by current power
            if ((exp & 1) == 1)
            {
                result *= currentPower;
            }

            // Square the current power for next iteration
            currentPower *= currentPower;

            // Shift to next bit
            exp >>= 1;
        }

        // Handle negative exponents by inverting
        return exponent < 0 ? 1m / result : result;
    }
}
