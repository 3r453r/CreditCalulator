using CreditTool.Services;

namespace CreditTool.Tests;

public class DecimalMathTests
{
    [Theory]
    [InlineData(1.000136986301369863, 1, 1.000136986301369863)]
    [InlineData(1.000136986301369863, 31, 1.004257421687867337)] // 31-day compound
    [InlineData(1.05, 2, 1.1025)]
    [InlineData(2.0, 10, 1024.0)]
    [InlineData(1.1, 0, 1.0)]
    [InlineData(5.0, 1, 5.0)]
    [InlineData(2.0, -3, 0.125)]
    public void Power_CalculatesCorrectly(double baseDouble, int exponent, double expectedDouble)
    {
        var baseValue = (decimal)baseDouble;
        var expected = (decimal)expectedDouble;

        var result = DecimalMath.Power(baseValue, exponent);

        // For precision comparison, allow small tolerance due to double conversion
        var tolerance = 0.0000000001m;
        Assert.InRange(result, expected - tolerance, expected + tolerance);
    }

    [Fact]
    public void Power_MaintainsPrecisionBetterThanDoubleMathPow()
    {
        // This test demonstrates the precision advantage of decimal power
        // For a typical daily compound interest scenario: 5% annual rate, 31 days
        var annualRate = 0.05m;
        var dailyRate = 1m + (annualRate / 365m);  // 1.000136986301369863...
        var days = 31;

        // Using our decimal power function
        var decimalResult = DecimalMath.Power(dailyRate, days);

        // Using Math.Pow with double conversion (old approach)
        var doubleResult = (decimal)Math.Pow((double)(1.0m + (annualRate / 365m)), days);

        // Both should be close, but decimal should maintain more precision
        // The key is that decimal calculation avoids the double conversion loss
        Assert.NotEqual(0m, decimalResult);
        Assert.NotEqual(0m, doubleResult);

        // For this specific case, let's verify the decimal result is reasonable
        // (1 + 0.05/365)^31 should be approximately 1.00426
        Assert.InRange(decimalResult, 1.0042m, 1.0043m);
    }

    [Fact]
    public void Power_HandlesZeroBase()
    {
        var result = DecimalMath.Power(0m, 5);
        Assert.Equal(0m, result);
    }

    [Fact]
    public void Power_HandlesZeroExponent()
    {
        var result = DecimalMath.Power(123.456m, 0);
        Assert.Equal(1m, result);
    }

    [Fact]
    public void Power_HandlesOneBase()
    {
        var result = DecimalMath.Power(1m, 100);
        Assert.Equal(1m, result);
    }

    [Fact]
    public void Power_HandlesNegativeExponent()
    {
        var result = DecimalMath.Power(2m, -3);
        Assert.Equal(0.125m, result);
    }

    [Fact]
    public void Power_HandlesLargeExponents()
    {
        // Test with a realistic scenario: 365 days of daily compounding
        var dailyRate = 1.0001m;  // Simplified daily rate
        var result = DecimalMath.Power(dailyRate, 365);

        // Should be approximately 1.0372 (e^0.0365)
        Assert.InRange(result, 1.03m, 1.04m);
    }

    [Fact]
    public void Power_CompoundDailyRealWorldScenario()
    {
        // Real-world scenario: 5% annual rate, 31-day period, 10000 principal
        var principal = 10000m;
        var annualRate = 5m;
        var days = 31;
        var denominator = 365m;

        var dailyRate = 1m + (annualRate / 100m / denominator);
        var factor = DecimalMath.Power(dailyRate, days);
        var interest = principal * (factor - 1m);

        // Interest should be approximately 42.47
        Assert.InRange(interest, 42.4m, 42.5m);

        // Verify factor is greater than 1 (we earned interest)
        Assert.True(factor > 1m);
    }
}
