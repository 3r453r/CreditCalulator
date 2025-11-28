using System.Linq;
using CreditTool.Models;
using CreditTool.Services;

namespace CreditTool.Tests;

public class ScheduleCalculatorTests
{
    private static DayToDayScheduleCalculator CreateCalculator() => new();

    [Fact]
    public void CalculatesMonthlyInterestWithBankingRounding()
    {
        var calculator = CreateCalculator();
        var parameters = new CreditParameters
        {
            NetValue = 10000m,
            MarginRate = 1m,
            PaymentFrequency = PaymentFrequency.Monthly,
            PaymentDay = PaymentDayOption.LastOfMonth,
            CreditStartDate = new DateTime(2024, 1, 1),
            CreditEndDate = new DateTime(2024, 2, 1),
            DayCountBasis = DayCountBasis.Actual365,
            RoundingMode = RoundingModeOption.Bankers,
            RoundingDecimals = 4
        };

        var schedule = calculator.Calculate(parameters, new[]
        {
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2024, 1, 1),
                DateTo = new DateTime(2024, 12, 31),
                Rate = 4m
            }
        }).Schedule;

        Assert.Single(schedule);
        var payment = schedule[0];
        Assert.Equal(31, payment.DaysInPeriod);
        Assert.Equal(5m, payment.InterestRate);
        var expectedInterest = RoundingService.Round(10000m * 0.05m / 365m * 31, parameters.RoundingMode, parameters.RoundingDecimals);
        Assert.Equal(expectedInterest, payment.InterestAmount);
        Assert.Equal(10000m, payment.PrincipalPayment);
        Assert.Equal(0m, payment.RemainingPrincipal);
    }

    [Fact]
    public void HandlesDailyScheduleWithRateChange()
    {
        var calculator = CreateCalculator();
        var parameters = new CreditParameters
        {
            NetValue = 5000m,
            MarginRate = 0m,
            PaymentFrequency = PaymentFrequency.Daily,
            PaymentDay = PaymentDayOption.LastOfMonth,
            CreditStartDate = new DateTime(2024, 3, 30),
            CreditEndDate = new DateTime(2024, 4, 2),
            DayCountBasis = DayCountBasis.Actual365,
            RoundingMode = RoundingModeOption.Bankers,
            RoundingDecimals = 4,
            BulletRepayment = true
        };

        var rates = new[]
        {
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2024, 3, 1),
                DateTo = new DateTime(2024, 3, 31),
                Rate = 3m
            },
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2024, 4, 1),
                DateTo = new DateTime(2024, 4, 30),
                Rate = 6m
            }
        };

        var schedule = calculator.Calculate(parameters, rates).Schedule;

        Assert.Equal(3, schedule.Count);
        Assert.All(schedule.Take(2), payment => Assert.Equal(0m, payment.PrincipalPayment));
        Assert.Equal(5000m, schedule.Last().PrincipalPayment);
        Assert.Equal(0m, schedule.Last().RemainingPrincipal);

        var firstDayInterest = RoundingService.Round(5000m * 0.03m / 365m, parameters.RoundingMode, parameters.RoundingDecimals);
        var secondDayInterest = RoundingService.Round(5000m * 0.03m / 365m, parameters.RoundingMode, parameters.RoundingDecimals);
        var thirdDayInterest = RoundingService.Round(5000m * 0.06m / 365m, parameters.RoundingMode, parameters.RoundingDecimals);
        var totalInterest = firstDayInterest + secondDayInterest + thirdDayInterest;
        Assert.Equal(totalInterest, schedule.Sum(p => p.InterestAmount));
    }

    [Fact]
    public void AppliesAwayFromZeroRoundingWhenSelected()
    {
        var calculator = CreateCalculator();
        var parameters = new CreditParameters
        {
            NetValue = 1234.56m,
            MarginRate = 2m,
            PaymentFrequency = PaymentFrequency.Monthly,
            PaymentDay = PaymentDayOption.FirstOfMonth,
            CreditStartDate = new DateTime(2024, 5, 15),
            CreditEndDate = new DateTime(2024, 6, 15),
            DayCountBasis = DayCountBasis.Actual360,
            RoundingMode = RoundingModeOption.AwayFromZero,
            RoundingDecimals = 4
        };

        var rates = new[]
        {
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2024, 1, 1),
                DateTo = new DateTime(2024, 12, 31),
                Rate = 1.3333m
            }
        };

        var schedule = calculator.Calculate(parameters, rates).Schedule;

        Assert.Equal(2, schedule.Count);
        var totalRate = rates[0].Rate + parameters.MarginRate;

        var firstDays = (new DateTime(2024, 6, 1) - parameters.CreditStartDate).Days;
        var firstRaw = parameters.NetValue * totalRate / 100m / 360m * firstDays;
        var firstRounded = RoundingService.Round(firstRaw, parameters.RoundingMode, parameters.RoundingDecimals);
        Assert.Equal(firstRounded, schedule[0].InterestAmount);

        var secondDays = (parameters.CreditEndDate - new DateTime(2024, 6, 1)).Days;
        var remainingPrincipal = parameters.NetValue - schedule[0].PrincipalPayment;
        var secondRaw = remainingPrincipal * totalRate / 100m / 360m * secondDays;
        var secondRounded = RoundingService.Round(secondRaw, parameters.RoundingMode, parameters.RoundingDecimals);
        Assert.Equal(secondRounded, schedule[1].InterestAmount);
    }

    [Fact]
    public void CalculatesEqualInstallmentsWithoutNegativePrincipal()
    {
        var calculator = CreateCalculator();
        var parameters = new CreditParameters
        {
            NetValue = 100_000m,
            MarginRate = 0m,
            PaymentFrequency = PaymentFrequency.Monthly,
            PaymentDay = PaymentDayOption.LastOfMonth,
            CreditStartDate = new DateTime(2025, 1, 1),
            CreditEndDate = new DateTime(2026, 1, 1),
            DayCountBasis = DayCountBasis.Actual365,
            RoundingMode = RoundingModeOption.Bankers,
            RoundingDecimals = 4,
            PaymentType = PaymentType.EqualInstallments
        };

        var rates = new[]
        {
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2025, 1, 1),
                DateTo = new DateTime(2025, 12, 31),
                Rate = 5m
            }
        };

        var schedule = calculator.Calculate(parameters, rates).Schedule;

        Assert.Equal(12, schedule.Count);
        Assert.All(schedule, item =>
        {
            Assert.True(item.InterestAmount >= 0m, "Interest should not be negative");
            Assert.True(item.PrincipalPayment >= 0m, "Principal should not be negative");
            Assert.True(item.RemainingPrincipal >= 0m, "Remaining principal should not be negative");
        });

        var firstTotal = schedule.First().TotalPayment;
        Assert.All(schedule.Take(schedule.Count - 1), item =>
            Assert.InRange(Math.Abs(item.TotalPayment - firstTotal), 0m, 0.05m));

        var remaining = schedule.Select(item => item.RemainingPrincipal).ToList();
        for (var i = 1; i < remaining.Count; i++)
        {
            Assert.True(remaining[i] <= remaining[i - 1] + 0.0001m, "Remaining principal should decrease over time");
        }

        Assert.InRange(schedule.First().PrincipalPayment, 0.01m, parameters.NetValue - 0.01m);
        Assert.True(schedule.First().PrincipalPayment < schedule.Last().PrincipalPayment,
            "For equal installments, principal part should grow over time");
        Assert.Equal(0m, schedule.Last().RemainingPrincipal);
    }

    [Fact]
    public void DifferentiatesAccrualMethodsWhenRateChangesMidPeriod()
    {
        var calculator = CreateCalculator();

        var startDate = new DateTime(2025, 1, 1);
        var endDate = new DateTime(2025, 4, 1);

        const decimal netValue = 100_000m;
        const decimal marginRate = 1m;

        CreditParameters CreateParameters(InterestRateApplication application) => new()
        {
            NetValue = netValue,
            MarginRate = marginRate,
            PaymentFrequency = PaymentFrequency.Monthly,
            PaymentDay = PaymentDayOption.LastOfMonth,
            CreditStartDate = startDate,
            CreditEndDate = endDate,
            DayCountBasis = DayCountBasis.Actual365,
            RoundingMode = RoundingModeOption.Bankers,
            RoundingDecimals = 4,
            PaymentType = PaymentType.DecreasingInstallments,
            InterestRateApplication = application
        };

        var ratePeriods = new[]
        {
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2025, 1, 1),
                DateTo = new DateTime(2025, 2, 14),
                Rate = 10m
            },
            new InterestRatePeriod
            {
                DateFrom = new DateTime(2025, 2, 15),
                DateTo = new DateTime(2025, 12, 31),
                Rate = 3m
            }
        };

        var dailySchedule = calculator.Calculate(CreateParameters(InterestRateApplication.DailyAccrual), ratePeriods).Schedule;
        var nextPeriodSchedule = calculator.Calculate(CreateParameters(InterestRateApplication.ApplyChangedRateNextPeriod), ratePeriods).Schedule;

        Assert.Equal(dailySchedule.Count, nextPeriodSchedule.Count);

        var firstPeriodDays = dailySchedule[0].DaysInPeriod;
        var daysAtHighRate = (new DateTime(2025, 2, 15) - startDate).Days;
        var daysAtLowRate = firstPeriodDays - daysAtHighRate;

        var expectedFirstPeriodDailyRaw = netValue * (10m + marginRate) / 100m / 365m * daysAtHighRate
                                         + netValue * (3m + marginRate) / 100m / 365m * daysAtLowRate;
        var expectedFirstPeriodDaily = RoundingService.Round(expectedFirstPeriodDailyRaw, RoundingModeOption.Bankers, 4);
        Assert.Equal(expectedFirstPeriodDaily, dailySchedule[0].InterestAmount);

        var expectedFirstPeriodNextRaw = netValue * (10m + marginRate) / 100m / 365m * firstPeriodDays;
        var expectedFirstPeriodNext = RoundingService.Round(expectedFirstPeriodNextRaw, RoundingModeOption.Bankers, 4);

        Assert.Equal(expectedFirstPeriodNext, nextPeriodSchedule[0].InterestAmount);
        Assert.NotEqual(nextPeriodSchedule[0].InterestAmount, dailySchedule[0].InterestAmount);
    }
}
