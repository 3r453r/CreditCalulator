using CreditTool.Models;

namespace CreditTool.Services.ScheduleCalculation;

/// <summary>
/// Interface for calculating amortization schedules based on credit parameters and rate periods.
/// </summary>
public interface IScheduleCalculator
{
    /// <summary>
    /// Calculates a full amortization schedule for the given credit parameters.
    /// </summary>
    /// <param name="parameters">The credit parameters defining the loan structure.</param>
    /// <param name="ratePeriods">Collection of interest rate periods over the loan duration.</param>
    /// <param name="configuration">Optional configuration for calculation behavior.</param>
    /// <param name="includeLog">Whether to include detailed calculation logs.</param>
    /// <returns>The complete schedule calculation result including items, logs, and warnings.</returns>
    ScheduleCalculationResult Calculate(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        CalculatorConfiguration? configuration = null,
        bool includeLog = false);
}
