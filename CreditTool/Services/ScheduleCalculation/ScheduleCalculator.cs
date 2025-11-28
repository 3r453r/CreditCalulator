using CreditTool.Models;
using CreditTool.Services.ScheduleCalculation.Configuration;
using CreditTool.Services.ScheduleCalculation.Strategies.Interest;
using CreditTool.Services.ScheduleCalculation.Strategies.PaymentDate;
using CreditTool.Services.ScheduleCalculation.Strategies.Principal;

namespace CreditTool.Services.ScheduleCalculation;

/// <summary>
/// Main orchestrator for schedule calculations using strategy pattern.
/// Coordinates payment date generation, interest calculation, and principal payment strategies.
/// </summary>
public class ScheduleCalculator : IScheduleCalculator
{
    private readonly IPaymentDateGenerator _paymentDateGenerator;

    public ScheduleCalculator(IPaymentDateGenerator paymentDateGenerator)
    {
        _paymentDateGenerator = paymentDateGenerator;
    }

    public ScheduleCalculationResult Calculate(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        CalculatorConfiguration? configuration = null,
        bool includeLog = false)
    {
        configuration ??= new CalculatorConfiguration();

        if (parameters.CreditEndDate <= parameters.CreditStartDate)
        {
            throw new ArgumentException("Credit end date must be after start date");
        }

        var paymentDates = _paymentDateGenerator.GeneratePaymentDates(
            parameters.CreditStartDate,
            parameters.CreditEndDate,
            parameters.PaymentFrequency,
            parameters.PaymentDay);

        var interestStrategy = CreateInterestStrategy(parameters.InterestRateApplication);
        var principalStrategy = CreatePrincipalStrategy(parameters, paymentDates.Count);

        decimal? fixedTotalPayment = null;
        if (parameters.PaymentType == PaymentType.EqualInstallments)
        {
            fixedTotalPayment = CalculateLevelPayment(
                parameters,
                ratePeriods,
                paymentDates,
                configuration,
                interestStrategy);
        }

        return GenerateSchedule(
            parameters,
            ratePeriods,
            paymentDates,
            interestStrategy,
            principalStrategy,
            configuration,
            fixedTotalPayment,
            includeLog);
    }

    private IInterestCalculationStrategy CreateInterestStrategy(InterestRateApplication application)
    {
        return application switch
        {
            InterestRateApplication.ApplyChangedRateNextPeriod => new ApplyRateNextPeriodStrategy(),
            InterestRateApplication.CompoundDaily => new CompoundDailyStrategy(),
            InterestRateApplication.CompoundMonthly => new CompoundPeriodicStrategy(12),
            InterestRateApplication.CompoundQuarterly => new CompoundPeriodicStrategy(4),
            _ => new SimpleInterestStrategy()
        };
    }

    private IPrincipalPaymentStrategy CreatePrincipalStrategy(CreditParameters parameters, int paymentCount)
    {
        return parameters.PaymentType switch
        {
            PaymentType.Bullet => new BulletPaymentStrategy(),
            PaymentType.DecreasingInstallments => new DecreasingInstallmentStrategy(
                parameters.NetValue,
                RoundingService.Round(
                    parameters.NetValue / paymentCount,
                    parameters.RoundingMode,
                    parameters.RoundingDecimals)),
            PaymentType.EqualInstallments => new AnnuityStrategy(),
            _ => throw new ArgumentException($"Unsupported payment type: {parameters.PaymentType}")
        };
    }

    private ScheduleCalculationResult GenerateSchedule(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        List<DateTime> paymentDates,
        IInterestCalculationStrategy interestStrategy,
        IPrincipalPaymentStrategy principalStrategy,
        CalculatorConfiguration configuration,
        decimal? fixedTotalPayment,
        bool includeLog)
    {
        var schedule = new List<ScheduleItem>();
        var calculationLog = new List<CalculationLogEntry>();
        var warnings = new List<string>();
        var principalRemaining = parameters.NetValue;
        var previousDate = parameters.CreditStartDate;

        if (includeLog)
        {
            LogRateChanges(calculationLog, parameters, ratePeriods);
        }

        for (var index = 0; index < paymentDates.Count; index++)
        {
            var paymentDate = paymentDates[index];
            var isLastPayment = index == paymentDates.Count - 1;
            var daysInPeriod = (paymentDate - previousDate).Days;

            if (includeLog)
            {
                LogPaymentHeader(calculationLog, index + 1, previousDate, paymentDate, principalRemaining);
                LogDaysInPeriod(calculationLog, index + 1, paymentDate, previousDate, daysInPeriod);
            }

            var interestResult = interestStrategy.Calculate(
                previousDate,
                paymentDate,
                principalRemaining,
                parameters.MarginRate,
                ratePeriods,
                parameters.DayCountBasis);

            var interestRounded = RoundingService.Round(
                interestResult.Interest,
                parameters.RoundingMode,
                parameters.RoundingDecimals);

            if (includeLog)
            {
                LogInterestCalculation(
                    calculationLog,
                    index + 1,
                    paymentDate,
                    parameters,
                    interestResult,
                    interestRounded,
                    principalRemaining,
                    daysInPeriod);
            }

            var principalContext = new PrincipalPaymentContext(
                principalRemaining,
                interestRounded,
                index,
                paymentDates.Count,
                isLastPayment,
                fixedTotalPayment);

            var principalPaymentRaw = principalStrategy.CalculatePrincipalPayment(principalContext);

            if (includeLog)
            {
                LogPrincipalCalculation(
                    calculationLog,
                    index + 1,
                    paymentDate,
                    parameters,
                    principalPaymentRaw,
                    principalRemaining,
                    fixedTotalPayment,
                    interestRounded,
                    isLastPayment);
            }

            var principalPayment = RoundingService.Round(
                principalPaymentRaw,
                parameters.RoundingMode,
                parameters.RoundingDecimals);

            var itemWarnings = ValidatePayment(
                parameters,
                configuration,
                warnings,
                index,
                principalPaymentRaw,
                interestRounded,
                fixedTotalPayment);

            principalRemaining = RoundingService.Round(
                principalRemaining - principalPayment,
                parameters.RoundingMode,
                parameters.RoundingDecimals);

            if (includeLog)
            {
                LogRemainingBalance(calculationLog, index + 1, paymentDate, principalRemaining, principalPayment, interestRounded, principalPayment + interestRounded);
            }

            var totalPayment = interestRounded + principalPayment;
            var isFinalAdjusted = false;

            if (parameters.PaymentType == PaymentType.EqualInstallments &&
                isLastPayment &&
                fixedTotalPayment.HasValue &&
                Math.Abs(totalPayment - fixedTotalPayment.Value) > 0.01m)
            {
                isFinalAdjusted = true;
                itemWarnings |= ScheduleWarnings.FinalPaymentAdjusted;
            }

            schedule.Add(new ScheduleItem
            {
                PaymentDate = paymentDate,
                DaysInPeriod = daysInPeriod,
                InterestRate = interestResult.EffectiveRate,
                NominalRate = interestResult.NominalRate,
                EffectivePeriodRate = interestResult.EffectivePeriodRate,
                InterestAmount = interestRounded,
                PrincipalPayment = principalPayment,
                TotalPayment = totalPayment,
                RemainingPrincipal = Math.Max(principalRemaining, 0m),
                IsFinalPaymentAdjusted = isFinalAdjusted,
                Warnings = itemWarnings
            });

            previousDate = paymentDate;
        }

        return new ScheduleCalculationResult
        {
            Schedule = schedule,
            CalculationLog = calculationLog,
            Warnings = warnings,
            TargetLevelPayment = fixedTotalPayment,
            ActualFinalPayment = schedule.Count > 0 ? schedule[^1].TotalPayment : null
        };
    }

    private decimal CalculateLevelPayment(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates,
        CalculatorConfiguration configuration,
        IInterestCalculationStrategy interestStrategy)
    {
        var low = 0m;
        var high = Math.Max(parameters.NetValue, parameters.NetValue + 1000m);
        var tolerance = configuration.LevelPaymentTolerance;

        var remainingHigh = SimulateRemainingPrincipal(
            parameters,
            ratePeriods,
            paymentDates,
            high,
            interestStrategy);

        var guard = 0;
        while (remainingHigh > 0m && guard < configuration.MaxRangeFindingIterations)
        {
            low = high;
            high *= 2m;
            remainingHigh = SimulateRemainingPrincipal(
                parameters,
                ratePeriods,
                paymentDates,
                high,
                interestStrategy);
            guard++;
        }

        for (var i = 0; i < configuration.MaxLevelPaymentIterations; i++)
        {
            var mid = (low + high) / 2m;
            var remaining = SimulateRemainingPrincipal(
                parameters,
                ratePeriods,
                paymentDates,
                mid,
                interestStrategy);

            if (Math.Abs(high - low) <= tolerance)
            {
                break;
            }

            if (remaining > 0)
            {
                low = mid;
            }
            else
            {
                high = mid;
            }
        }

        return RoundingService.Round(
            (low + high) / 2m,
            parameters.RoundingMode,
            parameters.RoundingDecimals);
    }

    private decimal SimulateRemainingPrincipal(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates,
        decimal totalPayment,
        IInterestCalculationStrategy interestStrategy)
    {
        var principalRemaining = parameters.NetValue;
        var previousDate = parameters.CreditStartDate;

        foreach (var paymentDate in paymentDates)
        {
            var result = interestStrategy.Calculate(
                previousDate,
                paymentDate,
                principalRemaining,
                parameters.MarginRate,
                ratePeriods,
                parameters.DayCountBasis);

            var interestRounded = RoundingService.Round(
                result.Interest,
                parameters.RoundingMode,
                parameters.RoundingDecimals);

            var principalPayment = totalPayment - interestRounded;
            if (principalPayment < 0m)
            {
                principalPayment = 0m;
            }

            principalPayment = Math.Min(principalPayment, principalRemaining);
            principalPayment = RoundingService.Round(
                principalPayment,
                parameters.RoundingMode,
                parameters.RoundingDecimals);
            principalRemaining = RoundingService.Round(
                principalRemaining - principalPayment,
                parameters.RoundingMode,
                parameters.RoundingDecimals);
            previousDate = paymentDate;
        }

        return principalRemaining;
    }

    private ScheduleWarnings ValidatePayment(
        CreditParameters parameters,
        CalculatorConfiguration configuration,
        List<string> warnings,
        int index,
        decimal principalPaymentRaw,
        decimal interestRounded,
        decimal? fixedTotalPayment)
    {
        var itemWarnings = ScheduleWarnings.None;

        if (parameters.PaymentType != PaymentType.EqualInstallments || !fixedTotalPayment.HasValue)
        {
            return itemWarnings;
        }

        if (interestRounded > fixedTotalPayment.Value)
        {
            itemWarnings |= ScheduleWarnings.InterestExceedsPayment;
        }

        if (principalPaymentRaw < 0m)
        {
            itemWarnings |= ScheduleWarnings.NegativeAmortization;

            if (configuration.EnableValidation)
            {
                var warning = $"Rata {index + 1}: Odsetki ({interestRounded:F2}) przekraczają ratę docelową ({fixedTotalPayment.Value:F2})";
                warnings.Add(warning);

                if (configuration.ThrowOnNegativeAmortization)
                {
                    throw new InvalidOperationException(warning);
                }
            }
        }

        return itemWarnings;
    }

    #region Logging Methods

    private void LogRateChanges(
        List<CalculationLogEntry> log,
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods)
    {
        var recalculatedRatePeriods = ratePeriods
            .Where(period => period.DateFrom > parameters.CreditStartDate &&
                           period.DateFrom < parameters.CreditEndDate)
            .OrderBy(period => period.DateFrom);

        foreach (var period in recalculatedRatePeriods)
        {
            log.Add(new CalculationLogEntry
            {
                ShortDescription = "Zmiana stopy procentowej",
                SymbolicFormula = "stopa_efektywna = stawka_bazowa + marża",
                SubstitutedFormula = $"{period.Rate:F4}% + {parameters.MarginRate:F4}%",
                Result = $"{(period.Rate + parameters.MarginRate):F4}% (obowiązuje od {period.DateFrom:yyyy-MM-dd} do {period.DateTo:yyyy-MM-dd})",
                Context = new LogEntryContext
                {
                    Type = LogEntryType.RateChange,
                    Metadata = new Dictionary<string, string>
                    {
                        ["BaseRate"] = $"{period.Rate:F4}%",
                        ["MarginRate"] = $"{parameters.MarginRate:F4}%",
                        ["DateFrom"] = period.DateFrom.ToString("yyyy-MM-dd"),
                        ["DateTo"] = period.DateTo.ToString("yyyy-MM-dd")
                    }
                }
            });
        }
    }

    private void LogPaymentHeader(
        List<CalculationLogEntry> log,
        int paymentNumber,
        DateTime periodStart,
        DateTime periodEnd,
        decimal principalBefore)
    {
        log.Add(new CalculationLogEntry
        {
            ShortDescription = $"═══ RATA {paymentNumber} ({periodStart:yyyy-MM-dd} → {periodEnd:yyyy-MM-dd}) ═══",
            SymbolicFormula = "",
            SubstitutedFormula = "",
            Result = $"Saldo początkowe: {principalBefore:F2} PLN",
            Context = new LogEntryContext
            {
                PaymentNumber = paymentNumber,
                PaymentDate = periodEnd,
                Type = LogEntryType.Header,
                Metadata = new Dictionary<string, string>
                {
                    ["PeriodStart"] = periodStart.ToString("yyyy-MM-dd"),
                    ["PeriodEnd"] = periodEnd.ToString("yyyy-MM-dd"),
                    ["PrincipalBefore"] = $"{principalBefore:F2}"
                }
            }
        });
    }

    private void LogDaysInPeriod(
        List<CalculationLogEntry> log,
        int paymentNumber,
        DateTime paymentDate,
        DateTime previousDate,
        int daysInPeriod)
    {
        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Liczba dni w okresie",
            SymbolicFormula = "dni = data_końcowa - data_początkowa",
            SubstitutedFormula = $"{paymentDate:yyyy-MM-dd} - {previousDate:yyyy-MM-dd}",
            Result = $"{daysInPeriod} dni",
            Context = new LogEntryContext
            {
                PaymentNumber = paymentNumber,
                PaymentDate = paymentDate,
                Type = LogEntryType.PeriodCalculation
            }
        });
    }

    private void LogInterestCalculation(
        List<CalculationLogEntry> log,
        int paymentNumber,
        DateTime paymentDate,
        CreditParameters parameters,
        InterestCalculationResult result,
        decimal interestRounded,
        decimal principal,
        int days)
    {
        var denominator = parameters.DayCountBasis == DayCountBasis.Actual360 ? 360m : 365m;

        var interestDescription = parameters.InterestRateApplication switch
        {
            InterestRateApplication.ApplyChangedRateNextPeriod =>
                "Odsetki (stopa z początku okresu)",
            InterestRateApplication.CompoundDaily =>
                "Odsetki (kapitalizacja dzienna)",
            InterestRateApplication.CompoundMonthly =>
                "Odsetki (kapitalizacja miesięczna)",
            InterestRateApplication.CompoundQuarterly =>
                "Odsetki (kapitalizacja kwartalna)",
            _ => "Odsetki (naliczanie dzienne)"
        };

        var interestFormula = parameters.InterestRateApplication switch
        {
            InterestRateApplication.ApplyChangedRateNextPeriod =>
                "odsetki = saldo × (stopa_początkowa / 100 / baza) × dni",
            InterestRateApplication.CompoundDaily =>
                "odsetki = saldo × (∏(1 + r_dzienne/100/baza) - 1)",
            InterestRateApplication.CompoundMonthly =>
                "odsetki = saldo × ((1 + r_nom/100/12)^(12×t) - 1)",
            InterestRateApplication.CompoundQuarterly =>
                "odsetki = saldo × ((1 + r_nom/100/4)^(4×t) - 1)",
            _ => "odsetki = saldo × średnia_stopa_dzienna × dni"
        };

        string interestSubstitution;
        string calculationDetails;
        if (result.NominalRate.HasValue && result.EffectivePeriodRate.HasValue)
        {
            interestSubstitution = $"{principal:F2} × {result.EffectivePeriodRate.Value:F6}";
            calculationDetails = $"(stopa nom.: {result.NominalRate.Value:F4}%, dni: {days}, baza: {denominator})";
        }
        else
        {
            interestSubstitution = $"{principal:F2} × średnia_dzienna";
            calculationDetails = $"(stopa śr.: {result.EffectiveRate:F4}%, dni: {days}, baza: {denominator})";
        }

        var rawInterest = result.Interest;
        var roundingNeeded = Math.Abs(rawInterest - interestRounded) > 0.000001m;
        var resultText = roundingNeeded
            ? $"{interestRounded:F2} PLN (przed zaokr.: {rawInterest:F6})"
            : $"{interestRounded:F2} PLN";

        log.Add(new CalculationLogEntry
        {
            ShortDescription = interestDescription,
            SymbolicFormula = interestFormula,
            SubstitutedFormula = $"{interestSubstitution} {calculationDetails}",
            Result = resultText,
            Context = new LogEntryContext
            {
                PaymentNumber = paymentNumber,
                PaymentDate = paymentDate,
                Type = LogEntryType.InterestCalculation,
                Metadata = new Dictionary<string, string>
                {
                    ["Principal"] = $"{principal:F2}",
                    ["EffectiveRate"] = $"{result.EffectiveRate:F4}%",
                    ["Days"] = days.ToString(),
                    ["DayCountBasis"] = denominator.ToString(),
                    ["RawInterest"] = $"{rawInterest:F6}",
                    ["RoundedInterest"] = $"{interestRounded:F2}"
                }
            }
        });
    }

    private void LogPrincipalCalculation(
        List<CalculationLogEntry> log,
        int paymentNumber,
        DateTime paymentDate,
        CreditParameters parameters,
        decimal principalPayment,
        decimal principalRemaining,
        decimal? fixedTotal,
        decimal interestRounded,
        bool isLast)
    {
        var principalRounded = RoundingService.Round(principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
        var roundingNeeded = Math.Abs(principalPayment - principalRounded) > 0.000001m;

        var description = parameters.PaymentType switch
        {
            PaymentType.Bullet => isLast ? "Spłata kapitału (bullet - ostatnia)" : "Spłata kapitału (bullet)",
            PaymentType.DecreasingInstallments => isLast ? "Spłata kapitału (finalna)" : "Spłata kapitału (rata malejąca)",
            _ => "Spłata kapitału (rata równa)"
        };

        var formula = parameters.PaymentType switch
        {
            PaymentType.Bullet when !isLast => "kapitał = 0",
            PaymentType.Bullet => "kapitał = saldo",
            PaymentType.DecreasingInstallments => isLast ? "kapitał = saldo" : "kapitał = saldo_początkowe / liczba_rat",
            _ => "kapitał = rata_docelowa - odsetki"
        };

        var substitution = parameters.PaymentType switch
        {
            PaymentType.Bullet when !isLast => "0",
            PaymentType.Bullet => $"{principalRemaining:F2}",
            PaymentType.DecreasingInstallments when isLast => $"{principalRemaining:F2}",
            PaymentType.DecreasingInstallments => $"{parameters.NetValue:F2} / liczba_rat",
            _ => $"{fixedTotal!.Value:F2} - {interestRounded:F2}"
        };

        var resultText = roundingNeeded
            ? $"{principalRounded:F2} PLN (przed zaokr.: {principalPayment:F6})"
            : $"{principalRounded:F2} PLN";

        log.Add(new CalculationLogEntry
        {
            ShortDescription = description,
            SymbolicFormula = formula,
            SubstitutedFormula = substitution,
            Result = resultText,
            Context = new LogEntryContext
            {
                PaymentNumber = paymentNumber,
                PaymentDate = paymentDate,
                Type = LogEntryType.PrincipalCalculation,
                Metadata = new Dictionary<string, string>
                {
                    ["PrincipalRemaining"] = $"{principalRemaining:F2}",
                    ["Interest"] = $"{interestRounded:F2}",
                    ["RawPrincipal"] = $"{principalPayment:F6}",
                    ["RoundedPrincipal"] = $"{principalRounded:F2}",
                    ["FixedTotal"] = fixedTotal.HasValue ? $"{fixedTotal.Value:F2}" : "N/A"
                }
            }
        });
    }

    private void LogRemainingBalance(
        List<CalculationLogEntry> log,
        int paymentNumber,
        DateTime paymentDate,
        decimal principalRemaining,
        decimal principalPayment,
        decimal interestPayment,
        decimal totalPayment)
    {
        var principalBefore = principalRemaining + principalPayment;

        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Podsumowanie raty",
            SymbolicFormula = "saldo_nowe = saldo_przed - kapitał",
            SubstitutedFormula = $"{principalBefore:F2} - {principalPayment:F2}",
            Result = $"Saldo: {principalRemaining:F2} PLN | Rata: {totalPayment:F2} PLN (odsetki: {interestPayment:F2}, kapitał: {principalPayment:F2})",
            Context = new LogEntryContext
            {
                PaymentNumber = paymentNumber,
                PaymentDate = paymentDate,
                Type = LogEntryType.Summary,
                Metadata = new Dictionary<string, string>
                {
                    ["PrincipalBefore"] = $"{principalBefore:F2}",
                    ["PrincipalAfter"] = $"{principalRemaining:F2}",
                    ["PrincipalPayment"] = $"{principalPayment:F2}",
                    ["InterestPayment"] = $"{interestPayment:F2}",
                    ["TotalPayment"] = $"{totalPayment:F2}"
                }
            }
        });
    }

    #endregion
}
