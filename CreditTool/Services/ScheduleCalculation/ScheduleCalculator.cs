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
                LogDaysInPeriod(calculationLog, paymentDate, previousDate, daysInPeriod);
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
                LogRemainingBalance(calculationLog, principalRemaining, principalPayment);
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
                ShortDescription = "Aktualizacja stopy procentowej w trakcie kredytu",
                SymbolicFormula = "stawka_okresowa = stawka_bazowa + marża",
                SubstitutedFormula = $"stawka_okresowa = {period.Rate:F4}% + {parameters.MarginRate:F4}% (od {period.DateFrom:yyyy-MM-dd})",
                Result = $"Nowa stopa od {period.DateFrom:yyyy-MM-dd} do {period.DateTo:yyyy-MM-dd}: {(period.Rate + parameters.MarginRate):F4}%"
            });
        }
    }

    private void LogDaysInPeriod(
        List<CalculationLogEntry> log,
        DateTime paymentDate,
        DateTime previousDate,
        int daysInPeriod)
    {
        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Wyznaczenie liczby dni okresu",
            SymbolicFormula = "dni = data_płatności - poprzednia_data",
            SubstitutedFormula = $"dni = ({paymentDate:yyyy-MM-dd} - {previousDate:yyyy-MM-dd})",
            Result = $"{daysInPeriod} dni"
        });
    }

    private void LogInterestCalculation(
        List<CalculationLogEntry> log,
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
                "Naliczanie odsetek (zmiana stopy od kolejnego okresu)",
            InterestRateApplication.CompoundDaily =>
                "Naliczanie odsetek (kapitalizacja dzienna)",
            InterestRateApplication.CompoundMonthly =>
                "Naliczanie odsetek (kapitalizacja miesięczna)",
            InterestRateApplication.CompoundQuarterly =>
                "Naliczanie odsetek (kapitalizacja kwartalna)",
            _ => "Naliczanie odsetek"
        };

        var interestFormula = parameters.InterestRateApplication switch
        {
            InterestRateApplication.ApplyChangedRateNextPeriod =>
                "odsetki = saldo * (stopa_na_początek + marża) / 100 / baza_dni * dni",
            InterestRateApplication.CompoundDaily =>
                "odsetki = saldo * (∏(1 + r_dzienne/100/baza) - 1)",
            InterestRateApplication.CompoundMonthly =>
                "odsetki = saldo * ((1 + r_nom/100/12)^(12*t) - 1)",
            InterestRateApplication.CompoundQuarterly =>
                "odsetki = saldo * ((1 + r_nom/100/4)^(4*t) - 1)",
            _ => "odsetki = Σ(saldo * (stawka_dzienna + marża) / 100 / baza_dni)"
        };

        string interestSubstitution;
        if (result.NominalRate.HasValue && result.EffectivePeriodRate.HasValue)
        {
            interestSubstitution = $"odsetki = {principal:F2} * {result.EffectivePeriodRate.Value:F6} (stopa nom.: {result.NominalRate.Value:F4}%, okres: {days} dni)";
        }
        else
        {
            interestSubstitution = $"odsetki = {principal:F2} * średnia_stawka_dzienna (baza: {denominator}, dni: {days})";
        }

        log.Add(new CalculationLogEntry
        {
            ShortDescription = interestDescription,
            SymbolicFormula = interestFormula,
            SubstitutedFormula = interestSubstitution,
            Result = $"{interestRounded:F4} PLN"
        });

        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Zaokrąglenie odsetek",
            SymbolicFormula = "odsetki_zaokr = Round(wartość, tryb, miejsca)",
            SubstitutedFormula = $"Round({result.Interest:F6}, tryb: {parameters.RoundingMode}, miejsca: {parameters.RoundingDecimals})",
            Result = $"{interestRounded:F4} PLN"
        });
    }

    private void LogPrincipalCalculation(
        List<CalculationLogEntry> log,
        CreditParameters parameters,
        decimal principalPayment,
        decimal principalRemaining,
        decimal? fixedTotal,
        decimal interestRounded,
        bool isLast)
    {
        log.Add(new CalculationLogEntry
        {
            ShortDescription = parameters.PaymentType switch
            {
                PaymentType.Bullet => "Ustalenie spłaty kapitału (bullet)",
                PaymentType.DecreasingInstallments => "Ustalenie spłaty kapitału (raty malejące)",
                _ => "Ustalenie spłaty kapitału (rata równa)"
            },
            SymbolicFormula = parameters.PaymentType switch
            {
                PaymentType.Bullet when !isLast => "kapitał = 0",
                PaymentType.Bullet => "kapitał = saldo końcowe",
                PaymentType.DecreasingInstallments => isLast
                    ? "kapitał = saldo końcowe"
                    : "kapitał = saldo_początkowe / liczba_rat",
                _ => "kapitał = rata_całkowita - odsetki"
            },
            SubstitutedFormula = parameters.PaymentType switch
            {
                PaymentType.Bullet when !isLast => "kapitał = 0",
                PaymentType.Bullet => $"kapitał = {principalRemaining:F2}",
                PaymentType.DecreasingInstallments when isLast => $"kapitał = {principalRemaining:F4}",
                PaymentType.DecreasingInstallments => $"kapitał = {parameters.NetValue:F4} / liczba_rat",
                _ => $"kapitał = {fixedTotal!.Value:F4} - {interestRounded:F4}"
            },
            Result = $"{principalPayment:F4} PLN"
        });

        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Zaokrąglenie części kapitałowej",
            SymbolicFormula = "kapitał_zaokr = Round(wartość, tryb, miejsca)",
            SubstitutedFormula = $"Round({principalPayment:F6}, tryb: {parameters.RoundingMode}, miejsca: {parameters.RoundingDecimals})",
            Result = $"{principalPayment:F4} PLN"
        });
    }

    private void LogRemainingBalance(
        List<CalculationLogEntry> log,
        decimal principalRemaining,
        decimal principalPayment)
    {
        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Aktualizacja salda pozostałego",
            SymbolicFormula = "saldo_nowe = saldo_poprz - kapitał_zaokr",
            SubstitutedFormula = $"saldo_nowe = {principalRemaining + principalPayment:F4} - {principalPayment:F4}",
            Result = $"{principalRemaining:F4} PLN"
        });
    }

    #endregion
}
