// DayToDayScheduleCalculator.cs (complete replacement)
using CreditTool.Models;

namespace CreditTool.Services;

public interface IScheduleCalculator
{
    ScheduleCalculationResult Calculate(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        CalculatorConfiguration? configuration = null,
        bool includeLog = false);
}

public class DayToDayScheduleCalculator : IScheduleCalculator
{
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

        var paymentDates = BuildPaymentDates(
            parameters.CreditStartDate,
            parameters.CreditEndDate,
            parameters.PaymentFrequency,
            parameters.PaymentDay);

        if (paymentDates.Count == 0 || paymentDates[^1] != parameters.CreditEndDate)
        {
            paymentDates.Add(parameters.CreditEndDate);
        }

        var principalRemaining = parameters.NetValue;
        var paymentCount = paymentDates.Count;
        var principalStep = parameters.PaymentType == PaymentType.Bullet
            ? 0m
            : RoundingService.Round(parameters.NetValue / paymentCount, parameters.RoundingMode, parameters.RoundingDecimals);

        decimal? fixedTotalPayment = null;
        if (parameters.PaymentType == PaymentType.EqualInstallments)
        {
            fixedTotalPayment = CalculateLevelPayment(parameters, ratePeriods, paymentDates, configuration);
        }

        var schedule = new List<ScheduleItem>();
        var calculationLog = new List<CalculationLogEntry>();
        var warnings = new List<string>();
        var previousDate = parameters.CreditStartDate;

        var recalculatedRatePeriods = ratePeriods
            .Where(period => period.DateFrom > parameters.CreditStartDate && period.DateFrom < parameters.CreditEndDate)
            .OrderBy(period => period.DateFrom);

        if (includeLog)
        {
            foreach (var period in recalculatedRatePeriods)
            {
                calculationLog.Add(new CalculationLogEntry
                {
                    ShortDescription = "Aktualizacja stopy procentowej w trakcie kredytu",
                    SymbolicFormula = "stawka_okresowa = stawka_bazowa + marża",
                    SubstitutedFormula = $"stawka_okresowa = {period.Rate:F4}% + {parameters.MarginRate:F4}% (od {period.DateFrom:yyyy-MM-dd})",
                    Result = $"Nowa stopa od {period.DateFrom:yyyy-MM-dd} do {period.DateTo:yyyy-MM-dd}: {(period.Rate + parameters.MarginRate):F4}%"
                });
            }
        }

        for (var index = 0; index < paymentDates.Count; index++)
        {
            var paymentDate = paymentDates[index];
            var isLastPayment = index == paymentDates.Count - 1;

            var daysInPeriod = (paymentDate - previousDate).Days;
            if (includeLog)
            {
                calculationLog.Add(new CalculationLogEntry
                {
                    ShortDescription = "Wyznaczenie liczby dni okresu",
                    SymbolicFormula = "dni = data_płatności - poprzednia_data",
                    SubstitutedFormula = $"dni = ({paymentDate:yyyy-MM-dd} - {previousDate:yyyy-MM-dd})",
                    Result = $"{daysInPeriod} dni"
                });
            }

            var interestResult = CalculateInterest(
                previousDate,
                paymentDate,
                principalRemaining,
                parameters.MarginRate,
                ratePeriods,
                parameters.DayCountBasis,
                parameters.InterestRateApplication);

            var interestRounded = RoundingService.Round(interestResult.Interest, parameters.RoundingMode, parameters.RoundingDecimals);

            var baseRate = Math.Max(interestResult.EffectiveRate - parameters.MarginRate, 0m);
            var denominator = parameters.DayCountBasis == DayCountBasis.Actual360 ? 360m : 365m;

            if (includeLog)
            {
                LogInterestCalculation(calculationLog, parameters, interestResult, interestRounded,
                principalRemaining, baseRate, denominator, daysInPeriod);
            }
            
            decimal principalPayment;
            if (parameters.PaymentType == PaymentType.Bullet && !isLastPayment)
            {
                principalPayment = 0m;
            }
            else if (parameters.PaymentType == PaymentType.Bullet || parameters.PaymentType == PaymentType.DecreasingInstallments)
            {
                principalPayment = isLastPayment ? principalRemaining : principalStep;
            }
            else
            {
                var targetTotal = fixedTotalPayment!.Value;
                principalPayment = targetTotal - interestRounded;

                if (principalPayment < 0m)
                {
                    principalPayment = 0m;
                    if (configuration.EnableValidation)
                    {
                        var warning = $"Rata {index + 1}: Odsetki ({interestRounded:F2}) przekraczają ratę docelową ({targetTotal:F2})";
                        warnings.Add(warning);

                        if (configuration.ThrowOnNegativeAmortization)
                        {
                            throw new InvalidOperationException(warning);
                        }
                    }
                }

                if (!isLastPayment)
                {
                    principalPayment = Math.Min(principalPayment, principalRemaining);
                }
                else
                {
                    // Final payment: force exact balance closure
                    principalPayment = principalRemaining;
                }
            }

            var principalPaymentRaw = principalPayment;
            principalPayment = RoundingService.Round(principalPaymentRaw, parameters.RoundingMode, parameters.RoundingDecimals);

            if (includeLog)
            {
                LogPrincipalCalculation(calculationLog, parameters, paymentDate, paymentDates,
                principalPayment, principalPaymentRaw, principalRemaining,
                paymentCount, fixedTotalPayment, interestRounded);
            }           

            principalRemaining = RoundingService.Round(principalRemaining - principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            if (includeLog)
            {
                calculationLog.Add(new CalculationLogEntry
                {
                    ShortDescription = "Aktualizacja salda pozostałego",
                    SymbolicFormula = "saldo_nowe = saldo_poprz - kapitał_zaokr",
                    SubstitutedFormula = $"saldo_nowe = {principalRemaining + principalPayment:F4} - {principalPayment:F4}",
                    Result = $"{principalRemaining:F4} PLN"
                });
            }            

            var totalPayment = interestRounded + principalPayment;
            var isFinalAdjusted = false;
            var itemWarnings = ScheduleWarnings.None;

            if (parameters.PaymentType == PaymentType.EqualInstallments)
            {
                if (isLastPayment && Math.Abs(totalPayment - fixedTotalPayment!.Value) > 0.01m)
                {
                    isFinalAdjusted = true;
                    itemWarnings |= ScheduleWarnings.FinalPaymentAdjusted;
                }

                if (interestRounded > fixedTotalPayment!.Value)
                {
                    itemWarnings |= ScheduleWarnings.InterestExceedsPayment;
                }

                if (principalPaymentRaw < 0m)
                {
                    itemWarnings |= ScheduleWarnings.NegativeAmortization;
                }
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

    private void LogInterestCalculation(
        List<CalculationLogEntry> log,
        CreditParameters parameters,
        InterestCalculationResult result,
        decimal interestRounded,
        decimal principal,
        decimal baseRate,
        decimal denominator,
        int days)
    {
        var interestDescription = parameters.InterestRateApplication switch
        {
            InterestRateApplication.ApplyChangedRateNextPeriod => "Naliczanie odsetek (zmiana stopy od kolejnego okresu)",
            InterestRateApplication.CompoundDaily => "Naliczanie odsetek (kapitalizacja dzienna)",
            InterestRateApplication.CompoundMonthly => "Naliczanie odsetek (kapitalizacja miesięczna)",
            InterestRateApplication.CompoundQuarterly => "Naliczanie odsetek (kapitalizacja kwartalna)",
            _ => "Naliczanie odsetek"
        };

        var interestFormula = parameters.InterestRateApplication switch
        {
            InterestRateApplication.ApplyChangedRateNextPeriod => "odsetki = saldo * (stopa_na_początek + marża) / 100 / baza_dni * dni",
            InterestRateApplication.CompoundDaily => "odsetki = saldo * (∏(1 + r_dzienne/100/baza) - 1)",
            InterestRateApplication.CompoundMonthly => "odsetki = saldo * ((1 + r_nom/100/12)^(12*t) - 1)",
            InterestRateApplication.CompoundQuarterly => "odsetki = saldo * ((1 + r_nom/100/4)^(4*t) - 1)",
            _ => "odsetki = Σ(saldo * (stawka_dzienna + marża) / 100 / baza_dni)"
        };

        string interestSubstitution;
        if (parameters.InterestRateApplication == InterestRateApplication.ApplyChangedRateNextPeriod)
        {
            interestSubstitution = $"odsetki = {principal:F2} * ({baseRate:F4} + {parameters.MarginRate:F4}) / 100 / {denominator} * {days}";
        }
        else if (result.NominalRate.HasValue && result.EffectivePeriodRate.HasValue)
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
        DateTime paymentDate,
        List<DateTime> paymentDates,
        decimal principalPayment,
        decimal principalPaymentRaw,
        decimal principalRemaining,
        int paymentCount,
        decimal? fixedTotal,
        decimal interestRounded)
    {
        var isLast = paymentDate == paymentDates[^1];

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
                PaymentType.DecreasingInstallments => $"kapitał = {parameters.NetValue:F4} / {paymentCount}",
                _ => $"kapitał = {fixedTotal!.Value:F4} - {interestRounded:F4}"
            },
            Result = $"{principalPayment:F4} PLN"
        });

        log.Add(new CalculationLogEntry
        {
            ShortDescription = "Zaokrąglenie części kapitałowej",
            SymbolicFormula = "kapitał_zaokr = Round(wartość, tryb, miejsca)",
            SubstitutedFormula = $"Round({principalPaymentRaw:F6}, tryb: {parameters.RoundingMode}, miejsca: {parameters.RoundingDecimals})",
            Result = $"{principalPayment:F4} PLN"
        });
    }

    private decimal CalculateLevelPayment(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates,
        CalculatorConfiguration configuration)
    {
        var low = 0m;
        var high = Math.Max(parameters.NetValue, parameters.NetValue + 1000m);
        var tolerance = configuration.LevelPaymentTolerance;

        var remainingHigh = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, high);
        var guard = 0;
        while (remainingHigh > 0m && guard < configuration.MaxRangeFindingIterations)
        {
            low = high;
            high *= 2m;
            remainingHigh = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, high);
            guard++;
        }

        for (var i = 0; i < configuration.MaxLevelPaymentIterations; i++)
        {
            var mid = (low + high) / 2m;
            var remaining = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, mid);

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

        return RoundingService.Round((low + high) / 2m, parameters.RoundingMode, parameters.RoundingDecimals);
    }

    private decimal SimulateRemainingPrincipal(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates,
        decimal totalPayment)
    {
        var principalRemaining = parameters.NetValue;
        var previousDate = parameters.CreditStartDate;

        foreach (var paymentDate in paymentDates)
        {
            var result = CalculateInterest(
                previousDate,
                paymentDate,
                principalRemaining,
                parameters.MarginRate,
                ratePeriods,
                parameters.DayCountBasis,
                parameters.InterestRateApplication);

            var interestRounded = RoundingService.Round(result.Interest, parameters.RoundingMode, parameters.RoundingDecimals);

            var principalPayment = totalPayment - interestRounded;
            if (principalPayment < 0m)
            {
                principalPayment = 0m;
            }

            principalPayment = Math.Min(principalPayment, principalRemaining);
            principalPayment = RoundingService.Round(principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            principalRemaining = RoundingService.Round(principalRemaining - principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            previousDate = paymentDate;
        }

        return principalRemaining;
    }

    private static InterestCalculationResult CalculateInterest(
        DateTime from,
        DateTime to,
        decimal principal,
        decimal marginRate,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        DayCountBasis basis,
        InterestRateApplication application)
    {
        var daysInPeriod = Math.Max((to - from).Days, 0);
        var denominator = basis == DayCountBasis.Actual360 ? 360m : 365m;

        switch (application)
        {
            case InterestRateApplication.ApplyChangedRateNextPeriod:
                {
                    var baseRate = FindRateForDate(ratePeriods, from)?.Rate ?? 0m;
                    var effectiveRate = baseRate + marginRate;
                    var interest = principal * effectiveRate / 100m / denominator * daysInPeriod;
                    return new InterestCalculationResult(interest, effectiveRate, null, null);
                }
            case InterestRateApplication.CompoundDaily:
                {
                    decimal factor = 1m;
                    decimal totalRate = 0m;

                    for (var day = from; day < to; day = day.AddDays(1))
                    {
                        var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
                        totalRate += rate;
                        var effectiveRateDaily = rate + marginRate;
                        factor *= 1m + (effectiveRateDaily / 100m) / denominator;
                    }

                    var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
                    var nominalRate = averageRate + marginRate;
                    var periodRate = factor - 1m;
                    var interest = principal * periodRate;

                    return new InterestCalculationResult(interest, nominalRate, nominalRate, periodRate);
                }
            case InterestRateApplication.CompoundMonthly:
                {
                    decimal totalRate = 0m;
                    for (var day = from; day < to; day = day.AddDays(1))
                    {
                        var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
                        totalRate += rate;
                    }

                    var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
                    var nominalRate = averageRate + marginRate;
                    var n = 12m;
                    var tYears = daysInPeriod / denominator;
                    var rNom = nominalRate / 100m;

                    var periodFactor = (decimal)Math.Pow(1.0 + (double)(rNom / n), (double)(n * tYears));
                    var periodRate = periodFactor - 1m;
                    var interest = principal * periodRate;

                    return new InterestCalculationResult(interest, nominalRate, nominalRate, periodRate);
                }
            case InterestRateApplication.CompoundQuarterly:
                {
                    decimal totalRate = 0m;
                    for (var day = from; day < to; day = day.AddDays(1))
                    {
                        var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
                        totalRate += rate;
                    }

                    var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
                    var nominalRate = averageRate + marginRate;
                    var n = 4m;
                    var tYears = daysInPeriod / denominator;
                    var rNom = nominalRate / 100m;

                    var periodFactor = (decimal)Math.Pow(1.0 + (double)(rNom / n), (double)(n * tYears));
                    var periodRate = periodFactor - 1m;
                    var interest = principal * periodRate;

                    return new InterestCalculationResult(interest, nominalRate, nominalRate, periodRate);
                }
            default:
                {
                    decimal interest = 0m;
                    decimal totalRate = 0m;

                    for (var day = from; day < to; day = day.AddDays(1))
                    {
                        var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
                        totalRate += rate;
                        var effectiveRate = rate + marginRate;
                        interest += principal * effectiveRate / 100m / denominator;
                    }

                    var averageRate = daysInPeriod > 0 ? totalRate / daysInPeriod : 0m;
                    return new InterestCalculationResult(interest, averageRate + marginRate, null, null);
                }
        }
    }

    private static InterestRatePeriod? FindRateForDate(IEnumerable<InterestRatePeriod> periods, DateTime date)
    {
        return periods.FirstOrDefault(period => period.DateFrom.Date <= date.Date && period.DateTo.Date >= date.Date);
    }

    private static List<DateTime> BuildPaymentDates(DateTime start, DateTime end, PaymentFrequency frequency, PaymentDayOption paymentDay)
    {
        var dates = new List<DateTime>();
        var current = start;

        while (current < end)
        {
            DateTime next = frequency switch
            {
                PaymentFrequency.Daily => current.AddDays(1),
                PaymentFrequency.Monthly => NextMonthDate(current, 1, paymentDay),
                PaymentFrequency.Quarterly => NextMonthDate(current, 3, paymentDay),
                _ => current.AddMonths(1)
            };

            if (next > end)
            {
                next = end;
            }

            dates.Add(next);
            current = next;
        }

        return dates;
    }

    private static DateTime NextMonthDate(DateTime from, int monthsToAdd, PaymentDayOption paymentDay)
    {
        var tentative = from.AddMonths(monthsToAdd);
        var year = tentative.Year;
        var month = tentative.Month;

        return paymentDay switch
        {
            PaymentDayOption.FirstOfMonth => new DateTime(year, month, 1),
            PaymentDayOption.TenthOfMonth => new DateTime(year, month, 10),
            PaymentDayOption.LastOfMonth => new DateTime(year, month, DateTime.DaysInMonth(year, month)),
            _ => tentative
        };
    }
}

internal record InterestCalculationResult(
    decimal Interest,
    decimal EffectiveRate,
    decimal? NominalRate,
    decimal? EffectivePeriodRate);