using CreditTool.Models;

namespace CreditTool.Services;

public interface IScheduleCalculator
{
    ScheduleCalculationResult Calculate(CreditParameters parameters, IReadOnlyCollection<InterestRatePeriod> ratePeriods);
}

public class DayToDayScheduleCalculator : IScheduleCalculator
{
    public ScheduleCalculationResult Calculate(CreditParameters parameters, IReadOnlyCollection<InterestRatePeriod> ratePeriods)
    {
        if (parameters.CreditEndDate <= parameters.CreditStartDate)
        {
            throw new ArgumentException("Credit end date must be after start date");
        }

        var paymentDates = BuildPaymentDates(parameters.CreditStartDate, parameters.CreditEndDate, parameters.PaymentFrequency, parameters.PaymentDay);
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
            fixedTotalPayment = CalculateLevelPayment(parameters, ratePeriods, paymentDates);
        }

        var schedule = new List<ScheduleItem>();
        var calculationLog = new List<CalculationLogEntry>();
        var previousDate = parameters.CreditStartDate;

        foreach (var paymentDate in paymentDates)
        {
            var daysInPeriod = (paymentDate - previousDate).Days;
            calculationLog.Add(new CalculationLogEntry
            {
                ShortDescription = "Wyznaczenie liczby dni okresu",
                SymbolicFormula = "dni = data_p\u0142atno\u015bci - poprzednia_data",
                SubstitutedFormula = $"dni = ({paymentDate:yyyy-MM-dd} - {previousDate:yyyy-MM-dd})",
                Result = $"{daysInPeriod} dni"
            });

            var interest = CalculateInterestDaily(previousDate, paymentDate, principalRemaining, parameters.MarginRate, ratePeriods, parameters.DayCountBasis);
            var interestRounded = RoundingService.Round(interest, parameters.RoundingMode, parameters.RoundingDecimals);

            var periodRate = FindRateForDate(ratePeriods, previousDate)?.Rate ?? 0m;
            var denominator = parameters.DayCountBasis == DayCountBasis.Actual360 ? 360m : 365m;
            calculationLog.Add(new CalculationLogEntry
            {
                ShortDescription = "Naliczanie odsetek",
                SymbolicFormula = "odsetki = saldo * (stawka + mar\u017ca) / 100 / baza_dni * dni",
                SubstitutedFormula = $"odsetki = {principalRemaining:F2} * ({periodRate:F4} + {parameters.MarginRate:F4}) / 100 / {denominator} * {daysInPeriod}",
                Result = $"{interestRounded:F4} PLN"
            });

            calculationLog.Add(new CalculationLogEntry
            {
                ShortDescription = "Zaokr\u0105glenie odsetek",
                SymbolicFormula = "odsetki_zaokr = Round(warto\u015b\u0107, tryb, miejsca)",
                SubstitutedFormula = $"Round({interest:F6}, tryb: {parameters.RoundingMode}, miejsca: {parameters.RoundingDecimals})",
                Result = $"{interestRounded:F4} PLN"
            });

            decimal principalPayment;
            if (parameters.PaymentType == PaymentType.Bullet && paymentDate != paymentDates[^1])
            {
                principalPayment = 0m;
            }
            else if (parameters.PaymentType == PaymentType.Bullet || parameters.PaymentType == PaymentType.DecreasingInstallments)
            {
                principalPayment = paymentDate == paymentDates[^1] ? principalRemaining : principalStep;
            }
            else
            {
                var targetTotal = fixedTotalPayment!.Value;
                principalPayment = targetTotal - interestRounded;
                if (principalPayment < 0m)
                {
                    principalPayment = 0m;
                }

                if (paymentDate != paymentDates[^1])
                {
                    principalPayment = Math.Min(principalPayment, principalRemaining);
                }
                else
                {
                    principalPayment = principalRemaining;
                }
            }

            var principalPaymentRaw = principalPayment;
            principalPayment = RoundingService.Round(principalPaymentRaw, parameters.RoundingMode, parameters.RoundingDecimals);
            calculationLog.Add(new CalculationLogEntry
            {
                ShortDescription = parameters.PaymentType switch
                {
                    PaymentType.Bullet => "Ustalenie sp\u0142aty kapita\u0142u (bullet)",
                    PaymentType.DecreasingInstallments => "Ustalenie sp\u0142aty kapita\u0142u (raty malej\u0105ce)",
                    _ => "Ustalenie sp\u0142aty kapita\u0142u (rata r\u00f3wna)"
                },
                SymbolicFormula = parameters.PaymentType switch
                {
                    PaymentType.Bullet when paymentDate != paymentDates[^1] => "kapita\u0142 = 0",
                    PaymentType.Bullet => "kapita\u0142 = saldo\u00a0ko\u0144cowe",
                    PaymentType.DecreasingInstallments => "kapita\u0142 = saldo_pocz\u0105tkowe / liczba_rat",
                    _ => "kapita\u0142 = rata_ca\u0142kowita - odsetki"
                },
                SubstitutedFormula = parameters.PaymentType switch
                {
                    PaymentType.Bullet when paymentDate != paymentDates[^1] => "kapita\u0142 = 0",
                    PaymentType.Bullet => $"kapita\u0142 = {principalRemaining:F2}",
                    PaymentType.DecreasingInstallments => $"kapita\u0142 = {principalRemaining:F2} / {paymentCount}",
                    _ => $"kapita\u0142 = {fixedTotalPayment!.Value:F4} - {interestRounded:F4}"
                },
                Result = $"{principalPayment:F4} PLN"
            });

            calculationLog.Add(new CalculationLogEntry
            {
                ShortDescription = "Zaokr\u0105glenie cz\u0119\u015bci kapita\u0142owej",
                SymbolicFormula = "kapita\u0142_zaokr = Round(warto\u015b\u0107, tryb, miejsca)",
                SubstitutedFormula = $"Round({principalPaymentRaw:F6}, tryb: {parameters.RoundingMode}, miejsca: {parameters.RoundingDecimals})",
                Result = $"{principalPayment:F4} PLN"
            });

            principalRemaining = RoundingService.Round(principalRemaining - principalPayment, parameters.RoundingMode, parameters.RoundingDecimals);
            calculationLog.Add(new CalculationLogEntry
            {
                ShortDescription = "Aktualizacja salda pozosta\u0142ego",
                SymbolicFormula = "saldo_nowe = saldo_poprz - kapita\u0142_zaokr",
                SubstitutedFormula = $"saldo_nowe = {principalRemaining + principalPayment:F4} - {principalPayment:F4}",
                Result = $"{principalRemaining:F4} PLN"
            });
            var totalPayment = fixedTotalPayment ?? (interestRounded + principalPayment);

            schedule.Add(new ScheduleItem
            {
                PaymentDate = paymentDate,
                DaysInPeriod = daysInPeriod,
                InterestRate = periodRate + parameters.MarginRate,
                InterestAmount = interestRounded,
                PrincipalPayment = principalPayment,
                TotalPayment = totalPayment,
                RemainingPrincipal = Math.Max(principalRemaining, 0m)
            });

            previousDate = paymentDate;
        }

        return new ScheduleCalculationResult
        {
            Schedule = schedule,
            CalculationLog = calculationLog
        };
    }

    private decimal CalculateLevelPayment(
        CreditParameters parameters,
        IReadOnlyCollection<InterestRatePeriod> ratePeriods,
        IReadOnlyList<DateTime> paymentDates)
    {
        var low = 0m;
        var high = Math.Max(parameters.NetValue, parameters.NetValue + 1000m);
        var tolerance = 0.01m;

        var remainingHigh = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, high);
        var guard = 0;
        while (remainingHigh > 0m && guard < 25)
        {
            low = high;
            high *= 2m;
            remainingHigh = SimulateRemainingPrincipal(parameters, ratePeriods, paymentDates, high);
            guard++;
        }

        for (var i = 0; i < 200; i++)
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
            var interest = CalculateInterestDaily(previousDate, paymentDate, principalRemaining, parameters.MarginRate, ratePeriods, parameters.DayCountBasis);
            var interestRounded = RoundingService.Round(interest, parameters.RoundingMode, parameters.RoundingDecimals);

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

    private static decimal CalculateInterestDaily(DateTime from, DateTime to, decimal principal, decimal marginRate, IReadOnlyCollection<InterestRatePeriod> ratePeriods, DayCountBasis basis)
    {
        decimal interest = 0m;
        var denominator = basis == DayCountBasis.Actual360 ? 360m : 365m;

        for (var day = from; day < to; day = day.AddDays(1))
        {
            var rate = FindRateForDate(ratePeriods, day)?.Rate ?? 0m;
            var effectiveRate = rate + marginRate;
            interest += principal * effectiveRate / 100m / denominator;
        }

        return interest;
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
