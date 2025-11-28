using CreditTool.Models;

namespace CreditTool.Services;

public static class AprCalculator
{
    public static decimal CalculateAnnualPercentageRate(CreditParameters parameters, IEnumerable<ScheduleItem> schedule)
    {
        var cashFlows = BuildCashFlows(parameters, schedule);
        if (cashFlows.Count < 2)
        {
            return 0m;
        }

        var startDate = cashFlows[0].Date;
        double Npv(double rate)
        {
            return cashFlows.Sum(cf => (double)cf.Amount * Math.Pow(1 + rate, -(cf.Date - startDate).TotalDays / 365d));
        }

        double Derivative(double rate)
        {
            return cashFlows.Sum(cf =>
            {
                var years = (cf.Date - startDate).TotalDays / 365d;
                return -years * (double)cf.Amount * Math.Pow(1 + rate, -(years + 1));
            });
        }

        var guess = 0.1d;
        for (var iteration = 0; iteration < 50; iteration++)
        {
            var value = Npv(guess);
            var slope = Derivative(guess);
            if (Math.Abs(slope) < 1e-12)
            {
                break;
            }

            var nextGuess = guess - value / slope;
            if (nextGuess <= -0.99d || nextGuess > 10d)
            {
                break;
            }

            if (Math.Abs(nextGuess - guess) < 1e-8)
            {
                return (decimal)Math.Round(nextGuess * 100, 4, MidpointRounding.AwayFromZero);
            }

            guess = nextGuess;
        }

        var lower = -0.99d;
        var upper = 1.0d;
        for (var i = 0; i < 200; i++)
        {
            var mid = (lower + upper) / 2d;
            var value = Npv(mid);
            if (Math.Abs(value) < 1e-8)
            {
                return (decimal)Math.Round(mid * 100, 4, MidpointRounding.AwayFromZero);
            }

            if (value > 0)
            {
                lower = mid;
            }
            else
            {
                upper = mid;
            }
        }

        var rateEstimate = (lower + upper) / 2d;
        return (decimal)Math.Round(rateEstimate * 100, 4, MidpointRounding.AwayFromZero);
    }

    private static List<(DateTime Date, decimal Amount)> BuildCashFlows(CreditParameters parameters, IEnumerable<ScheduleItem> schedule)
    {
        var disbursement = parameters.NetValue;
        if (parameters.ProcessingFeeRate > 0)
        {
            disbursement -= parameters.NetValue * parameters.ProcessingFeeRate / 100m;
        }

        if (parameters.ProcessingFeeAmount > 0)
        {
            disbursement -= parameters.ProcessingFeeAmount;
        }

        var flows = new List<(DateTime Date, decimal Amount)> { (parameters.CreditStartDate, disbursement) };
        flows.AddRange(schedule.Select(item => (item.PaymentDate, -item.TotalPayment)));
        flows.Sort((a, b) => a.Date.CompareTo(b.Date));
        return flows;
    }
}
