using System.Linq;
using System.Net.Http.Json;
using CreditTool.Models;
using CreditTool.Services;
using Microsoft.AspNetCore.Mvc.Testing;

namespace CreditTool.Tests;

public class IntegrationTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient _client;

    public IntegrationTests(WebApplicationFactory<Program> factory)
    {
        _client = factory.CreateClient();
    }

    [Fact]
    public async Task CalculateEndpoint_ReturnsTotalInterestForYearConstantRate()
    {
        var request = new CalculationRequest
        {
            Parameters = new CreditParameters
            {
                NetValue = 10000m,
                MarginRate = 0m,
                PaymentFrequency = PaymentFrequency.Monthly,
                PaymentDay = PaymentDayOption.LastOfMonth,
                CreditStartDate = new DateTime(2024, 1, 1),
                CreditEndDate = new DateTime(2025, 1, 1),
                DayCountBasis = DayCountBasis.Actual365,
                RoundingMode = RoundingModeOption.Bankers,
                RoundingDecimals = 4,
                BulletRepayment = true
            },
            Rates = new List<InterestRatePeriod>
            {
                new()
                {
                    DateFrom = new DateTime(2024, 1, 1),
                    DateTo = new DateTime(2024, 12, 31),
                    Rate = 5m
                }
            }
        };

        var response = await _client.PostAsJsonAsync("/api/calculate", request);
        response.EnsureSuccessStatusCode();

        var payload = await response.Content.ReadFromJsonAsync<ScheduleResponse>();
        Assert.NotNull(payload);
        Assert.NotEmpty(payload!.Schedule);

        var calculator = new DayToDayScheduleCalculator();
        var expectedSchedule = calculator.Calculate(request.Parameters, request.Rates);
        var expectedTotalInterest = expectedSchedule
            .Select(item => RoundingService.Round(item.InterestAmount, request.Parameters.RoundingMode, 2))
            .Sum();

        Assert.Equal(expectedTotalInterest, payload.TotalInterest);
    }

    [Fact]
    public async Task RootRequest_IssuesAntiforgeryCookie()
    {
        var response = await _client.GetAsync("/");
        response.EnsureSuccessStatusCode();

        Assert.True(response.Headers.TryGetValues("Set-Cookie", out var cookies));
        Assert.Contains(cookies, cookie => cookie.Contains(".AspNetCore.Antiforgery"));
    }
}
