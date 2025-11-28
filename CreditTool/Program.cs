using CreditTool.Models;
using CreditTool.Services;
using Microsoft.AspNetCore.Antiforgery;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton<IScheduleCalculator, DayToDayScheduleCalculator>();
builder.Services.AddSingleton<ExcelService>();
builder.Services.AddSingleton<WordExportService>();
builder.Services.AddSingleton<WordImportService>();
builder.Services.AddEndpointsApiExplorer();
builder.Services.ConfigureHttpJsonOptions(options =>
{
    options.SerializerOptions.Converters.Add(new JsonStringEnumConverter());
});
builder.Services.AddAntiforgery();

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();
app.UseRouting();
app.Use(async (context, next) =>
{
    if (HttpMethods.IsGet(context.Request.Method) && (context.Request.Path == "/" || context.Request.Path == "/index.html"))
    {
        var antiforgery = context.RequestServices.GetRequiredService<IAntiforgery>();
        var tokens = antiforgery.GetAndStoreTokens(context);
        context.Response.Cookies.Append("XSRF-TOKEN", tokens.RequestToken!, new CookieOptions
        {
            HttpOnly = false,
            Secure = context.Request.IsHttps,
            SameSite = SameSiteMode.Strict
        });
    }

    await next();
});
app.UseAntiforgery();

app.MapPost("/api/import", async (IFormFile file, ExcelService excelService, WordImportService wordImportService) =>
{
    if (file.Length == 0)
    {
        return Results.BadRequest("File is empty");
    }

    await using var stream = file.OpenReadStream();

    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
    (CreditParameters Parameters, List<InterestRatePeriod> Rates) result = extension switch
    {
        ".docx" => wordImportService.Import(stream),
        _ => excelService.Import(stream)
    };

    return Results.Ok(new CalculationRequest { Parameters = result.Parameters, Rates = result.Rates });
});

app.MapPost("/api/import/word", async (IFormFile file, WordImportService wordImportService) =>
{
    if (file.Length == 0)
    {
        return Results.BadRequest("File is empty");
    }

    await using var stream = file.OpenReadStream();
    var result = wordImportService.Import(stream);
    return Results.Ok(new CalculationRequest { Parameters = result.Parameters, Rates = result.Rates });
});

app.MapPost("/api/calculate", (CalculationRequest request, IScheduleCalculator calculator) =>
{
    var schedule = calculator.Calculate(request.Parameters, request.Rates);
    var roundedSchedule = RoundCashSchedule(schedule, request.Parameters.RoundingMode);
    return Results.Ok(BuildScheduleResponse(roundedSchedule, request.Parameters));
});

app.MapPost("/api/export", (CalculationRequest request, string? format, IScheduleCalculator calculator, ExcelService excelService, WordExportService wordExportService) =>
{
    var schedule = calculator.Calculate(request.Parameters, request.Rates);
    var roundedSchedule = RoundCashSchedule(schedule, request.Parameters.RoundingMode);
    var response = BuildScheduleResponse(roundedSchedule, request.Parameters);

    if (string.Equals(format, "word", StringComparison.OrdinalIgnoreCase))
    {
        var wordPayload = wordExportService.Export(request.Parameters, request.Rates, roundedSchedule, response.TotalInterest);
        var wordFileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.docx";
        return Results.File(wordPayload, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", wordFileName);
    }

    var payload = excelService.Export(request.Parameters, request.Rates, roundedSchedule, response.TotalInterest);
    var fileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";
    return Results.File(payload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
});

app.Run();

static ScheduleResponse BuildScheduleResponse(IReadOnlyList<ScheduleItem> schedule, CreditParameters parameters)
{
    var totalInterest = RoundingService.Round(schedule.Sum(item => item.InterestAmount), parameters.RoundingMode, 2);

    return new ScheduleResponse
    {
        Schedule = schedule.ToList(),
        TotalInterest = totalInterest,
        AnnualPercentageRate = AprCalculator.CalculateAnnualPercentageRate(parameters, schedule)
    };
}

static IReadOnlyList<ScheduleItem> RoundCashSchedule(IEnumerable<ScheduleItem> schedule, RoundingModeOption roundingMode)
{
    return schedule.Select(item =>
    {
        var interest = RoundingService.Round(item.InterestAmount, roundingMode, 2);
        var principal = RoundingService.Round(item.PrincipalPayment, roundingMode, 2);
        var total = RoundingService.Round(interest + principal, roundingMode, 2);
        var remaining = RoundingService.Round(Math.Max(item.RemainingPrincipal, 0m), roundingMode, 2);

        return new ScheduleItem
        {
            PaymentDate = item.PaymentDate,
            DaysInPeriod = item.DaysInPeriod,
            InterestRate = item.InterestRate,
            InterestAmount = interest,
            PrincipalPayment = principal,
            TotalPayment = total,
            RemainingPrincipal = remaining
        };
    }).ToList();
}

public partial class Program;
