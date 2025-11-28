using CreditTool.Models;
using CreditTool.Services;
using CreditTool.Services.ScheduleCalculation;
using CreditTool.Services.ScheduleCalculation.Configuration;
using CreditTool.Services.ScheduleCalculation.Strategies.PaymentDate;
using Microsoft.AspNetCore.Antiforgery;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton<IPaymentDateGenerator, StandardPaymentDateGenerator>();
builder.Services.AddSingleton<IScheduleCalculator, ScheduleCalculator>();
builder.Services.AddSingleton(_ => new CalculatorConfiguration
{
    LevelPaymentTolerance = 0.0001m,
    EnableValidation = true,
    ThrowOnNegativeAmortization = false
});
builder.Services.AddSingleton<ExcelService>();
builder.Services.AddSingleton<LogExportService>();
builder.Services.AddEndpointsApiExplorer();
builder.Services.ConfigureHttpJsonOptions(options =>
{
    options.SerializerOptions.TypeInfoResolverChain.Insert(0, CreditJsonContext.Default);
    options.SerializerOptions.Converters.Add(new JsonStringEnumConverter());
});
builder.Services.AddAntiforgery();

var app = builder.Build();

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
app.UseDefaultFiles();
app.UseStaticFiles();
app.UseRouting();
app.UseAntiforgery();

app.MapPost("/api/import", async (IFormFile file, ExcelService excelService) =>
{
    if (file.Length == 0)
    {
        return Results.BadRequest("File is empty");
    }

    try
    {
        await using var stream = file.OpenReadStream();

        var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
        (CreditParameters Parameters, List<InterestRatePeriod> Rates) result = extension switch
        {
            ".xlsx" or ".ods" => excelService.Import(stream, extension),
            ".json" => await excelService.ImportJsonAsync(stream),
            _ => throw new InvalidOperationException("Unsupported import format. Use .xlsx, .ods or .json.")
        };

        return Results.Ok(new CalculationRequest { Parameters = result.Parameters, Rates = result.Rates });
    }
    catch (InvalidOperationException ex)
    {
        return Results.BadRequest(ex.Message);
    }
});

app.MapPost("/api/calculate", (
    CalculationRequest request,
    IScheduleCalculator calculator,
    CalculatorConfiguration config) =>
{
    var result = calculator.Calculate(request.Parameters, request.Rates, config);
    var roundedSchedule = RoundCashSchedule(result.Schedule, request.Parameters.RoundingMode);

    var response = BuildScheduleResponse(roundedSchedule, result.CalculationLog, request.Parameters);
    response.Warnings = result.Warnings;
    response.TargetLevelPayment = result.TargetLevelPayment;
    response.ActualFinalPayment = result.ActualFinalPayment;

    return Results.Ok(response);
});



app.MapPost("/api/export", (CalculationRequest request, string? format, IScheduleCalculator calculator, ExcelService excelService) =>
{
    var result = calculator.Calculate(request.Parameters, request.Rates);
    var roundedSchedule = RoundCashSchedule(result.Schedule, request.Parameters.RoundingMode);
    var response = BuildScheduleResponse(roundedSchedule, result.CalculationLog, request.Parameters);

    try
    {
        if (string.Equals(format, "json", StringComparison.OrdinalIgnoreCase))
        {
            var jsonPayload = excelService.ExportJson(request.Parameters, request.Rates, roundedSchedule, response.TotalInterest, response.AnnualPercentageRate);
            var jsonFileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.json";
            return Results.File(jsonPayload, "application/json", jsonFileName);
        }

        if (string.Equals(format, "ods", StringComparison.OrdinalIgnoreCase))
        {
            var odsPayload = excelService.ExportOds(request.Parameters, request.Rates, roundedSchedule, response.TotalInterest, response.AnnualPercentageRate);
            var odsFileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.ods";
            return Results.File(odsPayload, "application/vnd.oasis.opendocument.spreadsheet", odsFileName);
        }

        var payload = excelService.ExportXlsx(request.Parameters, request.Rates, roundedSchedule, response.TotalInterest, response.AnnualPercentageRate);
        var fileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";
        return Results.File(payload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    catch (InvalidOperationException ex)
    {
        return Results.BadRequest(ex.Message);
    }
});

app.MapPost("/api/export-log", (CalculationRequest request, IScheduleCalculator calculator, LogExportService logExportService) =>
{
    try
    {
        var result = calculator.Calculate(request.Parameters, request.Rates, includeLog: true);
        var logPayload = logExportService.Export(result.CalculationLog);
        var logFileName = $"Harmonogram_Log_{DateTime.UtcNow:yyyyMMddHHmmss}.md";
        return Results.File(logPayload, "text/markdown; charset=utf-8", logFileName);
    }
    catch (ArgumentException ex)
    {
        return Results.BadRequest(ex.Message);
    }
});

app.Run();

static ScheduleResponse BuildScheduleResponse(IReadOnlyList<ScheduleItem> schedule, List<CalculationLogEntry> calculationLog, CreditParameters parameters)
{
    var totalInterest = RoundingService.Round(schedule.Sum(item => item.InterestAmount), parameters.RoundingMode, 2);

    return new ScheduleResponse
    {
        Schedule = schedule.ToList(),
        CalculationLog = calculationLog,
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
