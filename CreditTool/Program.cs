using CreditTool.Models;
using CreditTool.Services;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton<IScheduleCalculator, DayToDayScheduleCalculator>();
builder.Services.AddSingleton<ExcelService>();
builder.Services.AddSingleton<WordExportService>();
builder.Services.AddEndpointsApiExplorer();
builder.Services.ConfigureHttpJsonOptions(options =>
{
    options.SerializerOptions.Converters.Add(new JsonStringEnumConverter());
});

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapPost("/api/import", async (IFormFile file, ExcelService excelService) =>
{
    if (file.Length == 0)
    {
        return Results.BadRequest("File is empty");
    }

    await using var stream = file.OpenReadStream();
    var result = excelService.Import(stream);
    return Results.Ok(new CalculationRequest { Parameters = result.Parameters, Rates = result.Rates });
});

app.MapPost("/api/calculate", (CalculationRequest request, IScheduleCalculator calculator) =>
{
    var schedule = calculator.Calculate(request.Parameters, request.Rates);
    return Results.Ok(BuildScheduleResponse(schedule));
});

app.MapPost("/api/export", (CalculationRequest request, string? format, IScheduleCalculator calculator, ExcelService excelService, WordExportService wordExportService) =>
{
    var schedule = calculator.Calculate(request.Parameters, request.Rates);
    var response = BuildScheduleResponse(schedule);

    if (string.Equals(format, "word", StringComparison.OrdinalIgnoreCase))
    {
        var wordPayload = wordExportService.Export(request.Parameters, request.Rates, schedule, response.TotalInterest);
        var wordFileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.docx";
        return Results.File(wordPayload, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", wordFileName);
    }

    var payload = excelService.Export(request.Parameters, request.Rates, schedule, response.TotalInterest);
    var fileName = $"Harmonogram_{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";
    return Results.File(payload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
});

app.Run();

static ScheduleResponse BuildScheduleResponse(IReadOnlyList<ScheduleItem> schedule)
{
    return new ScheduleResponse
    {
        Schedule = schedule.ToList(),
        TotalInterest = schedule.Sum(item => item.InterestAmount)
    };
}

public partial class Program;
