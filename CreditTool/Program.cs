using CreditTool.Models;
using CreditTool.Services;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddSingleton<IScheduleCalculator, DayToDayScheduleCalculator>();
builder.Services.AddSingleton<ExcelService>();
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
    return Results.Ok(schedule);
});

app.MapPost("/api/export", (CalculationRequest request, IScheduleCalculator calculator, ExcelService excelService) =>
{
    var schedule = calculator.Calculate(request.Parameters, request.Rates);
    var payload = excelService.Export(request.Parameters, request.Rates, schedule);
    var fileName = $"Schedule_{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";
    return Results.File(payload, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
});

app.Run();
