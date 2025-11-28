using System.Text.Json.Serialization;

namespace CreditTool.Models;

[JsonSerializable(typeof(CalculationRequest))]
[JsonSerializable(typeof(CreditParameters))]
[JsonSerializable(typeof(InterestRatePeriod))]
[JsonSerializable(typeof(ScheduleResponse))]
[JsonSerializable(typeof(ScheduleItem))]
[JsonSerializable(typeof(CalculationLogEntry))]
[JsonSerializable(typeof(LogEntryContext))]
[JsonSerializable(typeof(ScheduleCalculationResult))]
[JsonSerializable(typeof(PaymentFrequency))]
[JsonSerializable(typeof(PaymentDayOption))]
[JsonSerializable(typeof(DayCountBasis))]
[JsonSerializable(typeof(RoundingModeOption))]
[JsonSerializable(typeof(PaymentType))]
[JsonSerializable(typeof(InterestRateApplication))]
[JsonSerializable(typeof(LogEntryType))]
public partial class CreditJsonContext : JsonSerializerContext;
