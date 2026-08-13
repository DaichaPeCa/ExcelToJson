namespace ExcelToJson.Core;

public sealed record ConversionDiagnostic(
    string Message,
    string? Sheet = null,
    string? Cell = null,
    string? SettingKey = null)
{
    public override string ToString()
    {
        List<string> locations = [];
        if (Sheet is not null)
        {
            locations.Add($"シート '{Sheet}'");
        }

        if (Cell is not null)
        {
            locations.Add($"セル {Cell}");
        }

        if (SettingKey is not null)
        {
            locations.Add($"設定 '{SettingKey}'");
        }

        return locations.Count == 0
            ? Message
            : $"{string.Join("、", locations)}: {Message}";
    }
}
