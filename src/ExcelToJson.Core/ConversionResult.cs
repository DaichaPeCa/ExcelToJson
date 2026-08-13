namespace ExcelToJson.Core;

public abstract record ConversionResult
{
    private ConversionResult()
    {
    }

    public sealed record Succeeded(string OutputPath) : ConversionResult;

    public sealed record Failed(IReadOnlyList<ConversionDiagnostic> Diagnostics) : ConversionResult;
}
