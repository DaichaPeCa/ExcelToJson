using ExcelToJson.Core.Infrastructure;

namespace ExcelToJson.Core;

public sealed class ExcelToJsonConverter
{
    public static ConversionResult ConvertFile(string inputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);

        try
        {
            string fullInputPath = Path.GetFullPath(inputPath);
            WorkbookReadResult readResult = ClosedXmlWorkbookReader.Read(fullInputPath);
            if (readResult.Diagnostics.Count != 0)
            {
                return new ConversionResult.Failed(readResult.Diagnostics);
            }

            WorkbookTransformResult transformResult = WorkbookTransformer.Transform(readResult.Workbook!);
            if (transformResult.Diagnostics.Count != 0)
            {
                return new ConversionResult.Failed(transformResult.Diagnostics);
            }

            byte[] json = JsonOutput.Serialize(transformResult.Root!);
            string outputPath = Path.ChangeExtension(fullInputPath, ".json");
            SafeFileWriter.Write(outputPath, json);
            return new ConversionResult.Succeeded(outputPath);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return new ConversionResult.Failed([
                new ConversionDiagnostic($"ファイルを処理できませんでした。{exception.Message}"),
            ]);
        }
        catch (Exception exception)
        {
            return new ConversionResult.Failed([
                new ConversionDiagnostic($"変換中に予期しないエラーが発生しました。{exception.Message}"),
            ]);
        }
    }
}
