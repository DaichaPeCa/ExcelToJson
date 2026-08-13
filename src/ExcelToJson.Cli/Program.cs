using ExcelToJson.Core;
using System.Text;

namespace ExcelToJson.Cli;

public static class Program
{
    public static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return Run(args, Console.Out, Console.Error);
    }

    public static int Run(string[] args, TextWriter standardOutput, TextWriter standardError)
    {
        ArgumentNullException.ThrowIfNull(args);
        ArgumentNullException.ThrowIfNull(standardOutput);
        ArgumentNullException.ThrowIfNull(standardError);

        if (args.Length != 1)
        {
            standardError.WriteLine("入力Excelファイルを1つ指定してください。");
            standardError.WriteLine("使用法: ExcelToJson.exe <input.xlsx>");
            return 1;
        }

        if (!string.Equals(Path.GetExtension(args[0]), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            standardError.WriteLine("入力ファイルの拡張子は .xlsx でなければなりません。");
            return 1;
        }

        ConversionResult result = ExcelToJsonConverter.ConvertFile(args[0]);
        switch (result)
        {
            case ConversionResult.Succeeded succeeded:
                standardOutput.WriteLine(succeeded.OutputPath);
                return 0;
            case ConversionResult.Failed failed:
                foreach (ConversionDiagnostic diagnostic in failed.Diagnostics)
                {
                    standardError.WriteLine(diagnostic);
                }

                return 1;
            default:
                standardError.WriteLine("変換結果を判定できませんでした。");
                return 1;
        }
    }
}
