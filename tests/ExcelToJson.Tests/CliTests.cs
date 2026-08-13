using ClosedXML.Excel;
using ExcelToJson.Cli;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class CliTests
{
    [Fact]
    public void RejectsInvalidArgumentCountOnStandardError()
    {
        using StringWriter stdout = new();
        using StringWriter stderr = new();

        int exitCode = Program.Run([], stdout, stderr);

        Assert.Equal(1, exitCode);
        Assert.Equal(string.Empty, stdout.ToString());
        Assert.Contains("使用法", stderr.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void ReportsOutputPathOnStandardOutput()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("cli.XLSX");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "Alice";
            workbook.SaveAs(input);
        }

        using StringWriter stdout = new();
        using StringWriter stderr = new();
        int exitCode = Program.Run([input], stdout, stderr);

        Assert.Equal(0, exitCode);
        Assert.Equal(string.Empty, stderr.ToString());
        Assert.Equal(Path.ChangeExtension(Path.GetFullPath(input), ".json"), stdout.ToString().Trim());
    }
}
