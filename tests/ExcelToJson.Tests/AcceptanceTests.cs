using ClosedXML.Excel;
using ExcelToJson.Core;
using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class AcceptanceTests
{
    [Theory]
    [InlineData("true", true)]
    [InlineData("FALSE", false)]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("YeS", true)]
    [InlineData("NO", false)]
    public void ConvertsAllBooleanTextRepresentations(string source, bool expected)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("boolean.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, "boolean");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = source;
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(expected, json.RootElement.GetProperty("value").GetBoolean());
    }

    [Fact]
    public void StopsAtFirstEmptyRow()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("termination.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("array"))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "before";
            root.Cell("A4").Value = "R2";
            root.Cell("B4").Value = "after";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(1, json.RootElement.GetArrayLength());
        Assert.Equal("before", json.RootElement[0].GetProperty("name").GetString());
    }

    [Fact]
    public void AllowsAcyclicHierarchyWithinSameSheet()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("tree.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "node");
            TestWorkbook.SetType(root, 2, "object:nodes");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "A";

            IXLWorksheet nodes = workbook.AddWorksheet("nodes");
            nodes.Cell("A1").Value = "ID";
            nodes.Cell("B1").Value = "name";
            nodes.Cell("C1").Value = "next";
            TestWorkbook.SetType(nodes, 3, "object:nodes");
            nodes.Cell("A2").Value = "A";
            nodes.Cell("B2").Value = "first";
            nodes.Cell("C2").Value = "B";
            nodes.Cell("A3").Value = "B";
            nodes.Cell("B3").Value = "second";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal("second", json.RootElement.GetProperty("node").GetProperty("next").GetProperty("name").GetString());
        Assert.False(json.RootElement.GetProperty("node").GetProperty("next").TryGetProperty("next", out _));
    }

    [Fact]
    public void AppliesCustomDateFormatsAndRejectsTimeOnlyInput()
    {
        using TestWorkspace workspace = new();
        string validInput = workspace.File("custom-date.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet setting = workbook.Worksheet("setting");
            setting.Cell("A4").Value = "dateInputFormat";
            setting.Cell("B4").Value = "yyyyMMdd HHmmss";
            setting.Cell("A5").Value = "dateOutputFormat";
            setting.Cell("B5").Value = "yyyy/MM/dd HH:mm";
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, "date");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "20260813 143025";
            workbook.SaveAs(validInput);
        }

        ConversionResult.Succeeded valid = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(validInput));
        using (JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(valid.OutputPath)))
        {
            Assert.Equal("2026/08/13 14:30", json.RootElement.GetProperty("value").GetString());
        }

        string invalidInput = workspace.File("time-only.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, "date");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "14:30";
            workbook.SaveAs(invalidInput);
        }

        ConversionResult.Failed invalid = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(invalidInput));
        Assert.Contains("dateへ変換できません", invalid.Diagnostics.Single().Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ConvertsOffsetInputToLocalWallClockTime()
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            using TestWorkspace workspace = new();
            string input = workspace.File("offset.xlsx");
            using (XLWorkbook workbook = TestWorkbook.Create())
            {
                IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
                TestWorkbook.SetType(root, 2, "date");
                root.Cell("A2").Value = "R1";
                root.Cell("B2").Value = "2026-08-13T14:30:00+09:00";
                workbook.SaveAs(input);
            }

            ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
            DateTime expected = DateTimeOffset.Parse("2026-08-13T14:30:00+09:00", CultureInfo.InvariantCulture).LocalDateTime;
            using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
            Assert.Equal(expected.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture), json.RootElement.GetProperty("value").GetString());
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void IgnoresInvalidValuesOnUnreachableRowsButValidatesTheirStructure()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("unreachable.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "root";
            IXLWorksheet unused = workbook.AddWorksheet("unused");
            unused.Cell("A1").Value = "ID";
            unused.Cell("B1").Value = "number";
            TestWorkbook.SetType(unused, 2, "number");
            unused.Cell("A2").Value = "U1";
            unused.Cell("B2").Value = "invalid-number";
            workbook.SaveAs(input);
        }

        _ = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
    }

    [Fact]
    public void RejectsUnavailableFormulaResult()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("formula-error.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, "number");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").FormulaA1 = "UNSUPPORTEDFUNCTION(1)";
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains("数式", string.Join(Environment.NewLine, result.Diagnostics), StringComparison.Ordinal);
    }

    [Fact]
    public void UsesExcelCachedResultWhenWorkbookContainsOne()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("formula-cache.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, "number");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").FormulaA1 = "UNSUPPORTEDFUNCTION(1)";
            workbook.SaveAs(input);
        }

        SetCachedFormulaValue(input, "xl/worksheets/sheet2.xml", "B2", "42");

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(42m, json.RootElement.GetProperty("value").GetDecimal());
    }

    [Fact]
    public void TreatsFormulaEmptyStringsAsTerminatingEmptyRow()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("formula-empty-row.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("array"))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").FormulaA1 = "\"\"";
            root.Cell("B2").FormulaA1 = "\"\"";
            root.Cell("A3").Value = "R1";
            root.Cell("B3").Value = "ignored";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(0, json.RootElement.GetArrayLength());
    }

    [Fact]
    public void OverwritesExistingJsonAfterSuccessfulConversion()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("overwrite.xlsx");
        string output = workspace.File("overwrite.json");
        File.WriteAllText(output, "old");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "new";
            workbook.SaveAs(input);
        }

        _ = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains("new", File.ReadAllText(output), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("missing-setting", "settingシートが存在しません")]
    [InlineData("missing-root", "rootシートが存在しません")]
    [InlineData("setting-header", "A1=key")]
    [InlineData("duplicate-setting", "settingキーが重複")]
    [InlineData("invalid-empty", "emptyCellは")]
    [InlineData("root-count", "ちょうど1行")]
    [InlineData("unknown-type", "未知のJSON型")]
    [InlineData("missing-reference", "参照先シート")]
    [InlineData("invalid-boolean", "booleanへ変換")]
    [InlineData("date-as-number", "日付・日時セルはnumber")]
    public void RejectsMajorRequirementErrors(string scenario, string expectedMessage)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File($"{scenario}.xlsx");
        using (XLWorkbook workbook = CreateInvalidWorkbook(scenario))
        {
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(expectedMessage, string.Join(Environment.NewLine, result.Diagnostics), StringComparison.Ordinal);
    }

    private static XLWorkbook CreateInvalidWorkbook(string scenario)
    {
        XLWorkbook workbook = TestWorkbook.Create();
        IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
        root.Cell("A2").Value = "R1";
        root.Cell("B2").Value = "value";

        switch (scenario)
        {
            case "missing-setting":
                workbook.Worksheet("setting").Delete();
                break;
            case "missing-root":
                root.Delete();
                break;
            case "setting-header":
                workbook.Worksheet("setting").Cell("A1").Value = "invalid";
                break;
            case "duplicate-setting":
                workbook.Worksheet("setting").Cell("A4").Value = "ROOTTYPE";
                workbook.Worksheet("setting").Cell("B4").Value = "array";
                break;
            case "invalid-empty":
                workbook.Worksheet("setting").Cell("B3").Value = "invalid";
                break;
            case "root-count":
                root.Cell("A3").Value = "R2";
                break;
            case "unknown-type":
                TestWorkbook.SetType(root, 2, "auto");
                break;
            case "missing-reference":
                TestWorkbook.SetType(root, 2, "object:missing");
                break;
            case "invalid-boolean":
                TestWorkbook.SetType(root, 2, "boolean");
                root.Cell("B2").Value = "maybe";
                break;
            case "date-as-number":
                TestWorkbook.SetType(root, 2, "number");
                root.Cell("B2").Value = new DateTime(2026, 8, 13);
                root.Cell("B2").Style.DateFormat.Format = "yyyy-mm-dd";
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(scenario));
        }

        return workbook;
    }

    private static void SetCachedFormulaValue(string workbookPath, string entryPath, string cellAddress, string value)
    {
        using ZipArchive archive = ZipFile.Open(workbookPath, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(entryPath)
            ?? throw new InvalidOperationException($"Worksheet entry not found: {entryPath}");
        XDocument document;
        using (Stream input = entry.Open())
        {
            document = XDocument.Load(input);
        }

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XElement cell = document.Descendants(spreadsheet + "c")
            .Single(element => string.Equals((string?)element.Attribute("r"), cellAddress, StringComparison.Ordinal));
        XElement? cachedValue = cell.Element(spreadsheet + "v");
        if (cachedValue is null)
        {
            cell.Add(new XElement(spreadsheet + "v", value));
        }
        else
        {
            cachedValue.Value = value;
        }

        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(entryPath);
        using Stream output = replacement.Open();
        document.Save(output);
    }
}
