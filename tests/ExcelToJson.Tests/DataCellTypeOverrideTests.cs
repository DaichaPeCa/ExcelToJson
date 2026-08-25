using ClosedXML.Excel;
using ExcelToJson.Core;
using System.Text.Json;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class DataCellTypeOverrideTests
{
    [Fact]
    public void UsesHeaderDefaultsAndAllowsPerCellScalarOverrides()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("scalar-overrides.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("array"))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "headerNumber", "headerText");
            TestWorkbook.SetType(root, 2, "number");

            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "10";
            root.Cell("C2").Value = "20";
            TestWorkbook.SetType(root, 2, 1, "auto");

            root.Cell("A3").Value = "R2";
            root.Cell("B3").Value = "30";
            root.Cell("C3").Value = "40";
            TestWorkbook.SetType(root, 3, 2, " text ");
            TestWorkbook.SetType(root, 3, 3, " NuMbEr ");

            root.Cell("A4").Value = "R3";
            root.Cell("B4").Value = "50";
            root.Cell("C4").Value = "60";
            TestWorkbook.SetType(root, 4, 2, "   ");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));

        Assert.Equal(10m, json.RootElement[0].GetProperty("headerNumber").GetDecimal());
        Assert.Equal("20", json.RootElement[0].GetProperty("headerText").GetString());
        Assert.Equal("30", json.RootElement[1].GetProperty("headerNumber").GetString());
        Assert.Equal(40m, json.RootElement[1].GetProperty("headerText").GetDecimal());
        Assert.Equal(50m, json.RootElement[2].GetProperty("headerNumber").GetDecimal());
    }

    [Fact]
    public void SupportsAllScalarTypesInDataCellNotes()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("all-scalar-types.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "number", "boolean", "date", "text");
            TestWorkbook.SetType(root, 5, "number");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = 12.5;
            root.Cell("C2").Value = "YES";
            root.Cell("D2").Value = new DateTime(2026, 8, 13);
            root.Cell("D2").Style.DateFormat.Format = "yyyy-mm-dd";
            root.Cell("E2").Value = 7;
            TestWorkbook.SetType(root, 2, 2, " NUMBER ");
            TestWorkbook.SetType(root, 2, 3, " BoOlEaN ");
            TestWorkbook.SetType(root, 2, 4, " date ");
            TestWorkbook.SetType(root, 2, 5, " TeXt ");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        JsonElement rootJson = json.RootElement;

        Assert.Equal(12.5m, rootJson.GetProperty("number").GetDecimal());
        Assert.True(rootJson.GetProperty("boolean").GetBoolean());
        Assert.Equal("2026-08-13", rootJson.GetProperty("date").GetString());
        Assert.Equal("7", rootJson.GetProperty("text").GetString());
    }

    [Fact]
    public void SupportsObjectArrayAndMultistageReferenceOverrides()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("reference-overrides.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "profile", "items");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "P1";
            root.Cell("C2").Value = "G1";
            TestWorkbook.SetType(root, 2, 2, " OBJECT:PROFILE ");
            TestWorkbook.SetType(root, 2, 3, " array:items ");

            IXLWorksheet profile = workbook.AddWorksheet("profile");
            profile.Cell("A1").Value = "ID";
            profile.Cell("B1").Value = "detail";
            profile.Cell("A2").Value = "P1";
            profile.Cell("B2").Value = "D1";
            TestWorkbook.SetType(profile, 2, 2, "object:detail");

            IXLWorksheet detail = workbook.AddWorksheet("detail");
            detail.Cell("A1").Value = "ID";
            detail.Cell("B1").Value = "name";
            detail.Cell("A2").Value = "D1";
            detail.Cell("B2").Value = "nested";

            IXLWorksheet items = workbook.AddWorksheet("items");
            items.Cell("A1").Value = "ID";
            items.Cell("B1").Value = "name";
            items.Cell("A2").Value = "G1";
            items.Cell("B2").Value = "first";
            items.Cell("A3").Value = "G1";
            items.Cell("B3").Value = "second";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));

        Assert.Equal("nested", json.RootElement.GetProperty("profile").GetProperty("detail").GetProperty("name").GetString());
        Assert.Equal(2, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void AppliesDataCellTypeOverrideToFormulaResult()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("formula-override.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").FormulaA1 = "1+1";
            TestWorkbook.SetType(root, 2, 2, "number");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(2m, json.RootElement.GetProperty("value").GetDecimal());
    }

    [Fact]
    public void ReportsInvalidDataCellTypeDefinitionsAtTheirCells()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("invalid-overrides.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "unknown", "emptyTarget", "missing", "excluded");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "x";
            root.Cell("C2").Value = "x";
            root.Cell("D2").Value = "x";
            root.Cell("E2").Value = "x";
            TestWorkbook.SetType(root, 2, 2, "auto");
            TestWorkbook.SetType(root, 2, 3, "object:");
            TestWorkbook.SetType(root, 2, 4, "object:missing");
            TestWorkbook.SetType(root, 2, 5, "array:_memo");
            _ = workbook.AddWorksheet("_memo");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "B2" && diagnostic.Message.Contains("未知のJSON型", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "C2" && diagnostic.Message.Contains("参照先シート名が空", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "D2" && diagnostic.Message.Contains("存在しません", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "E2" && diagnostic.Message.Contains("変換対象外", StringComparison.Ordinal));
    }

    [Fact]
    public void ValidatesOverrideDefinitionsOnEmptyUnreachableCells()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("unreachable-empty-override.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "root";

            IXLWorksheet unused = workbook.AddWorksheet("unused");
            unused.Cell("A1").Value = "ID";
            unused.Cell("B1").Value = "unknown";
            unused.Cell("C1").Value = "reference";
            unused.Cell("A2").Value = "U1";
            TestWorkbook.SetType(unused, 2, 2, "auto");
            TestWorkbook.SetType(unused, 2, 3, "object:missing");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "B2" && diagnostic.Sheet == "unused" && diagnostic.Message.Contains("未知のJSON型", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "C2" && diagnostic.Sheet == "unused" && diagnostic.Message.Contains("存在しません", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("omit")]
    [InlineData("null")]
    [InlineData("emptyString")]
    public void AppliesEmptyCellBehaviorBeforeResolvingReferenceOverride(string emptyCell)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File($"empty-{emptyCell}.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create(emptyCell: emptyCell))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "reference");
            root.Cell("A2").Value = "R1";
            TestWorkbook.SetType(root, 2, 2, "object:child");
            IXLWorksheet child = workbook.AddWorksheet("child");
            child.Cell("A1").Value = "ID";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));

        if (emptyCell == "omit")
        {
            Assert.False(json.RootElement.TryGetProperty("reference", out _));
        }
        else if (emptyCell == "null")
        {
            Assert.Equal(JsonValueKind.Null, json.RootElement.GetProperty("reference").ValueKind);
        }
        else
        {
            Assert.Equal(string.Empty, json.RootElement.GetProperty("reference").GetString());
        }
    }

    [Fact]
    public void TreatsCommentOnlyRowAsTerminatingEmptyRow()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("comment-only-row.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("array"))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, 2, "auto");
            root.Cell("A3").Value = "R1";
            root.Cell("B3").Value = "ignored";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(0, json.RootElement.GetArrayLength());
    }
}
