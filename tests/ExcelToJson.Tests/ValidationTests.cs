using ClosedXML.Excel;
using ExcelToJson.Core;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class ValidationTests
{
    [Fact]
    public void AggregatesIndependentStaticValidationErrors()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("invalid.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("invalid"))
        {
            IXLWorksheet setting = workbook.Worksheet("setting");
            setting.Cell("A4").Value = "unknown";
            setting.Cell("B4").Value = "x";
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name", "name", "reference");
            TestWorkbook.SetType(root, 4, "object:");
            root.Cell("B2").Value = "x";
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        string combined = string.Join(Environment.NewLine, result.Diagnostics);
        Assert.Contains("rootType", combined, StringComparison.Ordinal);
        Assert.Contains("未知のsettingキー", combined, StringComparison.Ordinal);
        Assert.Contains("重複", combined, StringComparison.Ordinal);
        Assert.Contains("参照先シート名が空", combined, StringComparison.Ordinal);
        Assert.Contains("IDは空", combined, StringComparison.Ordinal);
    }

    [Fact]
    public void TreatsIdsAndPropertiesAsCaseSensitiveButControlsAsCaseInsensitive()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("case.xlsx");
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet setting = workbook.AddWorksheet("SETTING");
            setting.Cell("A1").Value = " KEY ";
            setting.Cell("B1").Value = " VALUE ";
            setting.Cell("A2").Value = " ROOTTYPE ";
            setting.Cell("B2").Value = " OBJECT ";
            IXLWorksheet root = workbook.AddWorksheet("ROOT");
            root.Cell("A1").Value = " id ";
            root.Cell("B1").Value = "name";
            root.Cell("C1").Value = "Name";
            root.Cell("D1").Value = "lower";
            root.Cell("E1").Value = "upper";
            TestWorkbook.SetType(root, 4, "OBJECT:CHILD");
            TestWorkbook.SetType(root, 5, "object:child");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "a";
            root.Cell("C2").Value = "b";
            root.Cell("D2").Value = "id";
            root.Cell("E2").Value = "ID";
            IXLWorksheet child = workbook.AddWorksheet("Child");
            child.Cell("A1").Value = "ID";
            child.Cell("B1").Value = "value";
            child.Cell("A2").Value = "id";
            child.Cell("B2").Value = "lower";
            child.Cell("A3").Value = "ID";
            child.Cell("B3").Value = "upper";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        string json = File.ReadAllText(result.OutputPath);
        Assert.Contains("\"name\"", json, StringComparison.Ordinal);
        Assert.Contains("\"Name\"", json, StringComparison.Ordinal);
        Assert.Contains("\"lower\"", json, StringComparison.Ordinal);
        Assert.Contains("\"upper\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsObjectMultiplicityAndArrayZeroMatches()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("counts.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "one", "many");
            TestWorkbook.SetType(root, 2, "object:child");
            TestWorkbook.SetType(root, 3, "array:child");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "duplicate";
            root.Cell("C2").Value = "missing";
            IXLWorksheet child = workbook.AddWorksheet("child");
            child.Cell("A1").Value = "ID";
            child.Cell("A2").Value = "duplicate";
            child.Cell("A3").Value = "duplicate";
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        string combined = string.Join(Environment.NewLine, result.Diagnostics);
        Assert.Contains("実際: 2件", combined, StringComparison.Ordinal);
        Assert.Contains("一致する行がありません", combined, StringComparison.Ordinal);
    }

    [Fact]
    public void DetectsRecordCycleWithPath()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("cycle.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "self");
            TestWorkbook.SetType(root, 2, "object:root");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "R1";
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains("root[2, ID=R1] -> root[2, ID=R1]", result.Diagnostics.Single().Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RejectsDataOutsideSchemaAndReferencesToExcludedSheet()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("outside.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "reference");
            TestWorkbook.SetType(root, 2, "object:_memo");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "M1";
            root.Cell("C2").Value = "unexpected";
            _ = workbook.AddWorksheet("_memo");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        string combined = string.Join(Environment.NewLine, result.Diagnostics);
        Assert.Contains("定義されていない列", combined, StringComparison.Ordinal);
        Assert.Contains("変換対象外シート", combined, StringComparison.Ordinal);
    }

    [Fact]
    public void DoesNotOverwriteExistingJsonWhenConversionFails()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("protected.xlsx");
        string output = workspace.File("protected.json");
        File.WriteAllText(output, "original");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "number");
            TestWorkbook.SetType(root, 2, "number");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "not-a-number";
            workbook.SaveAs(input);
        }

        _ = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Equal("original", File.ReadAllText(output));
    }
}
