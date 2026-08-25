using ClosedXML.Excel;
using ExcelToJson.Core;
using System.Text.Json;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class ScalarArrayTests
{
    [Fact]
    public void ConvertsNestedScalarArraysAndSourceCellOverride()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("nested-scalar-arrays.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(
                workbook,
                "nickNames",
                "numbers",
                "flags",
                "dates",
                "overridden");
            TestWorkbook.SetType(root, 2, " ScAlAr-ArRaY : NICKNAMES ");
            TestWorkbook.SetType(root, 3, "scalar-array:numbers");
            TestWorkbook.SetType(root, 4, "scalar-array:flags");
            TestWorkbook.SetType(root, 5, "scalar-array:dates");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = " N1 ";
            root.Cell("C2").Value = "N2";
            root.Cell("D2").Value = "N3";
            root.Cell("E2").Value = "N4";
            root.Cell("F2").Value = "N5";
            TestWorkbook.SetType(root, 2, 6, " scalar-array:overridden ");

            IXLWorksheet nickNames = AddScalarSheet(workbook, "nickNames", null, " VaLuE ");
            SetScalarRow(nickNames, 2, "N1", "Allie");
            SetScalarRow(nickNames, 3, "N1", "Ali");
            SetScalarRow(nickNames, 4, "n1", "case-sensitive");

            IXLWorksheet numbers = AddScalarSheet(workbook, "numbers", "number");
            numbers.Cell("A2").Value = "N2";
            numbers.Cell("B2").FormulaA1 = "1+1";
            SetScalarRow(numbers, 3, "N2", 2);

            IXLWorksheet flags = AddScalarSheet(workbook, "flags", "boolean");
            SetScalarRow(flags, 2, "N3", "yes");
            SetScalarRow(flags, 3, "N3", 0);

            IXLWorksheet dates = AddScalarSheet(workbook, "dates", "date");
            SetScalarRow(dates, 2, "N4", new DateTime(1996, 5, 10), "yyyy-mm-dd");
            SetScalarRow(dates, 3, "N4", new DateTime(2026, 8, 13, 0, 0, 0), "yyyy-mm-dd hh:mm:ss");

            IXLWorksheet overridden = AddScalarSheet(workbook, "overridden");
            SetScalarRow(overridden, 2, "N5", "cell override");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        JsonElement rootJson = json.RootElement;

        Assert.Equal(["Allie", "Ali"], rootJson.GetProperty("nickNames").EnumerateArray().Select(item => item.GetString()));
        Assert.Equal([2m, 2m], rootJson.GetProperty("numbers").EnumerateArray().Select(item => item.GetDecimal()));
        Assert.Equal([true, false], rootJson.GetProperty("flags").EnumerateArray().Select(item => item.GetBoolean()));
        Assert.Equal(
            ["1996-05-10", "2026-08-13T00:00:00"],
            rootJson.GetProperty("dates").EnumerateArray().Select(item => item.GetString()));
        Assert.Equal("cell override", rootJson.GetProperty("overridden")[0].GetString());
    }

    [Theory]
    [InlineData("text")]
    [InlineData("number")]
    [InlineData("boolean")]
    [InlineData("date")]
    public void ConvertsScalarRootForEveryScalarType(string type)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File($"scalar-root-{type}.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("scalar-array"))
        {
            IXLWorksheet root = AddScalarSheet(workbook, "root", type == "text" ? null : type);
            root.Cell("A2").Value = "R1";
            if (type == "text")
            {
                root.Cell("B2").Value = "Alpha";
            }
            else if (type == "number")
            {
                root.Cell("B2").Value = 12.5;
            }
            else if (type == "boolean")
            {
                root.Cell("B2").Value = true;
            }
            else
            {
                root.Cell("B2").Value = new DateTime(2026, 8, 13);
                root.Cell("B2").Style.DateFormat.Format = "yyyy-mm-dd";
            }

            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        JsonElement value = json.RootElement[0];

        if (type == "text")
        {
            Assert.Equal("Alpha", value.GetString());
        }
        else if (type == "number")
        {
            Assert.Equal(12.5m, value.GetDecimal());
        }
        else if (type == "boolean")
        {
            Assert.True(value.GetBoolean());
        }
        else
        {
            Assert.Equal("2026-08-13", value.GetString());
        }
    }

    [Fact]
    public void ConvertsEmptyScalarRootToEmptyArray()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("empty-scalar-root.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("scalar-array"))
        {
            _ = AddScalarSheet(workbook, "root");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Equal("[]", File.ReadAllText(result.OutputPath).Trim());
    }

    [Theory]
    [InlineData("omit")]
    [InlineData("null")]
    [InlineData("emptyString")]
    public void AppliesEmptyCellBehaviorToScalarArrayElements(string emptyCell)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File($"empty-scalar-element-{emptyCell}.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create(emptyCell: emptyCell))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "values");
            TestWorkbook.SetType(root, 2, "scalar-array:values");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "V1";

            IXLWorksheet values = AddScalarSheet(workbook, "values");
            values.Cell("A2").Value = "V1";
            SetScalarRow(values, 3, "V1", "kept");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        JsonElement valuesJson = json.RootElement.GetProperty("values");

        if (emptyCell == "omit")
        {
            Assert.Equal(["kept"], valuesJson.EnumerateArray().Select(item => item.GetString()));
        }
        else if (emptyCell == "null")
        {
            Assert.Equal(JsonValueKind.Null, valuesJson[0].ValueKind);
            Assert.Equal("kept", valuesJson[1].GetString());
        }
        else
        {
            Assert.Equal(string.Empty, valuesJson[0].GetString());
            Assert.Equal("kept", valuesJson[1].GetString());
        }
    }

    [Theory]
    [InlineData("omit")]
    [InlineData("null")]
    [InlineData("emptyString")]
    public void AppliesEmptyCellBehaviorBeforeResolvingScalarArray(string emptyCell)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File($"empty-scalar-reference-{emptyCell}.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create(emptyCell: emptyCell))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "values");
            TestWorkbook.SetType(root, 2, "scalar-array:values");
            root.Cell("A2").Value = "R1";
            _ = AddScalarSheet(workbook, "values");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));

        if (emptyCell == "omit")
        {
            Assert.False(json.RootElement.TryGetProperty("values", out _));
        }
        else if (emptyCell == "null")
        {
            Assert.Equal(JsonValueKind.Null, json.RootElement.GetProperty("values").ValueKind);
        }
        else
        {
            Assert.Equal(string.Empty, json.RootElement.GetProperty("values").GetString());
        }
    }

    [Fact]
    public void RejectsScalarArrayWithNoMatchingRows()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("missing-scalar-array-id.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "values");
            TestWorkbook.SetType(root, 2, "scalar-array:values");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "missing";
            _ = AddScalarSheet(workbook, "values");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Cell == "B2"
            && diagnostic.Message.Contains("一致する行がありません", StringComparison.Ordinal));
    }

    [Fact]
    public void ValidatesScalarArraySheetStructureAndElementType()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("invalid-scalar-array-sheets.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "missingValue", "wrongHeader", "extra", "nested", "override");
            string[] targets = ["missingValue", "wrongHeader", "extra", "nested", "override"];
            for (int index = 0; index < targets.Length; index++)
            {
                TestWorkbook.SetType(root, index + 2, $"scalar-array:{targets[index]}");
                root.Cell(2, index + 2).Value = $"V{index + 1}";
            }

            root.Cell("A2").Value = "R1";

            IXLWorksheet missingValue = workbook.AddWorksheet("missingValue");
            missingValue.Cell("A1").Value = "ID";
            missingValue.Cell("A2").Value = "V1";

            IXLWorksheet wrongHeader = AddScalarSheet(workbook, "wrongHeader", valueHeader: "item");
            SetScalarRow(wrongHeader, 2, "V2", "value");

            IXLWorksheet extra = AddScalarSheet(workbook, "extra");
            extra.Cell("C1").Value = "other";
            SetScalarRow(extra, 2, "V3", "value");

            IXLWorksheet nested = AddScalarSheet(workbook, "nested", "object:child");
            SetScalarRow(nested, 2, "V4", "C1");
            IXLWorksheet child = workbook.AddWorksheet("child");
            child.Cell("A1").Value = "ID";
            child.Cell("B1").Value = "name";
            SetScalarRow(child, 2, "C1", "child");

            IXLWorksheet overrideSheet = AddScalarSheet(workbook, "override");
            SetScalarRow(overrideSheet, 2, "V5", "10");
            TestWorkbook.SetType(overrideSheet, 2, 2, "number");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Sheet == "missingValue" && diagnostic.Cell == "B1");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Sheet == "wrongHeader" && diagnostic.Cell == "B1");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Sheet == "extra" && diagnostic.Cell == "C1");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Sheet == "nested"
            && diagnostic.Cell == "B1"
            && diagnostic.Message.Contains("text、number、boolean、date", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Sheet == "override"
            && diagnostic.Cell == "B2"
            && diagnostic.Message.Contains("上書きできません", StringComparison.Ordinal));
    }

    [Fact]
    public void ValidatesMissingAndExcludedScalarArrayTargets()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("invalid-scalar-array-targets.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "missing", "excluded");
            TestWorkbook.SetType(root, 2, "scalar-array:missing");
            TestWorkbook.SetType(root, 3, "scalar-array:_guide");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "M1";
            root.Cell("C2").Value = "G1";
            _ = workbook.AddWorksheet("_guide");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "B1" && diagnostic.Message.Contains("存在しません", StringComparison.Ordinal));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Cell == "C1" && diagnostic.Message.Contains("変換対象外", StringComparison.Ordinal));
    }

    [Fact]
    public void ValidatesScalarArrayDefinitionOnEmptyUnreachableCell()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("unreachable-scalar-array.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "root";

            IXLWorksheet unused = workbook.AddWorksheet("unused");
            unused.Cell("A1").Value = "ID";
            unused.Cell("B1").Value = "reference";
            unused.Cell("A2").Value = "U1";
            TestWorkbook.SetType(unused, 2, 2, "scalar-array:invalidTarget");

            IXLWorksheet invalidTarget = AddScalarSheet(workbook, "invalidTarget", valueHeader: "item");
            SetScalarRow(invalidTarget, 2, "V1", "value");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Sheet == "invalidTarget"
            && diagnostic.Cell == "B1"
            && diagnostic.Message.Contains("valueヘッダー", StringComparison.Ordinal));
    }

    [Fact]
    public void RejectsScalarArrayWithEmptyTargetName()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("empty-scalar-array-target.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "values");
            TestWorkbook.SetType(root, 2, "scalar-array:");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "V1";
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Cell == "B1"
            && diagnostic.Message.Contains("参照先シート名が空", StringComparison.Ordinal));
    }

    [Fact]
    public void ReportsScalarElementConversionFailureAtValueCell()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("invalid-scalar-element.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "values");
            TestWorkbook.SetType(root, 2, "scalar-array:values");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "V1";

            IXLWorksheet values = AddScalarSheet(workbook, "values", "number");
            SetScalarRow(values, 2, "V1", "not-a-number");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Sheet == "values"
            && diagnostic.Cell == "B2"
            && diagnostic.Message.Contains("numberへ変換できません", StringComparison.Ordinal));
    }

    [Fact]
    public void ValidatesScalarRootStructureAndRejectsValueCellOverride()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("invalid-scalar-root.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("scalar-array"))
        {
            IXLWorksheet root = AddScalarSheet(workbook, "root", valueHeader: "item");
            root.Cell("C1").Value = "extra";
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "10";
            TestWorkbook.SetType(root, 2, 2, "number");
            workbook.SaveAs(input);
        }

        ConversionResult.Failed result = Assert.IsType<ConversionResult.Failed>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Sheet == "root" && diagnostic.Cell == "B1");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Sheet == "root" && diagnostic.Cell == "C1");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Sheet == "root"
            && diagnostic.Cell == "B2"
            && diagnostic.Message.Contains("上書きできません", StringComparison.Ordinal));
    }

    private static IXLWorksheet AddScalarSheet(
        XLWorkbook workbook,
        string name,
        string? type = null,
        string valueHeader = "value")
    {
        IXLWorksheet sheet = workbook.AddWorksheet(name);
        sheet.Cell("A1").Value = "ID";
        sheet.Cell("B1").Value = valueHeader;
        if (type is not null)
        {
            TestWorkbook.SetType(sheet, 2, type);
        }

        return sheet;
    }

    private static void SetScalarRow(
        IXLWorksheet sheet,
        int row,
        string id,
        XLCellValue value,
        string? numberFormat = null)
    {
        sheet.Cell(row, 1).Value = id;
        sheet.Cell(row, 2).Value = value;
        if (numberFormat is not null)
        {
            sheet.Cell(row, 2).Style.NumberFormat.Format = numberFormat;
        }
    }
}
