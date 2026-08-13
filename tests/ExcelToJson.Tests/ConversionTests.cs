using ClosedXML.Excel;
using ExcelToJson.Core;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class ConversionTests
{
    [Fact]
    public void ConvertsObjectWithAllTypesReferencesAndFormula()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("sample.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "name", "profile", "items", "active", "score");
            TestWorkbook.SetType(root, 3, " object:PROFILE ");
            TestWorkbook.SetType(root, 4, "array:items");
            TestWorkbook.SetType(root, 5, "boolean");
            TestWorkbook.SetType(root, 6, "number");
            root.Cell("A2").Value = "R001";
            root.Cell("B2").Value = " Alice ";
            root.Cell("C2").Value = "P001";
            root.Cell("D2").Value = "G001";
            root.Cell("E2").Value = "YES";
            root.Cell("F2").FormulaA1 = "1+2";

            IXLWorksheet profile = workbook.AddWorksheet("profile");
            profile.Cell("A1").Value = "ID";
            profile.Cell("B1").Value = "age";
            profile.Cell("C1").Value = "birthday";
            TestWorkbook.SetType(profile, 2, "number");
            TestWorkbook.SetType(profile, 3, "date");
            profile.Cell("A2").Value = "P001";
            profile.Cell("B2").Value = 30;
            profile.Cell("C2").Value = new DateTime(1996, 5, 10);
            profile.Cell("C2").Style.DateFormat.Format = "yyyy-mm-dd";

            IXLWorksheet items = workbook.AddWorksheet("Items");
            items.Cell("A1").Value = "ID";
            items.Cell("B1").Value = "name";
            items.Cell("C1").Value = "quantity";
            TestWorkbook.SetType(items, 3, "number");
            items.Cell("A2").Value = "G001";
            items.Cell("B2").Value = "Apple";
            items.Cell("C2").Value = 2;
            items.Cell("A3").Value = "G001";
            items.Cell("B3").Value = "Orange";
            items.Cell("C3").Value = 3;
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        byte[] bytes = File.ReadAllBytes(result.OutputPath);
        Assert.False(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble));

        using JsonDocument json = JsonDocument.Parse(bytes);
        JsonElement rootJson = json.RootElement;
        Assert.Equal(" Alice ", rootJson.GetProperty("name").GetString());
        Assert.Equal(30m, rootJson.GetProperty("profile").GetProperty("age").GetDecimal());
        Assert.Equal("1996-05-10", rootJson.GetProperty("profile").GetProperty("birthday").GetString());
        Assert.Equal(2, rootJson.GetProperty("items").GetArrayLength());
        Assert.True(rootJson.GetProperty("active").GetBoolean());
        Assert.Equal(3m, rootJson.GetProperty("score").GetDecimal());
        Assert.Contains(Environment.NewLine, Encoding.UTF8.GetString(bytes), StringComparison.Ordinal);
        Assert.Equal(["name", "profile", "items", "active", "score"], rootJson.EnumerateObject().Select(property => property.Name));
    }

    [Fact]
    public void ConvertsEmptyRootArray()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("empty.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create("array"))
        {
            _ = TestWorkbook.AddRoot(workbook, "name");
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Equal("[]", File.ReadAllText(result.OutputPath).Trim());
    }

    [Theory]
    [InlineData("omit", "{}")]
    [InlineData("null", "{\"value\":null}")]
    [InlineData("emptyString", "{\"value\":\"\"}")]
    public void AppliesEmptyCellBeforeTypeConversion(string behavior, string expectedCompactJson)
    {
        using TestWorkspace workspace = new();
        string input = workspace.File($"empty-{behavior}.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create(emptyCell: behavior))
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
            TestWorkbook.SetType(root, 2, "number");
            root.Cell("A2").Value = "R1";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument actual = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        using JsonDocument expected = JsonDocument.Parse(expectedCompactJson);
        Assert.True(JsonElement.DeepEquals(expected.RootElement, actual.RootElement));
    }

    [Fact]
    public void ParsesDecimalUsingCurrentCultureIncludingThousandsAndExponent()
    {
        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        CultureInfo originalUiCulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("de-DE");
            using TestWorkspace workspace = new();
            string input = workspace.File("culture.xlsx");
            using (XLWorkbook workbook = TestWorkbook.Create())
            {
                IXLWorksheet root = TestWorkbook.AddRoot(workbook, "value");
                TestWorkbook.SetType(root, 2, "number");
                root.Cell("A2").Value = "R1";
                root.Cell("B2").Value = " 1.234,5E+1 ";
                workbook.SaveAs(input);
            }

            ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
            using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
            Assert.Equal(12345m, json.RootElement.GetProperty("value").GetDecimal());
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public void PreservesMidnightAsDateTimeWhenDisplayFormatContainsTime()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("date.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "dateOnly", "midnight", "textMidnight");
            TestWorkbook.SetType(root, 2, "date");
            TestWorkbook.SetType(root, 3, "date");
            TestWorkbook.SetType(root, 4, "date");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = new DateTime(2026, 8, 13);
            root.Cell("B2").Style.DateFormat.Format = "yyyy-mm-dd";
            root.Cell("C2").Value = new DateTime(2026, 8, 13);
            root.Cell("C2").Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
            root.Cell("D2").Value = "2026/08/13 00:00:00";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal("2026-08-13", json.RootElement.GetProperty("dateOnly").GetString());
        Assert.Equal("2026-08-13T00:00:00", json.RootElement.GetProperty("midnight").GetString());
        Assert.Equal("2026-08-13T00:00:00", json.RootElement.GetProperty("textMidnight").GetString());
    }

    [Fact]
    public void UsesFormattedDisplayStringForReferenceId()
    {
        using TestWorkspace workspace = new();
        string input = workspace.File("display-id.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "child");
            TestWorkbook.SetType(root, 2, "object:child");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = 1;
            root.Cell("B2").Style.NumberFormat.Format = "000";

            IXLWorksheet child = workbook.AddWorksheet("child");
            child.Cell("A1").Value = "ID";
            child.Cell("B1").Value = "name";
            child.Cell("A2").Value = "001";
            child.Cell("B2").Value = "matched";
            workbook.SaveAs(input);
        }

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal("matched", json.RootElement.GetProperty("child").GetProperty("name").GetString());
    }
}
