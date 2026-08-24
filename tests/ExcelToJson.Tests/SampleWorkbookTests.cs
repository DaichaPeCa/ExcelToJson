using ClosedXML.Excel;
using ExcelToJson.Core;
using System.Text;
using System.Text.Json;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class SampleWorkbookTests
{
    [Fact]
    public void RepositorySampleMatchesDocumentedConversionExample()
    {
        string source = SampleAsset("sample.xlsx");
        using (XLWorkbook workbook = new(source))
        {
            Assert.Equal(["setting", "root", "profile", "items"], workbook.Worksheets.Select(sheet => sheet.Name));
            AssertSetting(workbook.Worksheet("setting"), "object");

            IXLWorksheet root = workbook.Worksheet("root");
            Assert.Equal("object:profile", CommentText(root.Cell("C1")));
            Assert.Equal("array:items", CommentText(root.Cell("D1")));

            IXLWorksheet profile = workbook.Worksheet("profile");
            Assert.Equal("number", CommentText(profile.Cell("B1")));
            Assert.Equal("date", CommentText(profile.Cell("C1")));

            IXLWorksheet items = workbook.Worksheet("items");
            Assert.Equal("number", CommentText(items.Cell("C1")));
        }

        using TestWorkspace workspace = new();
        string input = workspace.File("sample.xlsx");
        File.Copy(source, input);

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        byte[] bytes = File.ReadAllBytes(result.OutputPath);
        Assert.False(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble));

        using JsonDocument json = JsonDocument.Parse(bytes);
        JsonElement rootJson = json.RootElement;
        Assert.Equal(["name", "profile", "items"], rootJson.EnumerateObject().Select(property => property.Name));
        Assert.Equal("Alice", rootJson.GetProperty("name").GetString());
        Assert.Equal(30m, rootJson.GetProperty("profile").GetProperty("age").GetDecimal());
        Assert.Equal("1996-05-10", rootJson.GetProperty("profile").GetProperty("birthday").GetString());

        JsonElement itemsJson = rootJson.GetProperty("items");
        Assert.Equal(2, itemsJson.GetArrayLength());
        Assert.Equal("Apple", itemsJson[0].GetProperty("name").GetString());
        Assert.Equal(2m, itemsJson[0].GetProperty("quantity").GetDecimal());
        Assert.Equal("Orange", itemsJson[1].GetProperty("name").GetString());
        Assert.Equal(3m, itemsJson[1].GetProperty("quantity").GetDecimal());
    }

    [Fact]
    public void RepositoryTemplateIsAnEmptyValidArrayWorkbook()
    {
        string source = SampleAsset("template.xlsx");
        using (XLWorkbook workbook = new(source))
        {
            Assert.Equal(["_guide", "setting", "root"], workbook.Worksheets.Select(sheet => sheet.Name));
            AssertSetting(workbook.Worksheet("setting"), "array");
            Assert.False(workbook.Worksheet("setting").Cell("A4").IsEmpty());
            Assert.True(workbook.Worksheet("setting").Cell("B4").IsEmpty());
            Assert.False(workbook.Worksheet("setting").Cell("A5").IsEmpty());
            Assert.True(workbook.Worksheet("setting").Cell("B5").IsEmpty());

            IXLWorksheet root = workbook.Worksheet("root");
            Assert.Equal("ID", root.Cell("A1").GetString());
            Assert.Equal(1, root.LastRowUsed()!.RowNumber());
            Assert.False(workbook.Worksheet("_guide").Cell("A1").IsEmpty());
        }

        using TestWorkspace workspace = new();
        string input = workspace.File("template.xlsx");
        File.Copy(source, input);

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        Assert.Equal("[]", File.ReadAllText(result.OutputPath).Trim());
    }

    private static string SampleAsset(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "samples", fileName);

    private static void AssertSetting(IXLWorksheet setting, string rootType)
    {
        Assert.Equal("key", setting.Cell("A1").GetString());
        Assert.Equal("value", setting.Cell("B1").GetString());
        Assert.Equal("rootType", setting.Cell("A2").GetString());
        Assert.Equal(rootType, setting.Cell("B2").GetString());
        Assert.Equal("emptyCell", setting.Cell("A3").GetString());
        Assert.Equal("omit", setting.Cell("B3").GetString());
    }

    private static string CommentText(IXLCell cell)
    {
        Assert.True(cell.HasComment);
        return cell.GetComment().Text;
    }
}
