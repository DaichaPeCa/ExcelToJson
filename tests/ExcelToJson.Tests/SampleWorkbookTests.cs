using ClosedXML.Excel;
using ExcelToJson.Core;
using System.IO.Compression;
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
            Assert.Equal(["setting", "root", "profile", "items", "nickNames"], workbook.Worksheets.Select(sheet => sheet.Name));
            AssertSetting(workbook.Worksheet("setting"), "object");

            IXLWorksheet root = workbook.Worksheet("root");
            Assert.Equal("object:profile", CommentText(root.Cell("C1")));
            Assert.Equal("array:items", CommentText(root.Cell("D1")));
            Assert.Equal("scalar-array:nickNames", CommentText(root.Cell("E1")));

            IXLWorksheet profile = workbook.Worksheet("profile");
            Assert.Equal("number", CommentText(profile.Cell("B1")));
            Assert.Equal("date", CommentText(profile.Cell("C1")));

            IXLWorksheet items = workbook.Worksheet("items");
            Assert.Equal("number", CommentText(items.Cell("C1")));

            IXLWorksheet nickNames = workbook.Worksheet("nickNames");
            Assert.Equal("ID", nickNames.Cell("A1").GetString());
            Assert.Equal("value", nickNames.Cell("B1").GetString());
            Assert.False(nickNames.Cell("B1").HasComment);
        }

        AssertLegacyCommentsOnly(source);

        using TestWorkspace workspace = new();
        string input = workspace.File("sample.xlsx");
        File.Copy(source, input);

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        byte[] bytes = File.ReadAllBytes(result.OutputPath);
        Assert.False(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble));

        using JsonDocument json = JsonDocument.Parse(bytes);
        JsonElement rootJson = json.RootElement;
        Assert.Equal(["name", "profile", "items", "nickNames"], rootJson.EnumerateObject().Select(property => property.Name));
        Assert.Equal("Alice", rootJson.GetProperty("name").GetString());
        Assert.Equal(30m, rootJson.GetProperty("profile").GetProperty("age").GetDecimal());
        Assert.Equal("1996-05-10", rootJson.GetProperty("profile").GetProperty("birthday").GetString());

        JsonElement itemsJson = rootJson.GetProperty("items");
        Assert.Equal(2, itemsJson.GetArrayLength());
        Assert.Equal("Apple", itemsJson[0].GetProperty("name").GetString());
        Assert.Equal(2m, itemsJson[0].GetProperty("quantity").GetDecimal());
        Assert.Equal("Orange", itemsJson[1].GetProperty("name").GetString());
        Assert.Equal(3m, itemsJson[1].GetProperty("quantity").GetDecimal());
        Assert.Equal(["Allie", "Ali"], rootJson.GetProperty("nickNames").EnumerateArray().Select(item => item.GetString()));
    }

    [Fact]
    public void RepositoryScalarArraySampleProducesStringArrayRoot()
    {
        string source = SampleAsset("scalar-array.xlsx");
        using (XLWorkbook workbook = new(source))
        {
            Assert.Equal(["setting", "root"], workbook.Worksheets.Select(sheet => sheet.Name));
            AssertSetting(workbook.Worksheet("setting"), "scalar-array");

            IXLWorksheet root = workbook.Worksheet("root");
            Assert.Equal("ID", root.Cell("A1").GetString());
            Assert.Equal("value", root.Cell("B1").GetString());
            Assert.False(root.Cell("B1").HasComment);
            Assert.Equal("N001", root.Cell("A2").GetString());
            Assert.Equal("Allie", root.Cell("B2").GetString());
            Assert.Equal("N002", root.Cell("A3").GetString());
            Assert.Equal("Ali", root.Cell("B3").GetString());
        }

        using TestWorkspace workspace = new();
        string input = workspace.File("scalar-array.xlsx");
        File.Copy(source, input);

        ConversionResult.Succeeded result = Assert.IsType<ConversionResult.Succeeded>(ExcelToJsonConverter.ConvertFile(input));
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(result.OutputPath));
        Assert.Equal(["Allie", "Ali"], json.RootElement.EnumerateArray().Select(item => item.GetString()));
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
            IXLWorksheet guide = workbook.Worksheet("_guide");
            Assert.False(guide.Cell("A1").IsEmpty());
            Assert.Equal("セル型上書き", guide.Cell("A6").GetString());
            Assert.Contains("そのセルだけ", guide.Cell("B6").GetString(), StringComparison.Ordinal);
            Assert.Contains("空または空白だけのメモは上書きしません", guide.Cell("B6").GetString(), StringComparison.Ordinal);
            Assert.Equal("scalar-array", guide.Cell("A11").GetString());
            Assert.Contains("IDとvalueの2列", guide.Cell("B11").GetString(), StringComparison.Ordinal);
            Assert.Contains("メモなしはtext", guide.Cell("B11").GetString(), StringComparison.Ordinal);
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

    private static void AssertLegacyCommentsOnly(string path)
    {
        using ZipArchive archive = ZipFile.OpenRead(path);
        Assert.Contains(archive.Entries, entry => entry.FullName.StartsWith("xl/comments", StringComparison.Ordinal));
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.StartsWith("xl/threadedComments", StringComparison.Ordinal));
    }
}
