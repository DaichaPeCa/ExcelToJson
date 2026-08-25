using ClosedXML.Excel;

namespace ExcelToJson.Tests;

internal sealed class TestWorkspace : IDisposable
{
    public TestWorkspace()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelToJson.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose()
    {
        const int maximumAttempts = 20;
        for (int attempt = 1; attempt <= maximumAttempts; attempt++)
        {
            try
            {
                if (Directory.Exists(Path))
                {
                    Directory.Delete(Path, recursive: true);
                }

                return;
            }
            catch (Exception exception) when (
                attempt < maximumAttempts &&
                exception is IOException or UnauthorizedAccessException)
            {
                Thread.Sleep(100);
            }
        }
    }
}

internal static class TestWorkbook
{
    public static XLWorkbook Create(string rootType = "object", string emptyCell = "omit")
    {
        XLWorkbook workbook = new();
        IXLWorksheet setting = workbook.AddWorksheet("setting");
        setting.Cell("A1").Value = "key";
        setting.Cell("B1").Value = "value";
        setting.Cell("A2").Value = "rootType";
        setting.Cell("B2").Value = rootType;
        setting.Cell("A3").Value = "emptyCell";
        setting.Cell("B3").Value = emptyCell;
        return workbook;
    }

    public static IXLWorksheet AddRoot(XLWorkbook workbook, params string[] properties)
    {
        IXLWorksheet root = workbook.AddWorksheet("root");
        root.Cell(1, 1).Value = "ID";
        for (int index = 0; index < properties.Length; index++)
        {
            root.Cell(1, index + 2).Value = properties[index];
        }

        return root;
    }

    public static void SetType(IXLWorksheet sheet, int column, string specification) =>
        sheet.Cell(1, column).CreateComment().AddText(specification);

    public static void SetType(IXLWorksheet sheet, int row, int column, string specification) =>
        sheet.Cell(row, column).CreateComment().AddText(specification);
}
