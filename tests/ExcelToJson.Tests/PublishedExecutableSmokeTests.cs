using ClosedXML.Excel;
using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace ExcelToJson.Tests;

public sealed class PublishedExecutableSmokeTests
{
    [Fact]
    public void PublishedExecutableRunsInIsolationWhenPathIsProvided()
    {
        string? publishedExecutable = Environment.GetEnvironmentVariable("EXCELTOJSON_PUBLISHED_EXE");
        if (string.IsNullOrWhiteSpace(publishedExecutable))
        {
            return;
        }

        Assert.True(File.Exists(publishedExecutable), $"Published executable not found: {publishedExecutable}");
        using TestWorkspace workspace = new();
        string applicationDirectory = workspace.File("application");
        Directory.CreateDirectory(applicationDirectory);
        string isolatedExecutable = System.IO.Path.Combine(applicationDirectory, "ExcelToJson.exe");
        File.Copy(publishedExecutable, isolatedExecutable);
        Assert.Equal(["ExcelToJson.exe"], Directory.EnumerateFiles(applicationDirectory).Select(System.IO.Path.GetFileName));

        string input = workspace.File("smoke.xlsx");
        using (XLWorkbook workbook = TestWorkbook.Create())
        {
            IXLWorksheet root = TestWorkbook.AddRoot(workbook, "message");
            root.Cell("A2").Value = "R1";
            root.Cell("B2").Value = "standalone";
            workbook.SaveAs(input);
        }

        ProcessStartInfo startInfo = new(isolatedExecutable)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        startInfo.ArgumentList.Add(input);
        startInfo.Environment["PATH"] = string.Empty;
        startInfo.Environment["DOTNET_ROOT"] = workspace.File("missing-dotnet");
        startInfo.Environment["DOTNET_ROOT_X64"] = workspace.File("missing-dotnet-x64");

        string standardOutput;
        string standardError;
        int exitCode;
        using (Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Published executable could not be started."))
        {
            standardOutput = process.StandardOutput.ReadToEnd();
            standardError = process.StandardError.ReadToEnd();
            Assert.True(process.WaitForExit(30_000), "Published executable did not exit within 30 seconds.");
            exitCode = process.ExitCode;
        }

        Assert.Equal(0, exitCode);
        Assert.Equal(string.Empty, standardError);
        Assert.Equal(System.IO.Path.ChangeExtension(input, ".json"), standardOutput.Trim());

        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(System.IO.Path.ChangeExtension(input, ".json")));
        Assert.Equal("standalone", json.RootElement.GetProperty("message").GetString());
        Assert.Equal(["ExcelToJson.exe"], Directory.EnumerateFiles(applicationDirectory).Select(System.IO.Path.GetFileName));
    }
}
