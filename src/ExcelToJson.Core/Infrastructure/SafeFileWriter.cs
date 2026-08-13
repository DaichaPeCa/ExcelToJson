namespace ExcelToJson.Core.Infrastructure;

internal static class SafeFileWriter
{
    public static void Write(string outputPath, byte[] contents)
    {
        string directory = Path.GetDirectoryName(outputPath)
            ?? throw new IOException("出力先ディレクトリを特定できません。");
        string temporaryPath = Path.Combine(directory, $".{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.tmp");

        try
        {
            File.WriteAllBytes(temporaryPath, contents);
            if (File.Exists(outputPath))
            {
                File.Replace(temporaryPath, outputPath, null);
            }
            else
            {
                File.Move(temporaryPath, outputPath);
            }
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }
}
