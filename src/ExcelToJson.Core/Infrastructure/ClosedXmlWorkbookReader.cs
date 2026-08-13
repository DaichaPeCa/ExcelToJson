using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelToJson.Core.Infrastructure;

internal sealed record WorkbookReadResult(WorkbookModel? Workbook, IReadOnlyList<ConversionDiagnostic> Diagnostics);

internal static partial class ClosedXmlWorkbookReader
{
    private static readonly StringComparer ControlComparer = StringComparer.OrdinalIgnoreCase;

    public static WorkbookReadResult Read(string inputPath)
    {
        List<ConversionDiagnostic> diagnostics = [];

        if (!File.Exists(inputPath))
        {
            diagnostics.Add(new ConversionDiagnostic($"入力Excelファイルが存在しません: {inputPath}"));
            return new WorkbookReadResult(null, diagnostics);
        }

        if (!string.Equals(Path.GetExtension(inputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            diagnostics.Add(new ConversionDiagnostic("入力ファイルの拡張子は .xlsx でなければなりません。"));
            return new WorkbookReadResult(null, diagnostics);
        }

        try
        {
            using XLWorkbook workbook = new(inputPath);
            return ReadWorkbook(workbook, CultureInfo.CurrentCulture);
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            diagnostics.Add(new ConversionDiagnostic($"入力Excelファイルを読み取れませんでした。{exception.Message}"));
            return new WorkbookReadResult(null, diagnostics);
        }
    }

    internal static WorkbookReadResult ReadWorkbook(XLWorkbook workbook, CultureInfo culture)
    {
        List<ConversionDiagnostic> diagnostics = [];
        List<IXLWorksheet> settingSheets = workbook.Worksheets
            .Where(sheet => ControlComparer.Equals(sheet.Name, "setting"))
            .ToList();
        List<IXLWorksheet> rootSheets = workbook.Worksheets
            .Where(sheet => ControlComparer.Equals(sheet.Name, "root"))
            .ToList();

        if (settingSheets.Count != 1)
        {
            diagnostics.Add(new ConversionDiagnostic(
                settingSheets.Count == 0
                    ? "settingシートが存在しません。"
                    : "settingシートが複数存在します。"));
        }

        if (rootSheets.Count != 1)
        {
            diagnostics.Add(new ConversionDiagnostic(
                rootSheets.Count == 0
                    ? "rootシートが存在しません。"
                    : "rootシートが複数存在します。"));
        }

        ConversionSettings? settings = settingSheets.Count == 1
            ? ReadSettings(settingSheets[0], culture, diagnostics)
            : null;

        Dictionary<string, SheetModel> sheets = new(ControlComparer);
        foreach (IXLWorksheet worksheet in workbook.Worksheets)
        {
            if (ControlComparer.Equals(worksheet.Name, "setting") || worksheet.Name.StartsWith('_'))
            {
                continue;
            }

            SheetModel? sheet = ReadDataSheet(worksheet, culture, diagnostics);
            if (sheet is not null)
            {
                sheets.Add(sheet.Name, sheet);
            }
        }

        ValidateReferences(sheets, diagnostics);

        if (settings is not null && rootSheets.Count == 1 && sheets.TryGetValue(rootSheets[0].Name, out SheetModel? root))
        {
            if (settings.RootType == RootType.Object && root.Rows.Count != 1)
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "rootType=objectの場合、rootシートのデータ行はちょうど1行でなければなりません。",
                    root.Name));
            }
        }

        if (diagnostics.Count != 0 || settings is null || rootSheets.Count != 1)
        {
            return new WorkbookReadResult(null, diagnostics);
        }

        return new WorkbookReadResult(
            new WorkbookModel(settings, sheets, rootSheets[0].Name),
            diagnostics);
    }

    private static ConversionSettings? ReadSettings(
        IXLWorksheet worksheet,
        CultureInfo culture,
        List<ConversionDiagnostic> diagnostics)
    {
        string keyHeader = ReadControlText(worksheet.Cell(1, 1), culture).Trim();
        string valueHeader = ReadControlText(worksheet.Cell(1, 2), culture).Trim();
        if (!ControlComparer.Equals(keyHeader, "key") || !ControlComparer.Equals(valueHeader, "value"))
        {
            diagnostics.Add(new ConversionDiagnostic(
                "1行目は A1=key、B1=value でなければなりません。",
                worksheet.Name,
                "A1:B1"));
        }

        Dictionary<string, (string Value, string Cell)> values = new(ControlComparer);
        int lastRow = Math.Max(1, worksheet.LastRowUsed(XLCellsUsedOptions.AllContents)?.RowNumber() ?? 1);
        for (int rowNumber = 2; rowNumber <= lastRow; rowNumber++)
        {
            string key = ReadControlText(worksheet.Cell(rowNumber, 1), culture).Trim();
            string value = ReadControlText(worksheet.Cell(rowNumber, 2), culture).Trim();
            if (key.Length == 0 && value.Length == 0)
            {
                continue;
            }

            if (key.Length == 0)
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "設定値に対応するキーがありません。",
                    worksheet.Name,
                    $"B{rowNumber}"));
                continue;
            }

            if (!IsKnownSettingKey(key))
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "未知のsettingキーです。",
                    worksheet.Name,
                    $"A{rowNumber}",
                    key));
                continue;
            }

            if (!values.TryAdd(key, (value, $"B{rowNumber}")))
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "settingキーが重複しています。",
                    worksheet.Name,
                    $"A{rowNumber}",
                    key));
            }
        }

        RootType? rootType = null;
        if (!values.TryGetValue("rootType", out (string Value, string Cell) rootSetting) || rootSetting.Value.Length == 0)
        {
            diagnostics.Add(new ConversionDiagnostic("rootTypeが指定されていません。", worksheet.Name, SettingKey: "rootType"));
        }
        else if (ControlComparer.Equals(rootSetting.Value, "object"))
        {
            rootType = RootType.Object;
        }
        else if (ControlComparer.Equals(rootSetting.Value, "array"))
        {
            rootType = RootType.Array;
        }
        else
        {
            diagnostics.Add(new ConversionDiagnostic(
                "rootTypeは object または array でなければなりません。",
                worksheet.Name,
                rootSetting.Cell,
                "rootType"));
        }

        EmptyCellBehavior? emptyCell = EmptyCellBehavior.Omit;
        if (values.TryGetValue("emptyCell", out (string Value, string Cell) emptySetting) && emptySetting.Value.Length != 0)
        {
            emptyCell = emptySetting.Value.ToUpperInvariant() switch
            {
                "OMIT" => EmptyCellBehavior.Omit,
                "NULL" => EmptyCellBehavior.Null,
                "EMPTYSTRING" => EmptyCellBehavior.EmptyString,
                _ => null,
            };

            if (emptyCell is null)
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "emptyCellは omit、null、emptyString のいずれかでなければなりません。",
                    worksheet.Name,
                    emptySetting.Cell,
                    "emptyCell"));
            }
        }

        string? dateInputFormat = GetOptionalSetting(values, "dateInputFormat");
        string? dateOutputFormat = GetOptionalSetting(values, "dateOutputFormat");
        ValidateDateFormat(dateInputFormat, culture, "dateInputFormat", worksheet.Name, values, diagnostics);
        ValidateDateFormat(dateOutputFormat, CultureInfo.InvariantCulture, "dateOutputFormat", worksheet.Name, values, diagnostics);

        return rootType is not null && emptyCell is not null
            ? new ConversionSettings(rootType.Value, emptyCell.Value, dateInputFormat, dateOutputFormat, culture)
            : null;
    }

    private static SheetModel? ReadDataSheet(
        IXLWorksheet worksheet,
        CultureInfo culture,
        List<ConversionDiagnostic> diagnostics)
    {
        string idHeader = ReadControlText(worksheet.Cell(1, 1), culture).Trim();
        if (!ControlComparer.Equals(idHeader, "ID"))
        {
            diagnostics.Add(new ConversionDiagnostic(
                "A1はID列でなければなりません。",
                worksheet.Name,
                "A1"));
        }

        int lastHeaderColumn = worksheet.Row(1)
            .CellsUsed(XLCellsUsedOptions.Contents)
            .Select(cell => cell.Address.ColumnNumber)
            .DefaultIfEmpty(1)
            .Max();

        List<ColumnModel> columns = [];
        HashSet<string> propertyNames = new(StringComparer.Ordinal);
        for (int columnNumber = 2; columnNumber <= lastHeaderColumn; columnNumber++)
        {
            IXLCell headerCell = worksheet.Cell(1, columnNumber);
            string propertyName = ReadControlText(headerCell, culture).Trim();
            if (propertyName.Length == 0)
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "表の途中に空のJSONプロパティ名があります。",
                    worksheet.Name,
                    headerCell.Address.ToString()));
                continue;
            }

            if (!propertyNames.Add(propertyName))
            {
                diagnostics.Add(new ConversionDiagnostic(
                    $"JSONプロパティ名 '{propertyName}' が重複しています。",
                    worksheet.Name,
                    headerCell.Address.ToString()));
            }

            ColumnType? type = ParseColumnType(headerCell, worksheet.Name, diagnostics);
            if (type is not null)
            {
                columns.Add(new ColumnModel(columnNumber, propertyName, type, headerCell.Address.ToString()!));
            }
        }

        int lastContentColumn = Math.Max(
            lastHeaderColumn,
            worksheet.LastColumnUsed(XLCellsUsedOptions.AllContents)?.ColumnNumber() ?? lastHeaderColumn);
        int lastContentRow = Math.Max(
            2,
            worksheet.LastRowUsed(XLCellsUsedOptions.AllContents)?.RowNumber() ?? 2);
        List<RowModel> rows = [];

        for (int rowNumber = 2; rowNumber <= lastContentRow + 1; rowNumber++)
        {
            List<CellValue> rowValues = new(lastContentColumn);
            bool formulaReadFailed = false;
            for (int columnNumber = 1; columnNumber <= lastContentColumn; columnNumber++)
            {
                CellValue value = ReadCellValue(
                    worksheet.Cell(rowNumber, columnNumber),
                    culture,
                    worksheet.Name,
                    diagnostics,
                    out bool failed);
                formulaReadFailed |= failed;
                rowValues.Add(value);
            }

            if (formulaReadFailed)
            {
                continue;
            }

            bool isEmptyRow = rowValues.All(IsEmpty);
            if (isEmptyRow)
            {
                break;
            }

            for (int columnNumber = lastHeaderColumn + 1; columnNumber <= lastContentColumn; columnNumber++)
            {
                if (!IsEmpty(rowValues[columnNumber - 1]))
                {
                    diagnostics.Add(new ConversionDiagnostic(
                        "ヘッダーで定義されていない列にデータがあります。",
                        worksheet.Name,
                        worksheet.Cell(rowNumber, columnNumber).Address.ToString()));
                }
            }

            string id = ToDisplayText(rowValues[0]).Trim();
            if (id.Length == 0)
            {
                diagnostics.Add(new ConversionDiagnostic(
                    "データ行のIDは空にできません。",
                    worksheet.Name,
                    $"A{rowNumber}"));
            }

            List<CellModel> cells = [];
            foreach (ColumnModel column in columns)
            {
                cells.Add(new CellModel(
                    worksheet.Cell(rowNumber, column.Number).Address.ToString()!,
                    rowValues[column.Number - 1]));
            }

            rows.Add(new RowModel(rowNumber, id, cells));
        }

        return new SheetModel(worksheet.Name, columns, rows);
    }

    private static ColumnType? ParseColumnType(
        IXLCell headerCell,
        string sheetName,
        List<ConversionDiagnostic> diagnostics)
    {
        string specification = headerCell.HasComment ? headerCell.GetComment().Text.Trim() : string.Empty;
        if (specification.Length == 0 || ControlComparer.Equals(specification, "text"))
        {
            return new ColumnType.Text();
        }

        if (ControlComparer.Equals(specification, "number"))
        {
            return new ColumnType.Number();
        }

        if (ControlComparer.Equals(specification, "boolean"))
        {
            return new ColumnType.Boolean();
        }

        if (ControlComparer.Equals(specification, "date"))
        {
            return new ColumnType.Date();
        }

        int separator = specification.IndexOf(':');
        if (separator >= 0)
        {
            string kind = specification[..separator].Trim();
            string referencedSheet = specification[(separator + 1)..].Trim();
            if (ControlComparer.Equals(kind, "object") || ControlComparer.Equals(kind, "array"))
            {
                if (referencedSheet.Length == 0)
                {
                    diagnostics.Add(new ConversionDiagnostic(
                        $"{kind}: の参照先シート名が空です。",
                        sheetName,
                        headerCell.Address.ToString()));
                    return null;
                }

                return ControlComparer.Equals(kind, "object")
                    ? new ColumnType.ObjectReference(referencedSheet)
                    : new ColumnType.ArrayReference(referencedSheet);
            }
        }

        diagnostics.Add(new ConversionDiagnostic(
            $"未知のJSON型 '{specification}' が指定されています。",
            sheetName,
            headerCell.Address.ToString()));
        return null;
    }

    private static CellValue ReadCellValue(
        IXLCell cell,
        CultureInfo culture,
        string sheetName,
        List<ConversionDiagnostic> diagnostics,
        out bool formulaReadFailed)
    {
        formulaReadFailed = false;
        XLCellValue value;
        try
        {
            value = cell.Value;
        }
        catch (Exception recalculationException) when (cell.HasFormula && recalculationException is not OutOfMemoryException)
        {
            value = cell.CachedValue;
            if (value.IsBlank)
            {
                diagnostics.Add(new ConversionDiagnostic(
                    $"数式の再計算結果も保存済み結果も取得できませんでした。{recalculationException.Message}",
                    sheetName,
                    cell.Address.ToString()));
                formulaReadFailed = true;
                return new CellValue.Empty();
            }
        }

        string display;
        try
        {
            display = cell.GetFormattedString(culture);
        }
        catch (Exception) when (cell.HasFormula)
        {
            display = value.ToString(culture);
        }

        return value.Type switch
        {
            XLDataType.Blank => new CellValue.Empty(),
            XLDataType.Text when value.GetText().Length == 0 => new CellValue.Empty(),
            XLDataType.Text => new CellValue.Text(value.GetText(), display),
            XLDataType.Number => new CellValue.Number(value.GetNumber(), display),
            XLDataType.Boolean => new CellValue.Boolean(value.GetBoolean(), display),
            XLDataType.DateTime => new CellValue.DateTime(
                value.GetDateTime(),
                HasTimeFormat(cell.Style.NumberFormat.Format),
                display),
            XLDataType.TimeSpan => new CellValue.TimeSpan(value.GetTimeSpan(), display),
            XLDataType.Error => new CellValue.Error(
                cell.HasFormula
                    ? $"数式結果 {value.GetError()}"
                    : value.GetError().ToString()),
            _ => new CellValue.Error(display),
        };
    }

    private static string ReadControlText(IXLCell cell, CultureInfo culture)
    {
        CellValue value = ReadCellValue(cell, culture, cell.Worksheet.Name, [], out _);
        return ToDisplayText(value);
    }

    private static string ToDisplayText(CellValue value) => value switch
    {
        CellValue.Empty => string.Empty,
        CellValue.Text text => text.Raw,
        CellValue.Number number => number.Display,
        CellValue.Boolean boolean => boolean.Display,
        CellValue.DateTime dateTime => dateTime.Display,
        CellValue.TimeSpan timeSpan => timeSpan.Display,
        CellValue.Error error => error.Display,
        _ => string.Empty,
    };

    private static bool IsEmpty(CellValue value) => value is CellValue.Empty;

    private static void ValidateReferences(
        IReadOnlyDictionary<string, SheetModel> sheets,
        List<ConversionDiagnostic> diagnostics)
    {
        foreach (SheetModel sheet in sheets.Values)
        {
            foreach (ColumnModel column in sheet.Columns)
            {
                string? target = column.Type switch
                {
                    ColumnType.ObjectReference reference => reference.SheetName,
                    ColumnType.ArrayReference reference => reference.SheetName,
                    _ => null,
                };

                if (target is null)
                {
                    continue;
                }

                if (target.StartsWith('_'))
                {
                    diagnostics.Add(new ConversionDiagnostic(
                        $"変換対象外シート '{target}' は参照できません。",
                        sheet.Name,
                        column.Address));
                }
                else if (!sheets.ContainsKey(target))
                {
                    diagnostics.Add(new ConversionDiagnostic(
                        $"参照先シート '{target}' が存在しません。",
                        sheet.Name,
                        column.Address));
                }
            }
        }
    }

    private static bool IsKnownSettingKey(string key) =>
        ControlComparer.Equals(key, "rootType")
        || ControlComparer.Equals(key, "emptyCell")
        || ControlComparer.Equals(key, "dateInputFormat")
        || ControlComparer.Equals(key, "dateOutputFormat");

    private static string? GetOptionalSetting(
        Dictionary<string, (string Value, string Cell)> values,
        string key) =>
        values.TryGetValue(key, out (string Value, string Cell) setting) && setting.Value.Length != 0
            ? setting.Value
            : null;

    private static void ValidateDateFormat(
        string? format,
        CultureInfo culture,
        string key,
        string sheetName,
        Dictionary<string, (string Value, string Cell)> values,
        List<ConversionDiagnostic> diagnostics)
    {
        if (format is null)
        {
            return;
        }

        try
        {
            _ = new DateTime(2026, 8, 13, 14, 30, 25).ToString(format, culture);
        }
        catch (FormatException exception)
        {
            diagnostics.Add(new ConversionDiagnostic(
                $"日付形式が不正です。{exception.Message}",
                sheetName,
                values[key].Cell,
                key));
        }
    }

    private static bool HasTimeFormat(string format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            return false;
        }

        string normalized = QuotedFormatPartRegex().Replace(format, string.Empty);
        normalized = BracketedFormatPartRegex().Replace(normalized, string.Empty);
        return TimeFormatTokenRegex().IsMatch(normalized);
    }

    [GeneratedRegex("\"(?:[^\"]|\"\")*\"")]
    private static partial Regex QuotedFormatPartRegex();

    [GeneratedRegex("\\[[^\\]]*\\]")]
    private static partial Regex BracketedFormatPartRegex();

    [GeneratedRegex("(?i)(h+|s+|AM/PM|A/P)")]
    private static partial Regex TimeFormatTokenRegex();
}
