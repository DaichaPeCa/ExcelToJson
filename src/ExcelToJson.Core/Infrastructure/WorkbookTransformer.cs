using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelToJson.Core.Infrastructure;

internal sealed record WorkbookTransformResult(JsonValue? Root, IReadOnlyList<ConversionDiagnostic> Diagnostics);

internal static partial class WorkbookTransformer
{
    private static readonly NumberStyles DecimalStyles = NumberStyles.Float | NumberStyles.AllowThousands;

    public static WorkbookTransformResult Transform(WorkbookModel workbook)
    {
        TransformContext context = new(workbook);
        SheetModel rootSheet = workbook.Sheets[workbook.RootSheetName];

        JsonValue? root = workbook.Settings.RootType switch
        {
            RootType.Object => context.TransformRow(rootSheet, rootSheet.Rows.Single(), []),
            RootType.Array => new JsonValue.Array(
                rootSheet.Rows
                    .Select(row => context.TransformRow(rootSheet, row, []))
                    .Where(value => value is not null)
                    .Cast<JsonValue>()
                    .ToList()),
            _ => null,
        };

        return context.Diagnostics.Count == 0
            ? new WorkbookTransformResult(root, context.Diagnostics)
            : new WorkbookTransformResult(null, context.Diagnostics);
    }

    private sealed class TransformContext(WorkbookModel workbook)
    {
        public List<ConversionDiagnostic> Diagnostics { get; } = [];

        public JsonValue.Object? TransformRow(
            SheetModel sheet,
            RowModel row,
            IReadOnlyList<RecordKey> path)
        {
            RecordKey current = new(sheet.Name, row.Number, row.Id);
            int cycleStart = IndexOf(path, current);
            if (cycleStart >= 0)
            {
                IEnumerable<RecordKey> cycle = path.Skip(cycleStart).Append(current);
                Diagnostics.Add(new ConversionDiagnostic(
                    $"循環参照を検出しました: {string.Join(" -> ", cycle.Select(FormatRecordKey))}",
                    sheet.Name,
                    $"A{row.Number}"));
                return null;
            }

            List<RecordKey> nextPath = [.. path, current];
            List<KeyValuePair<string, JsonValue>> properties = [];
            for (int index = 0; index < sheet.Columns.Count; index++)
            {
                ColumnModel column = sheet.Columns[index];
                CellModel cell = row.Cells[index];
                if (cell.Value is CellValue.Empty)
                {
                    AddEmptyValue(properties, column.Name, workbook.Settings.EmptyCell);
                    continue;
                }

                JsonValue? value = ConvertValue(sheet, row, column, cell, nextPath);
                if (value is not null)
                {
                    properties.Add(new KeyValuePair<string, JsonValue>(column.Name, value));
                }
            }

            return new JsonValue.Object(properties);
        }

        private JsonValue? ConvertValue(
            SheetModel sourceSheet,
            RowModel sourceRow,
            ColumnModel column,
            CellModel cell,
            IReadOnlyList<RecordKey> path) => (cell.TypeOverride ?? column.DefaultType) switch
            {
                JsonType.Text => ConvertText(sourceSheet, cell),
                JsonType.Number => ConvertNumber(sourceSheet, cell),
                JsonType.Boolean => ConvertBoolean(sourceSheet, cell),
                JsonType.Date => ConvertDate(sourceSheet, cell),
                JsonType.ObjectReference reference => ResolveObject(sourceSheet, sourceRow, cell, reference, path),
                JsonType.ArrayReference reference => ResolveArray(sourceSheet, sourceRow, cell, reference, path),
                _ => null,
            };

        private JsonValue.String? ConvertText(SheetModel sheet, CellModel cell)
        {
            if (cell.Value is CellValue.Error error)
            {
                AddCellError(sheet, cell, $"Excelエラー値 '{error.Display}' は変換できません。");
                return null;
            }

            return new JsonValue.String(GetText(cell.Value, preserveTextWhitespace: true));
        }

        private JsonValue.Number? ConvertNumber(SheetModel sheet, CellModel cell)
        {
            try
            {
                decimal number = cell.Value switch
                {
                    CellValue.Number numeric => Convert.ToDecimal(numeric.Raw),
                    CellValue.Text text when decimal.TryParse(text.Raw, DecimalStyles, workbook.Settings.Culture, out decimal parsed) => parsed,
                    CellValue.DateTime => throw new FormatException("Excelの日付・日時セルはnumberとして変換できません。"),
                    CellValue.Error error => throw new FormatException($"Excelエラー値 '{error.Display}' は変換できません。"),
                    _ => throw new FormatException("numberへ変換できる数値または文字列ではありません。"),
                };

                return new JsonValue.Number(number);
            }
            catch (Exception exception) when (exception is FormatException or OverflowException)
            {
                AddCellError(sheet, cell, $"numberへ変換できません。{exception.Message}");
                return null;
            }
        }

        private JsonValue.Boolean? ConvertBoolean(SheetModel sheet, CellModel cell)
        {
            bool? value = cell.Value switch
            {
                CellValue.Boolean boolean => boolean.Raw,
                CellValue.Number number when number.Raw == 0d => false,
                CellValue.Number number when number.Raw == 1d => true,
                CellValue.Text text => ParseBoolean(text.Raw),
                _ => null,
            };

            if (value is null)
            {
                AddCellError(sheet, cell, "booleanへ変換できません。許容値は true、false、1、0、yes、no です。");
                return null;
            }

            return new JsonValue.Boolean(value.Value);
        }

        private JsonValue.String? ConvertDate(SheetModel sheet, CellModel cell)
        {
            ParsedDate? parsed = cell.Value switch
            {
                CellValue.DateTime date => new ParsedDate(date.Raw, date.HasTimeComponent),
                CellValue.Text text => ParseDateText(text.Raw, workbook.Settings),
                _ => null,
            };

            if (parsed is null)
            {
                AddCellError(sheet, cell, "dateへ変換できません。日付成分を含む有効な日付文字列またはExcel日付セルを指定してください。");
                return null;
            }

            string format = workbook.Settings.DateOutputFormat
                ?? (parsed.HasTimeComponent ? "yyyy-MM-dd'T'HH:mm:ss" : "yyyy-MM-dd");
            return new JsonValue.String(parsed.Value.ToString(format, CultureInfo.InvariantCulture));
        }

        private JsonValue.Object? ResolveObject(
            SheetModel sourceSheet,
            RowModel sourceRow,
            CellModel cell,
            JsonType.ObjectReference reference,
            IReadOnlyList<RecordKey> path)
        {
            SheetModel target = workbook.Sheets[reference.SheetName];
            string id = GetText(cell.Value, preserveTextWhitespace: false).Trim();
            List<RowModel> matches = target.FindRows(id).ToList();
            if (matches.Count != 1)
            {
                AddCellError(
                    sourceSheet,
                    cell,
                    $"object参照 '{target.Name}:{id}' の一致件数は1件でなければなりません（実際: {matches.Count}件）。");
                return null;
            }

            return TransformRow(target, matches[0], path);
        }

        private JsonValue.Array? ResolveArray(
            SheetModel sourceSheet,
            RowModel sourceRow,
            CellModel cell,
            JsonType.ArrayReference reference,
            IReadOnlyList<RecordKey> path)
        {
            SheetModel target = workbook.Sheets[reference.SheetName];
            string id = GetText(cell.Value, preserveTextWhitespace: false).Trim();
            List<RowModel> matches = target.FindRows(id).ToList();
            if (matches.Count == 0)
            {
                AddCellError(sourceSheet, cell, $"array参照 '{target.Name}:{id}' に一致する行がありません。");
                return null;
            }

            List<JsonValue> items = [];
            foreach (RowModel match in matches)
            {
                JsonValue? item = TransformRow(target, match, path);
                if (item is not null)
                {
                    items.Add(item);
                }
            }

            return new JsonValue.Array(items);
        }

        private void AddCellError(SheetModel sheet, CellModel cell, string message) =>
            Diagnostics.Add(new ConversionDiagnostic(message, sheet.Name, cell.Address));
    }

    private static void AddEmptyValue(
        List<KeyValuePair<string, JsonValue>> properties,
        string propertyName,
        EmptyCellBehavior behavior)
    {
        if (behavior == EmptyCellBehavior.Null)
        {
            properties.Add(new KeyValuePair<string, JsonValue>(propertyName, new JsonValue.Null()));
        }
        else if (behavior == EmptyCellBehavior.EmptyString)
        {
            properties.Add(new KeyValuePair<string, JsonValue>(propertyName, new JsonValue.String(string.Empty)));
        }
    }

    private static string GetText(CellValue value, bool preserveTextWhitespace) => value switch
    {
        CellValue.Text text => preserveTextWhitespace ? text.Raw : text.Display,
        CellValue.Number number => number.Display,
        CellValue.Boolean boolean => boolean.Display,
        CellValue.DateTime dateTime => dateTime.Display,
        CellValue.TimeSpan timeSpan => timeSpan.Display,
        CellValue.Error error => error.Display,
        _ => string.Empty,
    };

    private static bool? ParseBoolean(string value)
    {
        if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "0", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "no", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return null;
    }

    private static ParsedDate? ParseDateText(string text, ConversionSettings settings)
    {
        string? format = settings.DateInputFormat;
        bool hasOffset = OffsetSuffixRegex().IsMatch(text)
            || (format is not null && OffsetFormatTokenRegex().IsMatch(RemoveQuotedParts(format)));
        bool hasDateComponent = format is null
            ? DateComponentRegex().IsMatch(text)
            : DateFormatTokenRegex().IsMatch(RemoveQuotedParts(format));
        if (!hasDateComponent)
        {
            return null;
        }

        bool hasTimeComponent = format is null
            ? TimeComponentRegex().IsMatch(text)
            : TimeFormatTokenRegex().IsMatch(RemoveQuotedParts(format));

        if (hasOffset)
        {
            bool parsed = format is null
                ? DateTimeOffset.TryParse(text, settings.Culture, DateTimeStyles.AllowWhiteSpaces, out DateTimeOffset offsetValue)
                : DateTimeOffset.TryParseExact(text, format, settings.Culture, DateTimeStyles.AllowWhiteSpaces, out offsetValue);
            return parsed ? new ParsedDate(offsetValue.LocalDateTime, true) : null;
        }

        bool parsedDate = format is null
            ? System.DateTime.TryParse(text, settings.Culture, DateTimeStyles.AllowWhiteSpaces, out System.DateTime value)
            : System.DateTime.TryParseExact(text, format, settings.Culture, DateTimeStyles.AllowWhiteSpaces, out value);
        return parsedDate ? new ParsedDate(value, hasTimeComponent) : null;
    }

    private static string RemoveQuotedParts(string format) => QuotedPartRegex().Replace(format, string.Empty);

    private static int IndexOf(IReadOnlyList<RecordKey> path, RecordKey key)
    {
        for (int index = 0; index < path.Count; index++)
        {
            if (path[index].Equals(key))
            {
                return index;
            }
        }

        return -1;
    }

    private static string FormatRecordKey(RecordKey key) => $"{key.SheetName}[{key.RowNumber}, ID={key.Id}]";

    private readonly record struct RecordKey(string SheetName, int RowNumber, string Id);

    private sealed record ParsedDate(System.DateTime Value, bool HasTimeComponent);

    [GeneratedRegex("(?i)(Z|[+-]\\d{2}:?\\d{2})\\s*$")]
    private static partial Regex OffsetSuffixRegex();

    [GeneratedRegex("(?<!%)K|z{1,3}")]
    private static partial Regex OffsetFormatTokenRegex();

    [GeneratedRegex("(?i)(\\d{1,4}\\s*[/.-]\\s*\\d{1,2}|\\d{1,4}年|\\d{1,2}月|\\d{1,2}日|[A-Za-z]{3,})")]
    private static partial Regex DateComponentRegex();

    [GeneratedRegex("(?<!%)[yMd]")]
    private static partial Regex DateFormatTokenRegex();

    [GeneratedRegex("(?i)(\\d{1,2}:\\d{2}|T\\d{1,2}|AM|PM|午前|午後)")]
    private static partial Regex TimeComponentRegex();

    [GeneratedRegex("(?<!%)[Hhstm]|:")]
    private static partial Regex TimeFormatTokenRegex();

    [GeneratedRegex("\"(?:[^\"]|\"\")*\"")]
    private static partial Regex QuotedPartRegex();
}
