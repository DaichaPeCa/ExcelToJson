using System.Globalization;

namespace ExcelToJson.Core;

internal sealed record WorkbookModel(
    ConversionSettings Settings,
    IReadOnlyDictionary<string, SheetModel> Sheets,
    string RootSheetName);

internal sealed record ConversionSettings(
    RootType RootType,
    EmptyCellBehavior EmptyCell,
    string? DateInputFormat,
    string? DateOutputFormat,
    CultureInfo Culture);

internal enum RootType
{
    Object,
    Array,
    ScalarArray,
}

internal enum EmptyCellBehavior
{
    Omit,
    Null,
    EmptyString,
}

internal sealed record SheetModel(
    string Name,
    IReadOnlyList<ColumnModel> Columns,
    IReadOnlyList<RowModel> Rows)
{
    public IEnumerable<RowModel> FindRows(string id) => Rows.Where(row => string.Equals(row.Id, id, StringComparison.Ordinal));
}

internal sealed record ColumnModel(int Number, string Name, JsonType DefaultType, string Address);

internal abstract record JsonType
{
    private JsonType()
    {
    }

    internal sealed record Text : JsonType;

    internal sealed record Number : JsonType;

    internal sealed record Boolean : JsonType;

    internal sealed record Date : JsonType;

    internal sealed record ObjectReference(string SheetName) : JsonType;

    internal sealed record ArrayReference(string SheetName) : JsonType;

    internal sealed record ScalarArrayReference(string SheetName) : JsonType;
}

internal sealed record RowModel(int Number, string Id, IReadOnlyList<CellModel> Cells);

internal sealed record CellModel(string Address, CellValue Value, JsonType? TypeOverride);

internal abstract record CellValue
{
    private CellValue()
    {
    }

    internal sealed record Empty : CellValue;

    internal sealed record Text(string Raw, string Display) : CellValue;

    internal sealed record Number(double Raw, string Display) : CellValue;

    internal sealed record Boolean(bool Raw, string Display) : CellValue;

    internal sealed record DateTime(System.DateTime Raw, bool HasTimeComponent, string Display) : CellValue;

    internal sealed record TimeSpan(System.TimeSpan Raw, string Display) : CellValue;

    internal sealed record Error(string Display) : CellValue;
}

internal abstract record JsonValue
{
    private JsonValue()
    {
    }

    internal sealed record Object(IReadOnlyList<KeyValuePair<string, JsonValue>> Properties) : JsonValue;

    internal sealed record Array(IReadOnlyList<JsonValue> Items) : JsonValue;

    internal sealed record String(string Value) : JsonValue;

    internal sealed record Number(decimal Value) : JsonValue;

    internal sealed record Boolean(bool Value) : JsonValue;

    internal sealed record Null : JsonValue;
}
