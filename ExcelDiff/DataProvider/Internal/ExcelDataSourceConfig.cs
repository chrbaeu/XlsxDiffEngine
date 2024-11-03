namespace ExcelDiffEngine;

internal sealed record class ExcelDataSourceConfig
{
    public StringComparer StringComparer { get; init; } = StringComparer.OrdinalIgnoreCase;

    public string? RowNumberColumnName { get; init; }
    public string? WorksheetNameColumnName { get; init; }
    public string? MergedWorksheetNameColumnName { get; init; }

    public string? CustomColumnName { get; init; }
    public object? CustomColumnValue { get; init; }

    public IReadOnlyCollection<string> ColumnsToIgnore { get; init; } = [];
}