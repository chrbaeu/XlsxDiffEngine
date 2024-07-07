namespace ExcelDiffEngine;

public record class XlsxDataProviderConfig
{
    public bool IgnoreCase { get; init; } = true;
    public string? RowNumberColumnName { get; init; }
    public string? WorksheetNameColumnName { get; init; }
    public string? MergedWorksheetNameColumnName { get; init; }
    public string? DocumentNameColumnName { get; init; }
    public bool MergeWorkSheets { get; init; }
    public bool MergeDocuments { get; init; }
    public string? MergedDocumentName { get; init; }
    public IReadOnlyCollection<string>? WorksheetNames { get; init; }
    public IReadOnlyCollection<string> ColumnsToIgnore { get; init; } = [];
}