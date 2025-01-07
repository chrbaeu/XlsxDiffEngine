namespace XlsxDiffEngine;

/// <summary>
/// Defines configuration settings for the <see cref="XlsxDataProvider"/>, enabling customization 
/// of data handling options such as case sensitivity, merging, and column configurations.
/// </summary>
public record class XlsxDataProviderConfig
{
    /// <summary>
    /// Specifies whether comparisons should be case-insensitive. 
    /// Default is true.
    /// </summary>
    public bool IgnoreCase { get; init; } = true;

    /// <summary>
    /// Indicates whether multiple worksheets should be merged into a single worksheet.
    /// </summary>
    public bool MergeWorksheets { get; init; }

    /// <summary>
    /// Indicates whether multiple documents should be merged into a single document.
    /// </summary>
    public bool MergeDocuments { get; init; }

    /// <summary>
    /// The name of an additional column used to store the row numbers.
    /// </summary>
    public string? RowNumberColumnName { get; init; }

    /// <summary>
    /// The name of an additional column used to store the worksheet names.
    /// </summary>
    public string? WorksheetNameColumnName { get; init; }

    /// <summary>
    /// The name of an additional column used to store the names of merged worksheets.
    /// </summary>
    public string? MergedWorksheetNameColumnName { get; init; }

    /// <summary>
    /// The name of an additional column used to store the document names.
    /// </summary>
    public string? DocumentNameColumnName { get; init; }

    /// <summary>
    /// The name assigned to the merged document when multiple documents are merged.
    /// </summary>
    public string? MergedDocumentName { get; init; }

    /// <summary>
    /// A collection of worksheet names to include in the dataset. If null, all worksheets are included by default.
    /// </summary>
    public IReadOnlyCollection<string>? WorksheetNames { get; init; }

    /// <summary>
    /// A collection of column names to exclude during data processing.
    /// </summary>
    public IReadOnlyCollection<string> ColumnsToIgnore { get; init; } = [];
}
