namespace XlsxDiffEngine;

/// <summary>
/// Represents configuration options for comparing Excel sheets, including case sensitivity,
/// key columns, column inclusion/exclusion, styling, and value change markers.
/// </summary>
public record class ExcelDiffConfig
{
    /// <summary>
    /// Indicates whether comparisons should ignore case sensitivity. 
    /// Default is true.
    /// </summary>
    public bool IgnoreCase { get; init; } = true;

    /// <summary>
    /// The primary key columns for matching rows between sheets.
    /// </summary>
    public IReadOnlyCollection<string> KeyColumns { get; init; } = [];

    /// <summary>
    /// The secondary key columns used for additional matching criteria.
    /// </summary>
    public IReadOnlyCollection<string> SecondaryKeyColumns { get; init; } = [];

    /// <summary>
    /// The group key columns used to group rows when comparing sheets.
    /// </summary>
    public IReadOnlyCollection<string> GroupKeyColumns { get; init; } = [];

    /// <summary>
    /// The columns to compare explicitly. If null, all columns are compared by default.
    /// </summary>
    public IReadOnlyCollection<string>? ColumnsToCompare { get; init; }

    /// <summary>
    /// The columns to ignore during comparison.
    /// </summary>
    public IReadOnlyCollection<string>? ColumnsToIgnore { get; init; }

    /// <summary>
    /// The columns to omit from the output.
    /// </summary>
    public IReadOnlyCollection<string> ColumnsToOmit { get; init; } = [];

    /// <summary>
    /// The columns that should be compared as text only, ignoring data type differences.
    /// </summary>
    public IReadOnlyCollection<string> ColumnsToTextCompareOnly { get; init; } = [];

    /// <summary>
    /// A collection of modification rules that define how to handle changes based on specific criteria.
    /// </summary>
    public IReadOnlyList<ModificationRule> ModificationRules { get; init; } = [];

    /// <summary>
    /// A collection of value change markers that specify thresholds and styles for highlighting value changes.
    /// </summary>
    public IReadOnlyList<ValueChangedMarker> ValueChangedMarkers { get; init; } = [];

    /// <summary>
    /// Indicates whether cell formatting should be copied to the output.
    /// Default is true.
    /// </summary>
    public bool CopyCellFormat { get; init; } = true;

    /// <summary>
    /// Indicates whether cell styles (e.g., bold, italic) should be copied to the output.
    /// Default is false.
    /// </summary>
    public bool CopyCellStyle { get; init; }

    /// <summary>
    /// Indicates whether to include a column showing the old data in the comparison output.
    /// Default is true.
    /// </summary>
    public bool ShowOldDataColumn { get; init; } = true;

    /// <summary>
    /// Indicates whether to add the old values as comments to cells with the new value in the comparison output.
    /// </summary>
    public bool AddOldValueAsComment { get; init; }

    /// <summary>
    /// The prefix to be added to old value comments, if <see cref="AddOldValueAsComment"/> is enabled.
    /// </summary>
    public string? OldValueCommentPrefix { get; init; }

    /// <summary>
    /// The comment text for the headers of the old data columns in the comparison output, if specified.
    /// </summary>
    public string? OldHeaderColumnComment { get; init; }

    /// <summary>
    /// The comment text for the headers of the new data columns in the comparison output, if specified.
    /// </summary>
    public string? NewHeaderColumnComment { get; init; }

    /// <summary>
    /// A postfix for the headers of the old data columns in the comparison output, if specified.
    /// </summary>
    public string? OldHeaderColumnPostfix { get; init; }

    /// <summary>
    /// A postfix for the headers of the new data column in the comparison output, if specified.
    /// </summary>
    public string? NewHeaderColumnPostfix { get; init; }

    /// <summary>
    /// Indicates whether to skip unchanged rows in the comparison output.
    /// </summary>
    public bool SkipUnchangedRows { get; init; }

    /// <summary>
    /// Indicates whether to skip removed rows in the comparison output.
    /// </summary>
    public bool SkipRemovedRows { get; init; }

    /// <summary>
    /// Indicates whether to always set primary key column values in the comparison output.
    /// (Not supported in combination with ShowOldDataColumn = false)
    /// </summary>
    public bool AlwaysSetPrimaryKeyColumnValues { get; init; }

    /// <summary>
    /// A predicate to determine which rows should be skipped during comparison.
    /// </summary>
    public SkipRowPredicate? SkipRowRule { get; init; }

    /// <summary>
    /// The style for headers in the comparison output.
    /// </summary>
    public CellStyle HeaderStyle { get; init; } = DefaultCellStyles.Header;

    /// <summary>
    /// The style for rows that were removed in the comparison output.
    /// </summary>
    public CellStyle RemovedRowStyle { get; init; } = DefaultCellStyles.RemovedRow;

    /// <summary>
    /// The style for rows that were added in the comparison output.
    /// </summary>
    public CellStyle AddedRowStyle { get; init; } = DefaultCellStyles.AddedRow;

    /// <summary>
    /// The style for cells with changes in the comparison output.
    /// </summary>
    public CellStyle ChangedCellStyle { get; init; } = DefaultCellStyles.ChangedCell;

    /// <summary>
    /// The style for key columns in rows with changes in the comparison output.
    /// </summary>
    public CellStyle ChangedRowKeyColumnsStyle { get; init; } = DefaultCellStyles.ChangedRowKeyColumns;
}

/// <summary>
/// A delegate defining a predicate to determine if a row should be skipped in the comparison process.
/// </summary>
/// <param name="excelDataSource">The data source containing the row.</param>
/// <param name="row">The row index to check.</param>
/// <returns>True if the row should be skipped; otherwise, false.</returns>
public delegate bool SkipRowPredicate(IExcelDataSource excelDataSource, int row);

/// <summary>
/// Specifies the type of modification to apply to a cell.
/// </summary>
public enum ModificationKind
{
    /// <summary>
    /// Sets the number format of the cell.
    /// </summary>
    NumberFormat,

    /// <summary>
    /// Multiplies the cell value by a specified factor.
    /// </summary>
    Multiply,

    /// <summary>
    /// Sets the cell formula to a specified value.
    /// </summary>
    Formula,

    /// <summary>
    /// Replaces text in the cell using a regular expression pattern.
    /// </summary>
    RegexReplace
}

/// <summary>
/// Specifies the target data set for a modification rule, indicating if it applies to all data, only the old data, or only the new data.
/// </summary>
[Flags]
public enum DataKind
{
    /// <summary>
    /// The modification rule applies to the old data only.
    /// </summary>
    Old = 1,

    /// <summary>
    /// The modification rule applies to the new data only.
    /// </summary>
    New = 2,

    /// <summary>
    /// The modification rule applies to all data.
    /// </summary>
    All = Old | New,

    /// <summary>
    /// Specifices that the modification rule applies not to empty cells. Musst be combined with other flags.
    /// </summary>
    NonEmpty = 4,

    /// <summary>
    /// The modification rule applies to the old data only, but not to empty cells.
    /// </summary>
    OldNonEmpty = Old | NonEmpty,

    /// <summary>
    /// The modification rule applies to the new data only, but not to empty cells.
    /// </summary>
    NewNonEmpty = New | NonEmpty,

    /// <summary>
    /// The modification rule applies to all data, but not to empty cells.
    /// </summary>
    AllNonEmpty = Old | New | NonEmpty,
}

/// <summary>
/// Defines a modification rule for handling specific data changes.
/// </summary>
/// <param name="RegexPattern">The regex used to match the column names for which the modification is applied.</param>
/// <param name="ModificationKind">The type of modification to apply.</param>
/// <param name="Value">The value to use for the modification.</param>
/// <param name="Target">The target data for the rule (all, old, or new).</param>
/// <param name="AdditionalValue">An additional value used for the modification rule.</param>
public record class ModificationRule(string RegexPattern, ModificationKind ModificationKind, string Value, DataKind Target = DataKind.All, string? AdditionalValue = null);

/// <summary>
/// Defines a marker for highlighting cells where values have changed, based on a minimum deviation threshold.
/// </summary>
/// <param name="MinDeviationInPercent">The minimum percentage deviation required for the marker.</param>
/// <param name="MinDeviationAbsolute">The minimum absolute deviation required for the marker.</param>
/// <param name="CellStyle">The style to apply to cells where the change meets or exceeds the thresholds.</param>
public record class ValueChangedMarker(double MinDeviationInPercent, double MinDeviationAbsolute, CellStyle? CellStyle);
