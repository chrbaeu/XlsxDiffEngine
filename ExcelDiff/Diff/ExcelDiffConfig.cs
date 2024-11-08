namespace ExcelDiffEngine;

public record class ExcelDiffConfig
{
    public bool IgnoreCase { get; init; } = true;

    public IReadOnlyCollection<string> KeyColumns { get; init; } = [];
    public IReadOnlyCollection<string> SecondaryKeyColumns { get; init; } = [];
    public IReadOnlyCollection<string> GroupKeyColumns { get; init; } = [];

    public IReadOnlyCollection<string>? ColumnsToCompare { get; init; }
    public IReadOnlyCollection<string>? ColumnsToIgnore { get; init; }
    public IReadOnlyCollection<string> ColumnsToOmit { get; init; } = [];
    public IReadOnlyCollection<string> ColumnsToTextCompareOnly { get; init; } = [];

    public IReadOnlyList<ModificationRule> ModificationRules { get; init; } = [];

    public IReadOnlyList<ValueChangedMarker> ValueChangedMarkers { get; init; } = [];

    public bool CopyCellFormat { get; init; } = true;
    public bool CopyCellStyle { get; init; }

    public bool ShowOldDataColumn { get; init; } = true;

    public bool AddOldValueAsComment { get; init; }
    public string? OldValueCommentPrefix { get; init; }

    public string? OldHeaderColumnComment { get; init; }
    public string? NewHeaderColumnComment { get; init; }
    public string? OldHeaderColumnPostfix { get; init; }
    public string? NewHeaderColumnPostfix { get; init; }

    public bool IgnoreUnchangedRows { get; init; }

    public SkipRowPredicate? SkipRowRule { get; init; }

    public CellStyle HeaderStyle { get; init; } = DefaultCellStyles.Header;
    public CellStyle RemovedRowStyle { get; init; } = DefaultCellStyles.RemovedRow;
    public CellStyle AddedRowStyle { get; init; } = DefaultCellStyles.AddedRow;
    public CellStyle ChangedCellStyle { get; init; } = DefaultCellStyles.ChangedCell;
    public CellStyle ChangedRowKeyColumnsStyle { get; init; } = DefaultCellStyles.ChangedRowKeyColumns;
}

public delegate bool SkipRowPredicate(IExcelDataSource excelDataSource, int row);

public enum DataKind { All, Old, New }

public record class ModificationRule(string Match, char Type, string Value, DataKind Target = DataKind.All);

public record class ValueChangedMarker(double MinDeviationInPercent, double MinDeviationAbsolute, CellStyle? CellStyle);
