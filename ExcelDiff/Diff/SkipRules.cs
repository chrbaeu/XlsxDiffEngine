namespace ExcelDiffEngine.Diff;

/// <summary>
/// Specifies predifiend rules for skipping rows during comparison.
/// </summary>
public static class SkipRules
{
    /// <summary>
    /// A predicate that skips rows with empty cells.
    /// </summary>
    public static SkipRowPredicate SkipEmptyRows => (IExcelDataSource excelDataSource, int row) => excelDataSource.GetRow(row).Values.All(value => value is null);
}
