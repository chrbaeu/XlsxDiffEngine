namespace XlsxDiffEngine.Diff;

/// <summary>
/// Specifies predefined rules for skipping rows during comparison.
/// </summary>
public static class PredefinedSkipRules
{
    /// <summary>
    /// A predicate that skips rows with empty cells.
    /// </summary>
    public static SkipRowPredicate SkipEmptyRows => (IExcelDataSource excelDataSource, int row) => excelDataSource.GetRow(row).Values.All(value => value is null);
}
