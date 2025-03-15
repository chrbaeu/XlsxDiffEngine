using OfficeOpenXml;

namespace XlsxDiffEngine;

/// <summary>
/// A builder class for configuring and generating Excel comparison outputs, including settings for data sources,
/// columns, worksheet layout, and customization options for the comparison process.
/// </summary>
public class ExcelDiffBuilder
{
    private ExcelDiffConfig diffConfig = new();
    private XlsxDataProviderConfig xlsxConfig = new();
    private readonly List<XlsxFileInfo> oldFiles = [];
    private readonly List<XlsxFileInfo> newFiles = [];
    private bool hideOldColumns;
    private string[] columnsToHide = [];
    private string[] columnsToShow = [];
    private string[] header = [];
    private bool autoFitColumns = true;
    private bool autoFilter = true;
    private bool freezePanes = true;
    private readonly Dictionary<int, double> columnSizeDict = [];

    /// <summary>
    /// Adds files for comparison, allowing configuration of both "old" and "new" file sets.
    /// </summary>
    /// <param name="builderAction">An action to configure old and new files using <see cref="ExcelDiffXlsxFileConfigBuilder"/>.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddFiles(Action<ExcelDiffXlsxFileConfigBuilder> builderAction)
    {
        ArgumentNullThrowHelper.ThrowIfNull(builderAction);
        ExcelDiffXlsxFileConfigBuilder configBuilder = new();
        builderAction.Invoke(configBuilder);
        (XlsxFileInfo oldFile, XlsxFileInfo newFile) = configBuilder.Build();
        oldFiles.Add(oldFile);
        newFiles.Add(newFile);
        return this;
    }

    /// <summary>
    /// Adds files for comparison, specifying the "old" and "new" <see cref="XlsxFileInfo"/>.
    /// </summary>
    /// <param name="oldFile">The <see cref="XlsxFileInfo"/> for the old data.</param>
    /// <param name="newFile">The <see cref="XlsxFileInfo"/> for the old data.</param>
    /// <returns></returns>
    public ExcelDiffBuilder AddFiles(XlsxFileInfo oldFile, XlsxFileInfo newFile)
    {
        ArgumentNullThrowHelper.ThrowIfNull(oldFile);
        ArgumentNullThrowHelper.ThrowIfNull(newFile);
        oldFiles.Add(oldFile);
        newFiles.Add(newFile);
        return this;
    }

    /// <summary>
    /// Defines primary key columns used to match rows between compared files.
    /// </summary>
    /// <param name="keyColumns">Array of column names representing primary keys.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetKeyColumns(params string[] keyColumns)
    {
        diffConfig = diffConfig with { KeyColumns = keyColumns };
        return this;
    }

    /// <summary>
    /// Specifies secondary key columns for additional row-matching criteria.
    /// </summary>
    /// <param name="secondaryKeyColumns">Array of secondary key column names.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetSecondaryKeyColumns(params string[] secondaryKeyColumns)
    {
        diffConfig = diffConfig with { SecondaryKeyColumns = secondaryKeyColumns };
        return this;
    }

    /// <summary>
    /// Specifies grouping columns for organizing rows in the comparison output.
    /// </summary>
    /// <param name="groupKeyColumns">Array of column names to group rows by.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetGroupKeyColumns(params string[] groupKeyColumns)
    {
        diffConfig = diffConfig with { GroupKeyColumns = groupKeyColumns };
        return this;
    }

    /// <summary>
    /// Specifies the columns to include in the comparison.
    /// </summary>
    /// <param name="columnsToCompare">Array of column names to be compared.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnsToCompare(params string[] columnsToCompare)
    {
        diffConfig = diffConfig with { ColumnsToCompare = columnsToCompare };
        return this;
    }

    /// <summary>
    /// Specifies the columns to ignore during the comparison.
    /// </summary>
    /// <param name="columnsToIgnore">Array of column names to ignore.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnsToIgnore(params string[] columnsToIgnore)
    {
        diffConfig = diffConfig with { ColumnsToIgnore = columnsToIgnore };
        return this;
    }

    /// <summary>
    /// Specifies columns to omit from the output.
    /// </summary>
    /// <param name="columnsToOmit">Array of column names to omit.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnsToOmit(params string[] columnsToOmit)
    {
        diffConfig = diffConfig with { ColumnsToOmit = columnsToOmit };
        return this;
    }

    /// <summary>
    /// Specifies columns to be compared only as text, ignoring numeric data types.
    /// </summary>
    /// <param name="columnsToTextCompareOnly">Array of column names for text-only comparison.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnsToTextCompareOnly(params string[] columnsToTextCompareOnly)
    {
        diffConfig = diffConfig with { ColumnsToTextCompareOnly = columnsToTextCompareOnly };
        return this;
    }

    /// <summary>
    /// Specifies columns to sort by in the comparison output.
    /// </summary>
    /// <param name="columnsToSortBy">Array of column names to sort by.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnsToSortBy(params string[] columnsToSortBy)
    {
        diffConfig = diffConfig with { ColumnsToSortBy = columnsToSortBy };
        return this;
    }

    /// <summary>
    /// Specifies columns to fill with old values if no new value exists.
    /// </summary>
    /// <param name="columnsToFillWithOldValueIfNoNewValueExists">Array of column names to fill with old values if no new value exists.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnsToFillWithOldValueIfNoNewValueExists(params string[] columnsToFillWithOldValueIfNoNewValueExists)
    {
        diffConfig = diffConfig with { ColumnsToFillWithOldValueIfNoNewValueExists = columnsToFillWithOldValueIfNoNewValueExists };
        return this;
    }

    /// <summary>
    /// Sets the modification rules to apply to specific data changes.
    /// </summary>
    /// <param name="modificationRules">Array of modification rules.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetModificationRules(params ModificationRule[] modificationRules)
    {
        diffConfig = diffConfig with { ModificationRules = modificationRules };
        return this;
    }

    /// <summary>
    /// Adds additional modification rules to the existing set.
    /// </summary>
    /// <param name="modificationRules">Array of modification rules to add.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddModificationRules(params ModificationRule[] modificationRules)
    {
        diffConfig = diffConfig with { ModificationRules = [.. diffConfig.ModificationRules, .. modificationRules] };
        return this;
    }

    /// <summary>
    /// Adds a marker for highlighting value changes, specifying deviation thresholds and cell styling.
    /// </summary>
    /// <param name="minDeviationInPercent">Minimum percentage deviation to mark changes.</param>
    /// <param name="minDeviationAbsolute">Minimum absolute deviation to mark changes.</param>
    /// <param name="cellStyle">Optional cell style for marking changes.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddValueChangedMarker(double minDeviationInPercent, double minDeviationAbsolute, CellStyle? cellStyle)
    {
        diffConfig = diffConfig with
        {
            ValueChangedMarkers = [.. diffConfig.ValueChangedMarkers, new(minDeviationInPercent, minDeviationAbsolute, cellStyle)]
        };
        return this;
    }

    /// <summary>
    /// Specifies whether to always compare null values as text.
    /// </summary>
    /// <param name="alwaysCompareNullValuesAsText">Whether to always compare null values as text (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AlwaysCompareNullValuesAsText(bool alwaysCompareNullValuesAsText = true)
    {
        diffConfig = diffConfig with { AlwaysCompareNullValuesAsText = alwaysCompareNullValuesAsText };
        return this;
    }

    /// <summary>
    /// Specifies whether to add an empty row after groups in the comparison output.
    /// </summary>
    /// <param name="addEmptyRowAfterGroups">Whether to add an empty row after groups (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddEmptyRowAfterGroups(bool addEmptyRowAfterGroups = true)
    {
        diffConfig = diffConfig with { AddEmptyRowAfterGroups = addEmptyRowAfterGroups };
        return this;
    }

    /// <summary>
    /// Specifies whether to copy cell formatting from the original data.
    /// </summary>
    /// <param name="copyCellFormat">Whether to copy cell formats (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder CopyCellFormat(bool copyCellFormat = true)
    {
        diffConfig = diffConfig with { CopyCellFormat = copyCellFormat };
        return this;
    }

    /// <summary>
    /// Specifies whether to copy cell styling from the original data.
    /// </summary>
    /// <param name="copyCellStyle">Whether to copy cell styling (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder CopyCellStyle(bool copyCellStyle = true)
    {
        diffConfig = diffConfig with { CopyCellStyle = copyCellStyle };
        return this;
    }

    /// <summary>
    /// Specifies whether to show the old data column in the comparison output.
    /// </summary>
    /// <param name="showOldDataColumn">Whether to show the old data column (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder ShowOldDataColumn(bool showOldDataColumn = true)
    {
        diffConfig = diffConfig with { ShowOldDataColumn = showOldDataColumn };
        return this;
    }

    /// <summary>
    /// Adds the old value as a comment to the cell with the new data when differences are detected.
    /// </summary>
    /// <param name="prefix">Optional prefix for the old value comment.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddOldValueAsComment(string? prefix = null)
    {
        diffConfig = diffConfig with { AddOldValueAsComment = true, OldValueCommentPrefix = prefix };
        return this;
    }

    /// <summary>
    /// Sets a comment to be added to the header columns of the old data in the comparison output.
    /// </summary>
    /// <param name="oldHeaderColumnComment">Comment text for old data header columns.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetOldHeaderColumnComment(string oldHeaderColumnComment)
    {
        diffConfig = diffConfig with { OldHeaderColumnComment = oldHeaderColumnComment };
        return this;
    }

    /// <summary>
    /// Sets a comment to be added to the header columns of the new data in the comparison output.
    /// </summary>
    /// <param name="newHeaderColumnComment">Comment text for new data header columns.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetNewHeaderColumnComment(string newHeaderColumnComment)
    {
        diffConfig = diffConfig with { NewHeaderColumnComment = newHeaderColumnComment };
        return this;
    }

    /// <summary>
    /// Sets a postfix for the header columns of the old data in the comparison output.
    /// </summary>
    /// <param name="oldHeaderColumnPostfix">Postfix for old data header columns.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetOldHeaderColumnPostfix(string oldHeaderColumnPostfix)
    {
        diffConfig = diffConfig with { OldHeaderColumnPostfix = oldHeaderColumnPostfix };
        return this;
    }

    /// <summary>
    /// Sets a postfix for the header columns of the new data in the comparison output.
    /// </summary>
    /// <param name="newHeaderColumnPostfix">Postfix for new data header columns.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetNewHeaderColumnPostfix(string newHeaderColumnPostfix)
    {
        diffConfig = diffConfig with { NewHeaderColumnPostfix = newHeaderColumnPostfix };
        return this;
    }

    /// <summary>
    /// Configures whether to skip unchanged rows in the output.
    /// </summary>
    /// <param name="skipUnchangedRows">Whether to skip unchanged rows (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SkipUnchangedRows(bool skipUnchangedRows = true)
    {
        diffConfig = diffConfig with { SkipUnchangedRows = skipUnchangedRows };
        return this;
    }

    /// <summary>
    /// Configures whether to skip removed rows in the output.
    /// </summary>
    /// <param name="skipRemovedRows">Whether to skip removed rows (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SkipRemovedRows(bool skipRemovedRows = true)
    {
        diffConfig = diffConfig with { SkipRemovedRows = skipRemovedRows };
        return this;
    }

    /// <summary>
    /// Sets a custom rule for determining if a row should be skipped during the comparison.
    /// </summary>
    /// <param name="skipRowRule">The <see cref="SkipRowPredicate"/> to determine the rows to skip.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetSkipRowRule(SkipRowPredicate? skipRowRule)
    {
        diffConfig = diffConfig with { SkipRowRule = skipRowRule };
        return this;
    }

    /// <summary>
    /// Sets the style for headers in the comparison output.
    /// </summary>
    /// <param name="headerStyle">The <see cref="CellStyle"/> to apply to headers.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetHeaderStyle(CellStyle headerStyle)
    {
        diffConfig = diffConfig with { HeaderStyle = headerStyle };
        return this;
    }

    /// <summary>
    /// Sets the style for rows that were removed in the comparison output.
    /// </summary>
    /// <param name="removedRowStyle">The <see cref="CellStyle"/> to apply to removed rows.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetRemovedRowStyle(CellStyle removedRowStyle)
    {
        diffConfig = diffConfig with { RemovedRowStyle = removedRowStyle };
        return this;
    }

    /// <summary>
    /// Sets the style for rows that were added in the comparison output.
    /// </summary>
    /// <param name="addedRowStyle">The <see cref="CellStyle"/> to apply to added rows.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetAddedRowStyle(CellStyle addedRowStyle)
    {
        diffConfig = diffConfig with { AddedRowStyle = addedRowStyle };
        return this;
    }

    /// <summary>
    /// Sets the style for cells with changes in the comparison output.
    /// </summary>
    /// <param name="changedCellStyle">The <see cref="CellStyle"/> to apply to changed cells.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetChangedCellStyle(CellStyle changedCellStyle)
    {
        diffConfig = diffConfig with { ChangedCellStyle = changedCellStyle };
        return this;
    }

    /// <summary>
    /// Sets the style for key columns in rows with changes in the comparison output.
    /// </summary>
    /// <param name="changedRowKeyColumnsStyle">The <see cref="CellStyle"/> to apply to key columns in changed rows.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetChangedRowKeyColumnsStyle(CellStyle changedRowKeyColumnsStyle)
    {
        diffConfig = diffConfig with { ChangedRowKeyColumnsStyle = changedRowKeyColumnsStyle };
        return this;
    }

    /// <summary>
    /// Configures whether comparisons should ignore case sensitivity.
    /// </summary>
    /// <param name="ignoreCase">Whether to ignore case sensitivity (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder IgnoreCase(bool ignoreCase = true)
    {
        diffConfig = diffConfig with { IgnoreCase = ignoreCase };
        xlsxConfig = xlsxConfig with { IgnoreCase = ignoreCase };
        return this;
    }

    /// <summary>
    /// Specifies whether multiple worksheets should be merged into one for the comparison.
    /// </summary>
    /// <param name="mergeWorksheets">Whether to merge worksheets (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder MergeWorksheets(bool mergeWorksheets = true)
    {
        xlsxConfig = xlsxConfig with { MergeWorksheets = mergeWorksheets };
        return this;
    }

    /// <summary>
    /// Specifies whether multiple documents should be merged into one for the comparison.
    /// </summary>
    /// <param name="mergeDocuments">Whether to merge documents (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder MergeDocuments(bool mergeDocuments = true)
    {
        xlsxConfig = xlsxConfig with { MergeDocuments = mergeDocuments };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the row numbers.
    /// </summary>
    /// <param name="rowNumberColumnName">The name of the row number column.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddRowNumberAsColumn(string rowNumberColumnName = "RowNumber")
    {
        xlsxConfig = xlsxConfig with { RowNumberColumnName = rowNumberColumnName };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the worksheet names.
    /// </summary>
    /// <param name="worksheetColumnName">The name of the worksheet name column.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddWorksheetNameAsColumn(string worksheetColumnName = "WorksheetName")
    {
        xlsxConfig = xlsxConfig with { WorksheetNameColumnName = worksheetColumnName };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the merged worksheet names.
    /// </summary>
    /// <param name="mergedWorksheetColumnName">The name of the merged worksheet name column.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddMergedWorksheetNameAsColumn(string mergedWorksheetColumnName = "MergedWorksheetName")
    {
        xlsxConfig = xlsxConfig with { MergedWorksheetNameColumnName = mergedWorksheetColumnName };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the document names.
    /// </summary>
    /// <param name="documentNameColumnName">The name of the document name column.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder AddDocumentNameAsColumn(string documentNameColumnName = "DocumentName")
    {
        xlsxConfig = xlsxConfig with { DocumentNameColumnName = documentNameColumnName };
        return this;
    }

    /// <summary>
    /// Sets the name for the merged document.
    /// </summary>
    /// <param name="mergedDocumentName">The name of the merged document.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetMergedDocumentName(string mergedDocumentName)
    {
        xlsxConfig = xlsxConfig with { MergedDocumentName = mergedDocumentName };
        return this;
    }

    /// <summary>
    /// Configures whether to hide columns representing old data.
    /// </summary>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder HideOldColumns(bool hideOldColumns = true)
    {
        this.hideOldColumns = hideOldColumns;
        return this;
    }

    /// <summary>
    /// Specifies columns to hide in the output.
    /// </summary>
    /// <param name="columnsToHide">The names of the columns to hide.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder HideColumns(params string[] columnsToHide)
    {
        this.columnsToHide = columnsToHide;
        return this;
    }

    /// <summary>
    /// Specifies columns to display in the output.
    /// </summary>
    /// <param name="columnsToShow">The names of the columns to show.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder ShowColumns(params string[] columnsToShow)
    {
        this.columnsToShow = columnsToShow;
        return this;
    }

    /// <summary>
    /// Sets custom headers for the output worksheet.
    /// </summary>
    /// <param name="header">Header row strings.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetHeader(params string[] header)
    {
        this.header = header;
        return this;
    }

    /// <summary>
    /// Configures whether columns should automatically adjust to fit content.
    /// </summary>
    /// <param name="autoFitColumns">Whether to auto-fit columns (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetAutoFitColumns(bool autoFitColumns = true)
    {
        this.autoFitColumns = autoFitColumns;
        return this;
    }

    /// <summary>
    /// Configures whether an filter should be applied to the output worksheet.
    /// </summary>
    /// <param name="autoFilter">Whether to apply an auto-filter (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetAutoFilter(bool autoFilter = true)
    {
        this.autoFilter = autoFilter;
        return this;
    }

    /// <summary>
    /// Configures whether panes should be frozen in the output worksheet.
    /// </summary>
    /// <param name="freezePanes">Whether to freeze panes (default is true).</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetFreezePanes(bool freezePanes = true)
    {
        this.freezePanes = freezePanes;
        return this;
    }

    /// <summary>
    /// Sets a custom width for a specific column in the output worksheet.
    /// </summary>
    /// <param name="column">The column index (1-based) to set the width for.</param>
    /// <param name="size">The width size for the column.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnSize(int column, double size)
    {
        columnSizeDict[column] = size;
        return this;
    }

    /// <summary>
    /// Sets custom widths for multiple columns in the output worksheet.
    /// </summary>
    /// <param name="sizes">Array of column widths for the output worksheet.</param>
    /// <returns>The current builder instance for method chaining.</returns>
    public ExcelDiffBuilder SetColumnSizes(double[] sizes)
    {
        sizes ??= [];
        for (int i = 0; i < sizes.Length; i++)
        {
            columnSizeDict[i + 1] = sizes[i];
        }
        return this;
    }

    /// <summary>
    /// Builds the Excel comparison output, with an optional post-processing action.
    /// </summary>
    /// <param name="postProcessingAction">An optional action to perform additional processing on the <see cref="ExcelPackage"/>.</param>
    /// <returns>The generated <see cref="ExcelPackage"/> containing the comparison output.</returns>
    public ExcelPackage Build(Action<ExcelPackage>? postProcessingAction = null)
    {
        StringComparer stringComparer = diffConfig.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        using var oldDataProvider = new XlsxDataProvider(oldFiles, xlsxConfig);
        using var newDataProvider = new XlsxDataProvider(newFiles, xlsxConfig);
        IReadOnlyList<IExcelDataSource> oldDataSources = oldDataProvider.GetDataSources();
        if (oldDataSources.Select(x => x.Name).ToHashSet(stringComparer).Count != oldDataSources.Count)
        {
            throw new InvalidOperationException("The old excel files to compare must contain unique worksheet names!");
        }
        var oldDataSourcesDict = oldDataSources.ToDictionary(x => x.Name, stringComparer);
        IReadOnlyList<IExcelDataSource> newDataSources = newDataProvider.GetDataSources();
        if (newDataSources.Select(x => x.Name).ToHashSet(stringComparer).Count != newDataSources.Count)
        {
            throw new InvalidOperationException("The new excel files to compare must contain unique worksheet names!");
        }
        if (!newDataSources.Any(x => oldDataSourcesDict.ContainsKey(x.Name)))
        {
            throw new InvalidOperationException("The excel files to compare must contain worksheets with the same name!");
        }
        ExcelPackage? excelPackage = null;
        try
        {
            excelPackage = new();
            foreach (IExcelDataSource newDataSource in newDataSources)
            {
                if (oldDataSourcesDict.TryGetValue(newDataSource.Name, out IExcelDataSource? oldDataSource))
                {
                    var diffEngine = new ExcelDiffWriter(oldDataSource, newDataSource, diffConfig);
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(newDataSource.Name);
                    int row = 1;
                    int column = hideOldColumns ? 2 : 1;
                    foreach (string headerRow in header)
                    {
                        worksheet.Cells[row, column].Value = headerRow;
                        row++;
                    }
                    _ = diffEngine.WriteDiff(worksheet, row);
                    if (autoFitColumns) { worksheet.Cells.AutoFitColumns(); }
                    foreach (KeyValuePair<int, double> item in columnSizeDict)
                    {
                        worksheet.Column(item.Key).Width = item.Value;
                    }
                    if (autoFilter) { worksheet.Cells[row, column, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].AutoFilter = true; }
                    if (freezePanes) { worksheet.View.FreezePanes(row + 1, 1); }
                    if (hideOldColumns || columnsToHide.Length > 0)
                    {
                        for (column = 1; column <= worksheet.Dimension.End.Column; column++)
                        {
                            if (columnsToShow.Contains(worksheet.Cells[row, column].Text, stringComparer)) { continue; }
                            if (hideOldColumns && diffConfig.ShowOldDataColumn && column % 2 != 0) { worksheet.Column(column).Hidden = true; }
                            if (columnsToHide.Contains(worksheet.Cells[row, column].Text, stringComparer))
                            {
                                worksheet.Column(column).Hidden = true;
                            }
                        }
                    }
                }
            }
            postProcessingAction?.Invoke(excelPackage);
            ExcelPackage result = excelPackage;
            excelPackage = null;
            return result;
        }
        finally
        {
            excelPackage?.Dispose();
        }
    }

    /// <summary>
    /// Builds and saves the Excel comparison output to the specified file path, with an optional post-processing action.
    /// </summary>
    /// <param name="outputFilePath">The path where the output file should be saved.</param>
    /// <param name="postProcessingAction">An optional action to perform additional processing on the <see cref="ExcelPackage"/>.</param>
    public void Build(string outputFilePath, Action<ExcelPackage>? postProcessingAction = null)
    {
        using ExcelPackage excelPackage = Build(postProcessingAction);
        excelPackage.SaveAs(new FileInfo(outputFilePath));
    }

}
