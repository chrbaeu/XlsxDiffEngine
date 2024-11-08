using OfficeOpenXml;

namespace ExcelDiffEngine;

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
        ArgumentNullException.ThrowIfNull(builderAction);
        ExcelDiffXlsxFileConfigBuilder configBuilder = new();
        builderAction.Invoke(configBuilder);
        (XlsxFileInfo oldFile, XlsxFileInfo newFile) = configBuilder.Build();
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
    public ExcelDiffBuilder SetNewHeaderColumnComment(string newHeaderColumnComment)
    {
        diffConfig = diffConfig with { NewHeaderColumnComment = newHeaderColumnComment };
        return this;
    }

    /// <summary>
    /// Sets a postfix for the header columns of the old data in the comparison output.
    /// </summary>
    public ExcelDiffBuilder SetOldHeaderColumnPostfix(string oldHeaderColumnPostfix)
    {
        diffConfig = diffConfig with { OldHeaderColumnPostfix = oldHeaderColumnPostfix };
        return this;
    }

    /// <summary>
    /// Sets a postfix for the header columns of the new data in the comparison output.
    /// </summary>
    public ExcelDiffBuilder SetNewHeaderColumnPostfix(string newHeaderColumnPostfix)
    {
        diffConfig = diffConfig with { NewHeaderColumnPostfix = newHeaderColumnPostfix };
        return this;
    }

    /// <summary>
    /// Configures whether to ignore unchanged rows in the output.
    /// </summary>
    public ExcelDiffBuilder IgnoreUnchangedRows(bool ignoreUnchangedRows)
    {
        diffConfig = diffConfig with { IgnoreUnchangedRows = ignoreUnchangedRows };
        return this;
    }

    /// <summary>
    /// Sets a custom rule for determining if a row should be skipped during the comparison.
    /// </summary>
    public ExcelDiffBuilder SetSkipRowRule(SkipRowPredicate? skipRowRule)
    {
        diffConfig = diffConfig with { SkipRowRule = skipRowRule };
        return this;
    }

    /// <summary>
    /// Configures whether comparisons should ignore case sensitivity.
    /// </summary>
    public ExcelDiffBuilder IgnoreCase(bool ignoreCase = true)
    {
        xlsxConfig = xlsxConfig with { IgnoreCase = ignoreCase };
        return this;
    }

    /// <summary>
    /// Specifies whether multiple worksheets should be merged into one for the comparison.
    /// </summary>
    public ExcelDiffBuilder MergeWorkSheets(bool mergeWorksheets = true)
    {
        xlsxConfig = xlsxConfig with { MergeWorksheets = mergeWorksheets };
        return this;
    }

    /// <summary>
    /// Specifies whether multiple documents should be merged into one for the comparison.
    /// </summary>
    public ExcelDiffBuilder MergeDocuments(bool mergeDocuments = true)
    {
        xlsxConfig = xlsxConfig with { MergeDocuments = mergeDocuments };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the row numbers.
    /// </summary>
    public ExcelDiffBuilder AddRowNumberAsColumn(string rowNumberColumnName = "RowNumber")
    {
        xlsxConfig = xlsxConfig with { RowNumberColumnName = rowNumberColumnName };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the worksheet names.
    /// </summary>
    public ExcelDiffBuilder AddWorksheetNameAsColumn(string worksheetColumnName = "WorksheetName")
    {
        xlsxConfig = xlsxConfig with { WorksheetNameColumnName = worksheetColumnName };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the merged worksheet names.
    /// </summary>
    public ExcelDiffBuilder AddMergedWorksheetNameAsColumn(string mergedWorksheetColumnName = "MergedWorksheetName")
    {
        xlsxConfig = xlsxConfig with { MergedWorksheetNameColumnName = mergedWorksheetColumnName };
        return this;
    }

    /// <summary>
    /// Adds a column to the output containing the document names.
    /// </summary>
    public ExcelDiffBuilder AddDocumentNameAsColumn(string documentNameColumnName = "DocumentName")
    {
        xlsxConfig = xlsxConfig with { DocumentNameColumnName = documentNameColumnName };
        return this;
    }

    /// <summary>
    /// Sets the name for the merged document.
    /// </summary>
    public ExcelDiffBuilder SetMergedDocumentName(string mergedDocumentName)
    {
        xlsxConfig = xlsxConfig with { MergedDocumentName = mergedDocumentName };
        return this;
    }

    /// <summary>
    /// Configures whether to hide columns representing old data.
    /// </summary>
    public ExcelDiffBuilder HideOldColumns()
    {
        hideOldColumns = true;
        return this;
    }

    /// <summary>
    /// Specifies columns to hide in the output.
    /// </summary>
    public ExcelDiffBuilder HideColumns(params string[] columnsToHide)
    {
        this.columnsToHide = columnsToHide;
        return this;
    }

    /// <summary>
    /// Specifies columns to display in the output.
    /// </summary>
    public ExcelDiffBuilder ShowColumns(params string[] columnsToShow)
    {
        this.columnsToShow = columnsToShow;
        return this;
    }

    /// <summary>
    /// Sets custom headers for the output worksheet.
    /// </summary>
    public ExcelDiffBuilder SetHeader(params string[] header)
    {
        this.header = header;
        return this;
    }

    /// <summary>
    /// Configures whether columns should automatically adjust to fit content.
    /// </summary>
    public ExcelDiffBuilder SetAutoFitColumns(bool autoFitColumns = true)
    {
        this.autoFitColumns = autoFitColumns;
        return this;
    }

    /// <summary>
    /// Configures whether an filter should be applied to the output worksheet.
    /// </summary>
    public ExcelDiffBuilder SetAutoFilter(bool autoFilter = true)
    {
        this.autoFilter = autoFilter;
        return this;
    }

    /// <summary>
    /// Configures whether panes should be frozen in the output worksheet.
    /// </summary>
    public ExcelDiffBuilder SetFreezePanes(bool freezePanes = true)
    {
        this.freezePanes = freezePanes;
        return this;
    }

    /// <summary>
    /// Sets a custom width for a specific column in the output worksheet.
    /// </summary>
    public ExcelDiffBuilder SetColumnSize(int column, double size)
    {
        columnSizeDict[column] = size;
        return this;
    }

    /// <summary>
    /// Sets custom widths for multiple columns in the output worksheet.
    /// </summary>
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
    /// Builds and saves the Excel comparison output to the specified file path, with an optional post-processing action.
    /// </summary>
    /// <param name="outputFilePath">The path where the output file should be saved.</param>
    /// <param name="postProcessingAction">An optional action to perform additional processing on the <see cref="ExcelPackage"/>.</param>
    public void Build(string outputFilePath, Action<ExcelPackage>? postProcessingAction = null)
    {
        using var oldDataProvider = new XlsxDataProvider(oldFiles, xlsxConfig);
        using var newDataProvider = new XlsxDataProvider(newFiles, xlsxConfig);
        var oldDataSourcesDict = oldDataProvider.GetDataSources().ToDictionary(x => x.Name);
        using var excelPackage = new ExcelPackage();
        IReadOnlyList<IExcelDataSource> newDataSources = newDataProvider.GetDataSources();
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
                    StringComparer stringComparer = diffConfig.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
                    for (column = 1; column <= worksheet.Dimension.End.Column; column++)
                    {
                        if (columnsToShow.Contains(worksheet.Cells[row, column].Text, stringComparer)) { continue; }
                        if (hideOldColumns && column % 2 != 0) { worksheet.Column(column).Hidden = true; }
                        if (columnsToHide.Contains(worksheet.Cells[row, column].Text, stringComparer))
                        {
                            worksheet.Column(column).Hidden = true;
                        }
                    }
                }
            }
        }
        postProcessingAction?.Invoke(excelPackage);
        excelPackage.SaveAs(new FileInfo(outputFilePath));
    }
}
