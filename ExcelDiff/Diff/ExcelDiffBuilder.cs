using OfficeOpenXml;

namespace ExcelDiffEngine;

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

    public ExcelDiffBuilder AddFiles(Action<XlsxFileConfigurationBuilder> builderAction)
    {
        ArgumentNullException.ThrowIfNull(builderAction);
        XlsxFileConfigurationBuilder xlsxFileConfigurationBuilder = new();
        builderAction.Invoke(xlsxFileConfigurationBuilder);
        (XlsxFileInfo oldFile, XlsxFileInfo newFile) = xlsxFileConfigurationBuilder.Build();
        oldFiles.Add(oldFile);
        newFiles.Add(newFile);
        return this;
    }

    public ExcelDiffBuilder SetKeyColumns(params string[] keyColumns)
    {
        diffConfig = diffConfig with { KeyColumns = keyColumns };
        return this;
    }

    public ExcelDiffBuilder SetSecondaryKeyColumns(params string[] secondaryKeyColumns)
    {
        diffConfig = diffConfig with { SecondaryKeyColumns = secondaryKeyColumns };
        return this;
    }

    public ExcelDiffBuilder SetGroupKeyColumns(params string[] groupKeyColumns)
    {
        diffConfig = diffConfig with { GroupKeyColumns = groupKeyColumns };
        return this;
    }

    public ExcelDiffBuilder SetColumnsToCompare(params string[] columnsToCompare)
    {
        diffConfig = diffConfig with { ColumnsToCompare = columnsToCompare };
        return this;
    }
    public ExcelDiffBuilder SetColumnsToIgnore(params string[] columnsToIgnore)
    {
        diffConfig = diffConfig with { ColumnsToIgnore = columnsToIgnore };
        return this;
    }
    public ExcelDiffBuilder SetColumnsToOmit(params string[] columnsToOmit)
    {
        diffConfig = diffConfig with { ColumnsToOmit = columnsToOmit };
        return this;
    }

    public ExcelDiffBuilder SetColumnsToTextCompareOnly(params string[] columnsToTextCompareOnly)
    {
        diffConfig = diffConfig with { ColumnsToTextCompareOnly = columnsToTextCompareOnly };
        return this;
    }

    public ExcelDiffBuilder SetModificationRules(params ModificationRule[] modificationRules)
    {
        diffConfig = diffConfig with { ModificationRules = modificationRules };
        return this;
    }

    public ExcelDiffBuilder AddModificationRules(params ModificationRule[] modificationRules)
    {
        diffConfig = diffConfig with { ModificationRules = [.. diffConfig.ModificationRules, .. modificationRules] };
        return this;
    }

    public ExcelDiffBuilder AddValueChangedMarker(double minDeviationInPercent, double minDeviationAbsolute, CellStyle? cellStyle)
    {
        diffConfig = diffConfig with
        {
            ValueChangedMarkers = [.. diffConfig.ValueChangedMarkers, new(minDeviationInPercent, minDeviationAbsolute, cellStyle)]
        };
        return this;
    }

    public ExcelDiffBuilder CopyCellFormat(bool copyCellFormat = true)
    {
        diffConfig = diffConfig with { CopyCellFormat = copyCellFormat };
        return this;
    }

    public ExcelDiffBuilder CopyCellStyle(bool copyCellStyle = true)
    {
        diffConfig = diffConfig with { CopyCellStyle = copyCellStyle };
        return this;
    }

    public ExcelDiffBuilder AddOldValueAsComment(string? prefix = null)
    {
        diffConfig = diffConfig with { AddOldValueAsComment = true, OldValueCommentPrefix = prefix };
        return this;
    }

    public ExcelDiffBuilder SetOldHeaderColumnComment(string oldHeaderColumnComment)
    {
        diffConfig = diffConfig with { OldHeaderColumnComment = oldHeaderColumnComment };
        return this;
    }

    public ExcelDiffBuilder SetNewHeaderColumnComment(string newHeaderColumnComment)
    {
        diffConfig = diffConfig with { NewHeaderColumnComment = newHeaderColumnComment };
        return this;
    }

    public ExcelDiffBuilder SetOldHeaderColumnPostfix(string oldHeaderColumnPostfix)
    {
        diffConfig = diffConfig with { OldHeaderColumnPostfix = oldHeaderColumnPostfix };
        return this;
    }

    public ExcelDiffBuilder SetNewHeaderColumnPostfix(string newHeaderColumnPostfix)
    {
        diffConfig = diffConfig with { NewHeaderColumnPostfix = newHeaderColumnPostfix };
        return this;
    }

    public ExcelDiffBuilder IgnoreUnchangedRows()
    {
        diffConfig = diffConfig with { IgnoreUnchangedRows = true };
        return this;
    }

    public ExcelDiffBuilder IgnoreCase(bool ignoreCase = true)
    {
        xlsxConfig = xlsxConfig with { IgnoreCase = ignoreCase };
        return this;
    }

    public ExcelDiffBuilder MergeWorkSheets(bool mergeWorksheets = true)
    {
        xlsxConfig = xlsxConfig with { MergeWorksheets = mergeWorksheets };
        return this;
    }

    public ExcelDiffBuilder MergeDocuments(bool mergeDocuments = true)
    {
        xlsxConfig = xlsxConfig with { MergeDocuments = mergeDocuments };
        return this;
    }

    public ExcelDiffBuilder AddRowNumberAsColumn(string rowNumberColumnName = "RowNumber")
    {
        xlsxConfig = xlsxConfig with { RowNumberColumnName = rowNumberColumnName };
        return this;
    }

    public ExcelDiffBuilder AddWorksheetNameAsColumn(string worksheetColumnName = "WorksheetName")
    {
        xlsxConfig = xlsxConfig with { WorksheetNameColumnName = worksheetColumnName };
        return this;
    }

    public ExcelDiffBuilder AddMergedWorksheetNameAsColumn(string mergedWorksheetColumnName = "MergedWorksheetName")
    {
        xlsxConfig = xlsxConfig with { MergedWorksheetNameColumnName = mergedWorksheetColumnName };
        return this;
    }

    public ExcelDiffBuilder AddDocumentNameAsColumn(string documentNameColumnName = "DocumentName")
    {
        xlsxConfig = xlsxConfig with { DocumentNameColumnName = documentNameColumnName };
        return this;
    }

    public ExcelDiffBuilder SetMergedDocumentName(string mergedDocumentName)
    {
        xlsxConfig = xlsxConfig with { MergedDocumentName = mergedDocumentName };
        return this;
    }

    public ExcelDiffBuilder HideOldColumns()
    {
        hideOldColumns = true;
        return this;
    }

    public ExcelDiffBuilder HideColumns(params string[] columnsToHide)
    {
        this.columnsToHide = columnsToHide;
        return this;
    }

    public ExcelDiffBuilder ShowColumns(params string[] columnsToShow)
    {
        this.columnsToShow = columnsToShow;
        return this;
    }

    public ExcelDiffBuilder SetHeader(params string[] header)
    {
        this.header = header;
        return this;
    }

    public void Build(string outputFilePath)
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
                worksheet.Cells.AutoFitColumns();
                worksheet.Cells[row, column, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].AutoFilter = true;
                worksheet.View.FreezePanes(row + 1, 1);
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
        excelPackage.SaveAs(new FileInfo(outputFilePath));
    }

}
