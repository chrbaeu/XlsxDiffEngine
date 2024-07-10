using OfficeOpenXml;

namespace ExcelDiffEngine;

public class XlsxFileConfigurationBuilder
{
    private XlsxFileInfo oldFile = new("Unspecified.xlsx");
    private XlsxFileInfo newFile = new("Unspecified.xlsx");

    public XlsxFileConfigurationBuilder SetOldFile(string filePath, Action<ExcelPackage>? callback = null)
    {
        oldFile = oldFile with { FileInfo = new(filePath), Callback = callback };
        return this;
    }

    public XlsxFileConfigurationBuilder SetNewFile(string filePath, Action<ExcelPackage>? callback = null)
    {
        newFile = newFile with { FileInfo = new(filePath), Callback = callback };
        return this;
    }

    public XlsxFileConfigurationBuilder SetDocumentName(string documentName)
    {
        oldFile = oldFile with { DocumentName = documentName };
        newFile = newFile with { DocumentName = documentName };
        return this;
    }

    public XlsxFileConfigurationBuilder SetMergedWorksheetName(string mergedWorksheetName)
    {
        oldFile = oldFile with { MergedWorksheetName = mergedWorksheetName };
        newFile = newFile with { MergedWorksheetName = mergedWorksheetName };
        return this;
    }

    public XlsxFileConfigurationBuilder AddWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with
        {
            WorksheetInfos = [.. oldFile.WorksheetInfos, new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }]
        };
        newFile = newFile with
        {
            WorksheetInfos = [.. newFile.WorksheetInfos, new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }]
        };
        return this;
    }

    public XlsxFileConfigurationBuilder SetDataArea(int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with { FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn };
        newFile = newFile with { FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn };
        return this;
    }

    public XlsxFileConfigurationBuilder SetWorksheet(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with { WorksheetInfos = [new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }] };
        newFile = newFile with { WorksheetInfos = [new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }] };
        return this;
    }

    internal (XlsxFileInfo oldFile, XlsxFileInfo newFile) Build()
    {
        if (oldFile.MergedWorksheetName is null && newFile.MergedWorksheetName is null)
        {
            oldFile = oldFile with { MergedWorksheetName = Path.GetFileNameWithoutExtension(newFile.FileInfo.Name) };
            newFile = newFile with { MergedWorksheetName = Path.GetFileNameWithoutExtension(newFile.FileInfo.Name) };
        }
        return (oldFile, newFile);
    }
}

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
        XlsxFileConfigurationBuilder xlsxFileConfigurationBuilder = new();
        builderAction.Invoke(xlsxFileConfigurationBuilder);
        var (oldFile, newFile) = xlsxFileConfigurationBuilder.Build();
        oldFiles.Add(oldFile);
        newFiles.Add(newFile);
        return this;
    }

    public ExcelDiffBuilder SetKeyColumns(params string[] keyColumns)
    {
        diffConfig = diffConfig with { KeyColumns = keyColumns };
        return this;
    }

    public ExcelDiffBuilder SetGroupKeyColumns(params string[] groupKeyColumns)
    {
        diffConfig = diffConfig with { GroupKeyColumns = groupKeyColumns };
        return this;
    }

    public ExcelDiffBuilder SetColumnsToTextCompareOnly(params string[] columnsToTextCompareOnly)
    {
        diffConfig = diffConfig with { ColumnsToTextCompareOnly = columnsToTextCompareOnly };
        return this;
    }

    public ExcelDiffBuilder SetColumsToIgnore(params string[] columnsToIgnore)
    {
        diffConfig = diffConfig with { ColumnsToIgnore = columnsToIgnore };
        return this;
    }

    public ExcelDiffBuilder SetColumnsToOmit(params string[] columnsToOmit)
    {
        diffConfig = diffConfig with { ColumnsToOmit = columnsToOmit };
        return this;
    }

    public ExcelDiffBuilder AddModificationRules(params ModificationRule[] modificationRules)
    {
        diffConfig = diffConfig with { ModificationRules = modificationRules };
        return this;
    }

    public ExcelDiffBuilder MergeWorkSheets()
    {
        xlsxConfig = xlsxConfig with { MergeWorkSheets = true };
        return this;
    }

    public ExcelDiffBuilder MergeDocuments()
    {
        xlsxConfig = xlsxConfig with { MergeDocuments = true };
        return this;
    }

    public ExcelDiffBuilder SetMergedDocumentName(string mergedDocumentName)
    {
        xlsxConfig = xlsxConfig with { MergedDocumentName = mergedDocumentName };
        return this;
    }

    public ExcelDiffBuilder AddOldValueAsComment(string? prefix = null)
    {
        diffConfig = diffConfig with { AddOldValueAsComment = true, OldValueCommentPrefix = prefix };
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


    public ExcelDiffBuilder AddRowNumberAsColumn(string rowNumberColumnName = "RowNumber")
    {
        xlsxConfig = xlsxConfig with { RowNumberColumnName = rowNumberColumnName };
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

    public ExcelDiffBuilder SetOldColumnHeaderComment(string oldHeaderColumnComment)
    {
        diffConfig = diffConfig with { OldHeaderColumnComment = oldHeaderColumnComment };
        return this;
    }

    public ExcelDiffBuilder SetNewHeaderColumnComment(string newHeaderColumnComment)
    {
        diffConfig = diffConfig with { NewHeaderColumnComment = newHeaderColumnComment };
        return this;
    }

    public ExcelDiffBuilder IgnoreUnchangedRows()
    {
        diffConfig = diffConfig with { IgnoreUnchangedColumns = true };
        return this;
    }

    public void Build(string outputFilePath)
    {
        using var oldDataProvider = new XlsxDataProvider(oldFiles, xlsxConfig);
        using var newDataProvider = new XlsxDataProvider(newFiles, xlsxConfig);
        var oldDataSourcesDict = oldDataProvider.GetDataSources().ToDictionary(x => x.Name);
        using var excelPackage = new ExcelPackage();
        var newDataSources = newDataProvider.GetDataSources();
        foreach (var newDataSource in newDataSources)
        {
            if (oldDataSourcesDict.TryGetValue(newDataSource.Name, out var oldDataSource))
            {
                var diffEngine = new ExcelDiffWriter(oldDataSource, newDataSource, diffConfig);
                var worksheet = excelPackage.Workbook.Worksheets.Add(newDataSource.Name);
                var row = 1;
                var column = hideOldColumns ? 2 : 1;
                foreach (var headerRow in header)
                {
                    worksheet.Cells[row, column].Value = headerRow;
                    row++;
                }
                diffEngine.WriteDiff(worksheet, row);
                worksheet.Cells.AutoFitColumns();
                worksheet.Cells[row, column, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].AutoFilter = true;
                worksheet.View.FreezePanes(row + 1, 1);
                if (hideOldColumns || columnsToHide.Length > 0)
                {
                    var stringComparer = diffConfig.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
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
