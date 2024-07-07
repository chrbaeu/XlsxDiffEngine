using OfficeOpenXml;

namespace ExcelDiffEngine;

public class XlsxFileConfigurationBuilder
{
    private XlsxFileInfo oldFile = new("Unspecified.xlsx");
    private XlsxFileInfo newFile = new("Unspecified.xlsx");

    public XlsxFileConfigurationBuilder SetOldFile(string filePath)
    {
        oldFile = oldFile with { FileInfo = new(filePath) };
        return this;
    }

    public XlsxFileConfigurationBuilder SetNewFile(string filePath)
    {
        newFile = newFile with { FileInfo = new(filePath) };
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

    public ExcelDiffBuilder SetColumsToIgnore(params string[] columnsToIgnore)
    {
        diffConfig = diffConfig with { ColumnsToIgnore = columnsToIgnore };
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
                diffEngine.WriteDiff(worksheet);
                worksheet.Cells.AutoFitColumns();
                worksheet.Cells[worksheet.Dimension.Address].AutoFilter = true;
                worksheet.View.FreezePanes(2, 1);
            }
        }
        excelPackage.SaveAs(new FileInfo(outputFilePath));
    }

}
