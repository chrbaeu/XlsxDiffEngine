using OfficeOpenXml;

namespace ExcelDiffEngine;

public sealed class XlsxDataProvider : IDisposable
{
    private readonly ICollection<XlsxFileInfo> xlsxFileInfos;
    private readonly Dictionary<XlsxFileInfo, ExcelPackage> excelPackagesDict = [];
    private readonly XlsxDataProviderConfig config;
    private readonly ExcelDataSourceConfig excelDataSourceConfig;
    private List<IExcelDataSource> dataSources = [];

    public XlsxDataProvider(string xlsxFile, int headerRow = 1, XlsxDataProviderConfig? config = null)
        : this([new(xlsxFile) { FromRow = headerRow }], config)
    { }

    public XlsxDataProvider(XlsxFileInfo xlsxFile, XlsxDataProviderConfig? config = null)
        : this([xlsxFile], config)
    { }

    public XlsxDataProvider(ICollection<XlsxFileInfo> xlsxFiles, XlsxDataProviderConfig? config = null)
    {
        xlsxFileInfos = xlsxFiles;
        excelPackagesDict = xlsxFiles.ToDictionary(x => x, x => new ExcelPackage(x.FileInfo));
        this.config = config ?? new XlsxDataProviderConfig();
        excelDataSourceConfig = new()
        {
            StringComparer = this.config.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal,
            RowNumberColumnName = this.config.RowNumberColumnName,
            WorksheetNameColumnName = this.config.WorksheetNameColumnName,
            MergedWorksheetNameColumnName = this.config.MergedWorksheetNameColumnName,
            ColumnsToIgnore = this.config.ColumnsToIgnore
        };
    }

    private List<IExcelDataSource> GetExcelDataSources(XlsxFileInfo xlsxFileInfo)
    {
        List<IExcelDataSource> dataSources = [];
        var excelPackage = excelPackagesDict[xlsxFileInfo];
        HashSet<string>? workSheetNames = config.WorksheetNames is not null ? new(config.WorksheetNames, excelDataSourceConfig.StringComparer) : null;
        var xlsxFileDataSourceConfig = excelDataSourceConfig;
        if (config.DocumentNameColumnName is not null)
        {
            xlsxFileDataSourceConfig = excelDataSourceConfig with
            {
                CustomColumnName = config.DocumentNameColumnName,
                CustomColumnValue = xlsxFileInfo.FileInfo.Name
            };
        }
        foreach (var excelWorksheet in excelPackage.Workbook.Worksheets)
        {
            if (workSheetNames is not null && !workSheetNames.Contains(excelWorksheet.Name)) { continue; }
            ExcelAddress excelAddress = excelWorksheet.Dimension;
            if (xlsxFileInfo.FromRow != 1 || xlsxFileInfo.FromColumn != 1 || xlsxFileInfo.ToRow != null || xlsxFileInfo.ToColumn != null)
            {
                excelAddress = new(xlsxFileInfo.FromRow, xlsxFileInfo.FromColumn,
                    xlsxFileInfo.ToRow ?? excelAddress.End.Row,
                    xlsxFileInfo.ToColumn ?? excelAddress.End.Column);
            }
            dataSources.Add(new ExcelDataSource(excelWorksheet, xlsxFileDataSourceConfig, excelAddress));
        }
        if (config.MergeWorkSheets)
        {
            var name = xlsxFileInfo.MergedWorksheetName ?? xlsxFileInfo.FileInfo.Name;
            return [new MergedExcelDataSource(name, dataSources, excelDataSourceConfig)];
        }
        return dataSources;
    }

    public List<IExcelDataSource> GetDataSources()
    {
        if (dataSources.Count > 0) { return dataSources; }
        foreach (var xlsxFileInfo in xlsxFileInfos)
        {
            dataSources.AddRange(GetExcelDataSources(xlsxFileInfo));
        }
        if (config.MergeDocuments)
        {
            dataSources = dataSources.GroupBy(x => x.Name)
                .Select(x => new MergedExcelDataSource(x.Key, x.ToList(), excelDataSourceConfig))
                .ToList<IExcelDataSource>();
            return [new MergedExcelDataSource(config.MergedDocumentName ?? "MergedDocument", dataSources,
                        excelDataSourceConfig with { MergedWorksheetNameColumnName = null })];
        }
        return dataSources;
    }

    public void Dispose()
    {
        foreach (var excelPackage in excelPackagesDict.Values)
        {
            excelPackage.Dispose();
        }
    }
}
