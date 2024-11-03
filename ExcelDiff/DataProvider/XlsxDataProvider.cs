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

    public IReadOnlyList<IExcelDataSource> GetDataSources()
    {
        if (dataSources.Count > 0) { return dataSources; }
        foreach (XlsxFileInfo xlsxFileInfo in xlsxFileInfos)
        {
            dataSources.AddRange(GetExcelDataSources(xlsxFileInfo));
        }
        if (config.MergeDocuments)
        {
            dataSources = dataSources.GroupBy(x => x.Name)
                .Select(x => new MergedExcelDataSource(x.Key, [.. x], excelDataSourceConfig with { RowNumberColumnName = null }))
                .ToList<IExcelDataSource>();
            if (!config.MergeWorksheets) { return dataSources; }
            return [new MergedExcelDataSource(config.MergedDocumentName ?? "MergedDocument", dataSources,
                        excelDataSourceConfig with { MergedWorksheetNameColumnName = null, RowNumberColumnName = null })];
        }
        return dataSources;
    }

    public void Dispose()
    {
        foreach (ExcelPackage excelPackage in excelPackagesDict.Values)
        {
            excelPackage.Dispose();
        }
    }

    private List<IExcelDataSource> GetExcelDataSources(XlsxFileInfo xlsxFileInfo)
    {
        List<IExcelDataSource> dataSources = [];
        ExcelPackage excelPackage = excelPackagesDict[xlsxFileInfo];
        xlsxFileInfo.Callback?.Invoke(excelPackage);
        HashSet<string>? workSheetNames = config.WorksheetNames is not null ? new(config.WorksheetNames, excelDataSourceConfig.StringComparer) : null;
        ExcelDataSourceConfig xlsxFileDataSourceConfig = excelDataSourceConfig;
        if (config.DocumentNameColumnName is not null)
        {
            xlsxFileDataSourceConfig = excelDataSourceConfig with
            {
                CustomColumnName = config.DocumentNameColumnName,
                CustomColumnValue = xlsxFileInfo.DocumentName ?? xlsxFileInfo.FileInfo.Name
            };
        }
        Dictionary<string, XlsxWorksheetInfo> wsInfoDict = xlsxFileInfo.WorksheetInfos.ToDictionary(x => x.Name, excelDataSourceConfig.StringComparer);
        foreach (ExcelWorksheet excelWorksheet in excelPackage.Workbook.Worksheets)
        {
            if (workSheetNames is not null && !workSheetNames.Contains(excelWorksheet.Name)) { continue; }
            ExcelAddress excelAddress = excelWorksheet.Dimension;
            if (wsInfoDict.TryGetValue(excelWorksheet.Name, out XlsxWorksheetInfo? xlsxWorksheetInfo))
            {
                excelAddress = new(xlsxWorksheetInfo.FromRow, xlsxWorksheetInfo.FromColumn,
                    xlsxWorksheetInfo.ToRow ?? excelAddress.End.Row,
                    xlsxWorksheetInfo.ToColumn ?? excelAddress.End.Column);
            }
            else if (xlsxFileInfo.FromRow != 1 || xlsxFileInfo.FromColumn != 1 || xlsxFileInfo.ToRow != null || xlsxFileInfo.ToColumn != null)
            {
                excelAddress = new(xlsxFileInfo.FromRow, xlsxFileInfo.FromColumn,
                    xlsxFileInfo.ToRow ?? excelAddress.End.Row,
                    xlsxFileInfo.ToColumn ?? excelAddress.End.Column);
            }
            dataSources.Add(new ExcelDataSource(excelWorksheet, xlsxFileDataSourceConfig, excelAddress));
        }
        if (config.MergeWorksheets)
        {
            string name = xlsxFileInfo.MergedWorksheetName ?? xlsxFileInfo.FileInfo.Name;
            return [new MergedExcelDataSource(name, dataSources, excelDataSourceConfig)];
        }
        return dataSources;
    }

}
