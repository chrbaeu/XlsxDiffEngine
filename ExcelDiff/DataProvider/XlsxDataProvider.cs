using OfficeOpenXml;

namespace ExcelDiffEngine;

/// <summary>
/// Provides an unified interface to read data from one or multiple Excel (.xlsx) files, supporting 
/// advanced options like merging worksheets or documents and applying custom configurations.
/// </summary>
public sealed class XlsxDataProvider : IDisposable
{
    private readonly ICollection<XlsxFileInfo> xlsxFileInfos;
    private readonly Dictionary<XlsxFileInfo, ExcelPackage> excelPackagesDict = [];
    private readonly XlsxDataProviderConfig config;
    private readonly ExcelDataSourceConfig excelDataSourceConfig;

    private List<IExcelDataSource> dataSources = [];

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxDataProvider"/> class for a single Excel file.
    /// </summary>
    /// <param name="xlsxFile">The path to the Excel file.</param>
    /// <param name="headerRow">The header row number for the data. Default is 1.</param>
    /// <param name="config">Optional configuration for data processing.</param>
    public XlsxDataProvider(string xlsxFile, int headerRow = 1, XlsxDataProviderConfig? config = null)
        : this([new(xlsxFile) { FromRow = headerRow }], config)
    { }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxDataProvider"/> class for a single <see cref="XlsxFileInfo"/> object.
    /// </summary>
    /// <param name="xlsxFile">The file information for the Excel file.</param>
    /// <param name="config">Optional configuration for data processing.</param>
    public XlsxDataProvider(XlsxFileInfo xlsxFile, XlsxDataProviderConfig? config = null)
        : this([xlsxFile], config)
    { }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxDataProvider"/> class for multiple Excel files.
    /// </summary>
    /// <param name="xlsxFiles">A collection of file information for the Excel files.</param>
    /// <param name="config">Optional configuration for data processing.</param>
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

    /// <summary>
    /// Retrieves a list of data sources representing virtual worksheets that are build based on the files and configurations.
    /// </summary>
    /// <returns>A read-only list of <see cref="IExcelDataSource"/> instances.</returns>
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

    /// <summary>
    /// Disposes of the resources used by the <see cref="XlsxDataProvider"/> instance, including closing any open Excel packages.
    /// </summary>
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
        xlsxFileInfo.PrepareExcelPackageCallback?.Invoke(excelPackage);
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
