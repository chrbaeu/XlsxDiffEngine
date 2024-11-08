using OfficeOpenXml;

namespace ExcelDiffEngine;

/// <summary>
/// A builder class for configuring <see cref="XlsxFileInfo"/> instances representing old and new Excel (.xlsx) files,
/// including options for file paths, worksheet configurations, merged worksheet names, and data extraction ranges.
/// </summary>
public class ExcelDiffXlsxFileConfigBuilder
{
    private XlsxFileInfo oldFile = new("Unspecified.xlsx");
    private XlsxFileInfo newFile = new("Unspecified.xlsx");

    /// <summary>
    /// Sets the file path and optional callback for the "old" file in the comparison.
    /// </summary>
    /// <param name="filePath">The path to the old Excel file.</param>
    /// <param name="callback">An optional callback to execute with the <see cref="ExcelPackage"/> for additional configuration.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder SetOldFile(string filePath, Action<ExcelPackage>? callback = null)
    {
        oldFile = oldFile with { FileInfo = new(filePath), PrepareExcelPackageCallback = callback };
        return this;
    }

    /// <summary>
    /// Sets the file path and optional callback for the "new" file in the comparison.
    /// </summary>
    /// <param name="filePath">The path to the new Excel file.</param>
    /// <param name="callback">An optional callback to execute with the <see cref="ExcelPackage"/> for additional configuration.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder SetNewFile(string filePath, Action<ExcelPackage>? callback = null)
    {
        newFile = newFile with { FileInfo = new(filePath), PrepareExcelPackageCallback = callback };
        return this;
    }

    /// <summary>
    /// Sets a merged worksheet name to be used for both the old and new files in the comparison.
    /// </summary>
    /// <param name="mergedWorksheetName">The name to assign to the merged worksheet.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder SetMergedWorksheetName(string mergedWorksheetName)
    {
        oldFile = oldFile with { MergedWorksheetName = mergedWorksheetName };
        newFile = newFile with { MergedWorksheetName = mergedWorksheetName };
        return this;
    }

    /// <summary>
    /// Sets a document name to be used for both the old and new files in the comparison.
    /// </summary>
    /// <param name="documentName">The name to assign to the document.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder SetDocumentName(string documentName)
    {
        oldFile = oldFile with { DocumentName = documentName };
        newFile = newFile with { DocumentName = documentName };
        return this;
    }

    /// <summary>
    /// Adds worksheet-specific information for data extraction from the old file.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet to configure.</param>
    /// <param name="fromRow">The starting row for data extraction. Default is 1.</param>
    /// <param name="fromColumn">The starting column for data extraction. Default is 1.</param>
    /// <param name="toRow">The ending row for data extraction. If null, there is no limit.</param>
    /// <param name="toColumn">The ending column for data extraction. If null, there is no limit.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder AddOldFileWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with
        {
            WorksheetInfos = [.. oldFile.WorksheetInfos, new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }]
        };
        return this;
    }

    /// <summary>
    /// Adds worksheet-specific information for data extraction from the new file.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet to configure.</param>
    /// <param name="fromRow">The starting row for data extraction. Default is 1.</param>
    /// <param name="fromColumn">The starting column for data extraction. Default is 1.</param>
    /// <param name="toRow">The ending row for data extraction. If null, there is no limit.</param>
    /// <param name="toColumn">The ending column for data extraction. If null, there is no limit.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder AddNewFileWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        newFile = newFile with
        {
            WorksheetInfos = [.. newFile.WorksheetInfos, new() { Name = worksheetName, FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn }]
        };
        return this;
    }

    /// <summary>
    /// Adds worksheet-specific information for data extraction to both the old and new files.
    /// </summary>
    /// <param name="worksheetName">The name of the worksheet to configure.</param>
    /// <param name="fromRow">The starting row for data extraction. Default is 1.</param>
    /// <param name="fromColumn">The starting column for data extraction. Default is 1.</param>
    /// <param name="toRow">The ending row for data extraction. If null, there is no limit.</param>
    /// <param name="toColumn">The ending column for data extraction. If null, there is no limit.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder AddWorksheetInfo(string worksheetName, int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        _ = AddOldFileWorksheetInfo(worksheetName, fromRow, fromColumn, toRow, toColumn);
        _ = AddNewFileWorksheetInfo(worksheetName, fromRow, fromColumn, toRow, toColumn);
        return this;
    }

    /// <summary>
    /// Sets a data extraction area for both the old and new files.
    /// </summary>
    /// <param name="fromRow">The starting row for data extraction. Default is 1.</param>
    /// <param name="fromColumn">The starting column for data extraction. Default is 1.</param>
    /// <param name="toRow">The ending row for data extraction. If null, there is no limit.</param>
    /// <param name="toColumn">The ending column for data extraction. If null, there is no limit.</param>
    /// <returns>The current builder instance, allowing for method chaining.</returns>
    public ExcelDiffXlsxFileConfigBuilder SetDataArea(int fromRow = 1, int fromColumn = 1, int? toRow = null, int? toColumn = null)
    {
        oldFile = oldFile with { FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn };
        newFile = newFile with { FromRow = fromRow, FromColumn = fromColumn, ToRow = toRow, ToColumn = toColumn };
        return this;
    }

    /// <summary>
    /// Builds and returns the configured <see cref="XlsxFileInfo"/> instances for the old and new files.
    /// If no merged worksheet name has been set, it defaults to the file name of the new file (without extension).
    /// </summary>
    /// <returns>A tuple containing the configured old and new <see cref="XlsxFileInfo"/> instances.</returns>
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
