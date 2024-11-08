using OfficeOpenXml;

namespace ExcelDiffEngine;

/// <summary>
/// Represents metadata and configuration for extracting data from an Excel (.xlsx) file, 
/// including file details and row/column range specifications.
/// </summary>
public sealed record class XlsxFileInfo
{
    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxFileInfo"/> class with the specified file name.
    /// </summary>
    /// <param name="fileName">The name or path of the Excel file.</param>
    public XlsxFileInfo(string fileName)
    {
        FileInfo = new(fileName);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxFileInfo"/> class with the specified <see cref="FileInfo"/>.
    /// </summary>
    /// <param name="fileInfo">The <see cref="FileInfo"/> object representing the Excel file.</param>
    public XlsxFileInfo(FileInfo fileInfo)
    {
        FileInfo = fileInfo;
    }

    /// <summary>
    /// The <see cref="FileInfo"/> object representing the Excel file.
    /// </summary>
    public FileInfo FileInfo { get; init; }

    /// <summary>
    /// The starting row for data extraction. Default is 1.
    /// </summary>
    public int FromRow { get; init; } = 1;

    /// <summary>
    /// The starting column for data extraction. Default is 1.
    /// </summary>
    public int FromColumn { get; init; } = 1;

    /// <summary>
    /// The ending row for data extraction. If null, there is no limit.
    /// </summary>
    public int? ToRow { get; init; }

    /// <summary>
    /// The ending column for data extraction. If null, there is no limit.
    /// </summary>
    public int? ToColumn { get; init; }

    /// <summary>
    /// The name of the document, if specified.
    /// </summary>
    public string? DocumentName { get; init; }

    /// <summary>
    /// The name used for the merged worksheet, if worksheets are merged.
    /// </summary>
    public string? MergedWorksheetName { get; init; }

    /// <summary>
    /// The list of worksheet information, detailing each worksheet’s data extraction configuration.
    /// </summary>
    public IReadOnlyList<XlsxWorksheetInfo> WorksheetInfos { get; init; } = [];

    /// <summary>
    /// An optional callback action to configure or modify the <see cref="ExcelPackage"/> before processing.
    /// </summary>
    public Action<ExcelPackage>? PrepareExcelPackageCallback { get; init; }
}

/// <summary>
/// Represents configuration for a specific worksheet within an Excel (.xlsx) file, 
/// including row and column range specifications for data extraction.
/// </summary>
public sealed record XlsxWorksheetInfo
{
    /// <summary>
    /// The name of the worksheet.
    /// </summary>
    public string Name { get; set; } = "";

    /// <summary>
    /// The starting row for data extraction within the worksheet. Default is 1.
    /// </summary>
    public int FromRow { get; init; } = 1;

    /// <summary>
    /// The starting column for data extraction within the worksheet. Default is 1.
    /// </summary>
    public int FromColumn { get; init; } = 1;

    /// <summary>
    /// The ending row for data extraction. If null, there is no limit.
    /// </summary>
    public int? ToRow { get; init; }

    /// <summary>
    /// The ending column for data extraction. If null, there is no limit.
    /// </summary>
    public int? ToColumn { get; init; }
}
