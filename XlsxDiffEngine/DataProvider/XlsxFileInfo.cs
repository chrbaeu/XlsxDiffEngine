﻿using OfficeOpenXml;
using System.Diagnostics;

namespace XlsxDiffEngine;

/// <summary>
/// Represents metadata and configuration for extracting data from an Excel (.xlsx) file, 
/// including file details and row/column range specifications.
/// </summary>
public sealed record class XlsxFileInfo
{
    private readonly FileInfo? fileInfo;
    private readonly Stream? excelFileDataStream;

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxFileInfo"/> class with the specified file name.
    /// </summary>
    /// <param name="fileName">The name or path of the Excel file.</param>
    public XlsxFileInfo(string fileName)
    {
        ArgumentNullThrowHelper.ThrowIfNull(fileName);
        fileInfo = new(fileName);
        DocumentName = fileInfo.Name;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxFileInfo"/> class with the specified <see cref="FileInfo"/>.
    /// </summary>
    /// <param name="fileInfo">The <see cref="FileInfo"/> object representing the Excel file.</param>
    public XlsxFileInfo(FileInfo fileInfo)
    {
        ArgumentNullThrowHelper.ThrowIfNull(fileInfo);
        this.fileInfo = fileInfo;
        DocumentName = fileInfo.Name;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="XlsxFileInfo"/> class with the specified <see cref="Stream"/> and file name.
    /// </summary>
    /// <param name="excelFileDataStream">The stream containing the Excel file data.</param>
    /// <param name="documentName">The name of the Excel file.</param>
    public XlsxFileInfo(Stream excelFileDataStream, string documentName = "StreamedDocument.xlsx")
    {
        ArgumentNullThrowHelper.ThrowIfNull(excelFileDataStream);
        ArgumentNullThrowHelper.ThrowIfNull(documentName);
        this.excelFileDataStream = excelFileDataStream;
        DocumentName = documentName;
    }

    /// <summary>
    /// The name of the Excel file.
    /// </summary>
    public string DocumentName { get; init; }

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

    /// <summary>
    /// Indicates whether to recalculate formulas in the Excel file before processing.
    /// </summary>
    public bool RecalculateFormulas { get; init; }

    /// <summary>
    /// Retrieves a new <see cref="ExcelPackage"/> instance based on the file information or data stream.
    /// </summary>
    /// <returns>A new <see cref="ExcelPackage"/> instance.</returns>
    public ExcelPackage CreateExcelPackage()
    {
        if (fileInfo is not null)
        {
            return new(fileInfo);
        }
        else if (excelFileDataStream is not null)
        {
            return new(excelFileDataStream);
        }
        throw new UnreachableException("No file information or data is available.");
    }
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
