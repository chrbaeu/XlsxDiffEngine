using OfficeOpenXml;

namespace XlsxDiffEngine;

/// <summary>
/// Provides extension methods for exporting data from an <see cref="XlsxDataProvider"/> to Excel workbooks.
/// </summary>
public static class XlsxDataProviderExportExtensions
{
    /// <summary>
    /// Exports the data provider content to an Excel file.
    /// </summary>
    /// <param name="xlsxDataProvider">The data provider whose content should be exported.</param>
    /// <param name="fileInfo">The target Excel file.</param>
    public static void SaveAs(this XlsxDataProvider xlsxDataProvider, FileInfo fileInfo)
    {
        ArgumentNullException.ThrowIfNull(xlsxDataProvider);
        ArgumentNullException.ThrowIfNull(fileInfo);
        using ExcelPackage excelPackage = xlsxDataProvider.CreateExcelPackage();
        excelPackage.SaveAs(fileInfo);
    }

    /// <summary>
    /// Exports the data provider content to an Excel stream.
    /// </summary>
    /// <param name="xlsxDataProvider">The data provider whose content should be exported.</param>
    /// <param name="stream">The target stream.</param>
    public static void SaveAs(this XlsxDataProvider xlsxDataProvider, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(xlsxDataProvider);
        ArgumentNullException.ThrowIfNull(stream);
        using ExcelPackage excelPackage = xlsxDataProvider.CreateExcelPackage();
        excelPackage.SaveAs(stream);
    }

    /// <summary>
    /// Creates an <see cref="ExcelPackage"/> containing the data provider content.
    /// </summary>
    /// <param name="xlsxDataProvider">The data provider whose content should be exported.</param>
    /// <returns>A new <see cref="ExcelPackage"/>. The caller owns and must dispose it.</returns>
    public static ExcelPackage CreateExcelPackage(this XlsxDataProvider xlsxDataProvider)
    {
        ArgumentNullException.ThrowIfNull(xlsxDataProvider);
        ExcelPackage excelPackage = new();
        foreach (IExcelDataSource dataSource in xlsxDataProvider.GetDataSources())
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(dataSource.Name);
            WriteDataSource(dataSource, worksheet);
        }
        return excelPackage;
    }

    private static void WriteDataSource(IExcelDataSource dataSource, ExcelWorksheet worksheet)
    {
        string[] columnNames = dataSource.GetColumnNames().ToArray();
        for (int column = 0; column < columnNames.Length; column++)
        {
            worksheet.Cells[1, column + 1].Value = columnNames[column];
            worksheet.Cells[1, column + 1].Style.Font.Bold = true;
        }
        for (int row = 1; row <= dataSource.DataRows; row++)
        {
            for (int column = 0; column < columnNames.Length; column++)
            {
                ExcelRange dstCell = worksheet.Cells[row + 1, column + 1];
                ExcelRange? srcCell = dataSource.GetExcelRange(columnNames[column], row);
                dstCell.Value = srcCell?.Value ?? dataSource.GetCellValue(columnNames[column], row);
                ExcelHelper.CopyCellStyle(dstCell, srcCell);
                ExcelHelper.CopyCellFormat(dstCell, srcCell);
            }
        }
    }
}
