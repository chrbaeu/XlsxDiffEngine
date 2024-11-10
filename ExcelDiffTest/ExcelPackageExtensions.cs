using OfficeOpenXml;

namespace ExcelDiffTest;

internal static class ExcelPackageExtensions
{
    public static MemoryStream ToMemoryStream(this ExcelPackage excelPackage)
    {
        MemoryStream memoryStream = new();
        excelPackage.SaveAs(memoryStream);
        memoryStream.Position = 0;
        return memoryStream;
    }

    public static ExcelWorksheet AddWorksheet(this ExcelPackage excelPackage, object?[][] data, string worksheetName = "Table")
    {
        ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add(worksheetName);
        for (int row = 0; row < data.Length; row++)
        {
            var rowData = data[row];
            for (int column = 0; column < rowData.Length; column++)
            {
                excelWorksheet.Cells[row + 1, column + 1].Value = rowData[column];
                if (row == 0)
                {
                    excelWorksheet.Cells[row + 1, column + 1].Style.Font.Bold = true;
                }
            }
        }
        return excelWorksheet;
    }

}
