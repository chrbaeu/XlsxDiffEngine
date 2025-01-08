using OfficeOpenXml;
using XlsxDiffEngine;

namespace XlsxDiffSample;

internal class Program
{
    internal static void Main()
    {
        Console.WriteLine("ExcelDiff Sample");

        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using ExcelPackage excelPackage = new();
        ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Table");
        excelWorksheet.Cells[1, 1].Value = "Title";
        excelWorksheet.Cells[1, 2].Value = "Value";
        excelWorksheet.Cells[2, 1].Value = "A";
        excelWorksheet.Cells[2, 2].Value = 1;
        excelWorksheet.Cells[3, 1].Value = "B";
        excelWorksheet.Cells[3, 2].Value = 2;
        excelWorksheet.Cells[4, 1].Value = "C";
        excelWorksheet.Cells[4, 2].Value = 3;
        excelPackage.SaveAs(new FileInfo("OldFile.xlsx"));

        excelWorksheet.Cells[3, 2].Value = 4;
        excelPackage.SaveAs(new FileInfo("NewFile.xlsx"));

        new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile("OldFile.xlsx")
                .SetNewFile("NewFile.xlsx")
                )
            .SetKeyColumns("Title")
            .SetOldHeaderColumnComment("Old data")
            .SetNewHeaderColumnComment("New data")
            .Build("ComparisonOutput.xlsx");

    }
}
