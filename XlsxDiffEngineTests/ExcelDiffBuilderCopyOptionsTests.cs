using System.Drawing;

namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderCopyOptionsTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] oldFileContent = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3],
    ];

    private readonly object?[][] newFileContent = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 4],
        ["C", 3],
    ];

    [Test]
    public void Diff_WithNumberFormatCopyCellFormatTrue()
    {
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellFormat(true)
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        expectedResult.Workbook.Worksheets[0].Cells[2, 3, 4, 4].Style.Numberformat.Format = "0.00";

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithNumberFormatAndCopyCellFormatFalse()
    {
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellFormat(false)
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithStyleAndCopyCellStyleTrue()
    {
        CellStyle cellStyle = new()
        {
            Bold = true,
            Italic = true,
            Underline = true,
            FontColor = Color.Red,
            BackgroundColor = Color.Blue
        };
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellStyle(true)
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 4], cellStyle);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithStyleAndCopyCellStyleFalse()
    {
        CellStyle cellStyle = new()
        {
            Bold = true,
            Italic = true,
            Underline = true,
            FontColor = Color.Red,
            BackgroundColor = Color.Blue
        };
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellStyle(false)
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }
}
