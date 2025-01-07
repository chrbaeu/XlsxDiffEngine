using System.Drawing;

namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderStylesTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly CellStyle myHeader = new() { Underline = true, Bold = false, BackgroundColor = Color.FromArgb(200, 200, 200) };
    private readonly CellStyle myAddedRow = new() { Italic = true };
    private readonly CellStyle myRemovedRow = new() { Underline = true };
    private readonly CellStyle myChangedCell = new() { FontColor = Color.FromArgb(255, 178, 101) };
    private readonly CellStyle myChangedRowKeyColumns = new() { Bold = true, FontColor = Color.FromArgb(150, 175, 255) };

    private readonly object?[][] oldFileContent = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3],
        ["D", 4],
    ];

    private readonly object?[][] newFileContent = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 4],
        ["C", 3],
        ["E", 5],
    ];

    private readonly object?[][] diffFileContent = [
        ["Title", "Title", "Value", "Value"],
        ["A", "A", 1, 1],
        ["B", "B", 2, 4],
        ["C", "C", 3, 3],
        [null, "E", null, 5],
        ["D", null, 4, null],
    ];



    [Test]
    public void Diff_WithDefaultStyles()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetKeyColumns("Title")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage(diffFileContent);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 5, 4], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 1, 6, 4], DefaultCellStyles.RemovedRow);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithCustomStyles()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetHeaderStyle(myHeader)
            .SetAddedRowStyle(myAddedRow)
            .SetRemovedRowStyle(myRemovedRow)
            .SetChangedCellStyle(myChangedCell)
            .SetChangedRowKeyColumnsStyle(myChangedRowKeyColumns)
            .SetKeyColumns("Title")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage(diffFileContent);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[1, 1, 1, 4], myHeader);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], myChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], myChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 6, 2], myChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 3, 6, 4], myChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 5, 4], myAddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 1, 6, 4], myRemovedRow);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithCopyCellStyles()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[1, 1, 5, 2], new CellStyle() { FontColor = Color.FromArgb(100, 100, 100) });
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[1, 1, 5, 2], new CellStyle() { Italic = true });
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetKeyColumns("Title")
            .CopyCellStyle()
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage(diffFileContent);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 5, 4], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 1, 6, 4], DefaultCellStyles.RemovedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 1], new CellStyle() { FontColor = Color.FromArgb(100, 100, 100) });
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 3, 4, 3], new CellStyle() { FontColor = Color.FromArgb(100, 100, 100) });
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 2, 5, 2], new CellStyle() { Italic = true });
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 4, 5, 4], new CellStyle() { Italic = true });

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

}
