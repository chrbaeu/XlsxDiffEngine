using ExcelDiffEngine;

namespace ExcelDiffTest;

internal class ExcelDiffBuilderValueChangedMarkerTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object[][] oldFileContent = [
        ["Title", "Value"],
        ["A", 100.0],
        ["B", 100.0],
        ["C", 100.0],
        ["D", 100.0],
    ];

    private readonly object[][] newFileContent = [
        ["Title", "Value"],
        ["A", 100.5],
        ["B", 111.0],
        ["C", 130.0],
        ["D", 100.0],
    ];

    [Test]
    public async Task Diff_WithSingleAbsoluteMarker()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddValueChangedMarker(0, 0.6, DefaultCellStyles.YellowValueChangedMarker)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.0, 100.5],
            ["B", "B", 100.0, 111.0],
            ["C", "C", 100.0, 130.0],
            ["D", "D", 100.0, 100.0],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 4, 4], DefaultCellStyles.YellowValueChangedMarker);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithSingleProzentualMarker()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddValueChangedMarker(0.1, 0, DefaultCellStyles.YellowValueChangedMarker)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.0, 100.5],
            ["B", "B", 100.0, 111.0],
            ["C", "C", 100.0, 130.0],
            ["D", "D", 100.0, 100.0],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 4, 4], DefaultCellStyles.YellowValueChangedMarker);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithMultipleAbsolutMarkers()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddValueChangedMarker(0, 0.1, DefaultCellStyles.YellowValueChangedMarker)
            .AddValueChangedMarker(0, 10, DefaultCellStyles.OrangeValueChangedMarker)
            .AddValueChangedMarker(0, 25, DefaultCellStyles.RedValueChangedMarker)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.0, 100.5],
            ["B", "B", 100.0, 111.0],
            ["C", "C", 100.0, 130.0],
            ["D", "D", 100.0, 100.0],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 3, 2, 4], DefaultCellStyles.YellowValueChangedMarker);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.OrangeValueChangedMarker);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.RedValueChangedMarker);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithMultipleProzentualMarkers()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddValueChangedMarker(0.001, 0, DefaultCellStyles.YellowValueChangedMarker)
            .AddValueChangedMarker(0.10, 0, DefaultCellStyles.OrangeValueChangedMarker)
            .AddValueChangedMarker(0.25, 0, DefaultCellStyles.RedValueChangedMarker)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.0, 100.5],
            ["B", "B", 100.0, 111.0],
            ["C", "C", 100.0, 130.0],
            ["D", "D", 100.0, 100.0],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 3, 2, 4], DefaultCellStyles.YellowValueChangedMarker);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.OrangeValueChangedMarker);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.RedValueChangedMarker);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

}
