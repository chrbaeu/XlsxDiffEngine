using ExcelDiffEngine;

namespace ExcelDiffTest;

internal class ExcelDiffBuilderAdvancedTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object[][] dataTab1 = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3]
    ];

    private readonly object[][] dataTab2 = [
        ["Title", "Value"],
        ["D", 4],
        ["E", 5],
        ["F", 6]
    ];

    [Test]
    public async Task Diff_MultipleWorksheets_WithoutMerging()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        oldExcelPackage.AddWorksheet(dataTab2, "Tab2");
        using var oldFileStream = oldExcelPackage.ToMemoryStream();

        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        var ws = newExcelPackage.AddWorksheet(dataTab2, "Tab2");
        ws.Cells[3, 2].Value = 8;
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 2],
            ["C", "C", 3, 3],
        ], "Tab1");
        expectedResult.AddWorksheet([
            ["Title", "Title", "Value", "Value"],
            ["D", "D", 4, 4],
            ["E", "E", 5, 8],
            ["F", "F", 6, 6]
        ], "Tab2");
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[1].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_MultipleWorksheets_WithMerging()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        oldExcelPackage.AddWorksheet(dataTab2, "Tab2");
        using var oldFileStream = oldExcelPackage.ToMemoryStream();

        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        var ws = newExcelPackage.AddWorksheet(dataTab2, "Tab2");
        ws.Cells[3, 2].Value = 8;
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                .SetMergedWorksheetName("Merged")
                )
            .MergeWorkSheets()
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 2],
            ["C", "C", 3, 3],
            ["D", "D", 4, 4],
            ["E", "E", 5, 8],
            ["F", "F", 6, 6]
            ], "Merged");
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 3, 6, 4], DefaultCellStyles.ChangedCell);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

}
