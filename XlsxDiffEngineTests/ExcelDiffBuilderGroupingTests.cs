namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderGroupingTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] oldFileContent = [
        ["Title", "Group", "Value"],
        ["A", "1", 1],
        ["B", "1", 2],
        ["C", "2", 3],
        ["D", "2", 4],
    ];

    private readonly object?[][] newFileContent = [
        ["Title", "Group", "Value"],
        ["A", "1", 1],
        ["E", "1", 5],
        ["C", "2", 3],
        ["F", "2", 6],
    ];

    [Test]
    public void Diff_WithKeyAndGroupKey()
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
            .SetGroupKeyColumns("Group")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Group", "Group", "Value", "Value"],
            ["A", "A", "1", "1", 1, 1],
            [null, "E", null, "1", null, 5],
            ["B", null, "1", null, 2, null],
            ["C", "C", "2", "2", 3, 3],
            [null, "F", null, "2", null, 6],
            ["D", null, "2", null, 4, null],
        ]);

        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 6], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 1, 4, 6], DefaultCellStyles.RemovedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 1, 6, 6], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[7, 1, 7, 6], DefaultCellStyles.RemovedRow);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithKeyAndGroupKeyAndEmptyRowBetweenGroups()
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
            .SetGroupKeyColumns("Group")
            .AddEmptyRowAfterGroups()
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Group", "Group", "Value", "Value"],
            [null, null, null, null, null, null],
            ["A", "A", "1", "1", 1, 1],
            [null, "E", null, "1", null, 5],
            ["B", null, "1", null, 2, null],
            [null, null, null, null, null, null],
            ["C", "C", "2", "2", 3, 3],
            [null, "F", null, "2", null, 6],
            ["D", null, "2", null, 4, null],
        ]);

        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 1, 4, 6], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 5, 6], DefaultCellStyles.RemovedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[8, 1, 8, 6], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[9, 1, 9, 6], DefaultCellStyles.RemovedRow);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

}
