namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderSecondaryKeyColumnsTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    [Test]
    public async Task Diff_WithSecondaryKeyColumns_MatchesOnSecondaryKey()
    {
        // Arrange
        object?[][] oldFile = ExcelTestData.SecondaryKeyOld();
        object?[][] newFile = ExcelTestData.SecondaryKeyNew();
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
            )
            .SetKeyColumns("ID")
            .SetSecondaryKeyColumns("SecondaryID")
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["ID", "ID", "SecondaryID", "SecondaryID", "Value", "Value"],
            ["1", "3", "A", "A", 100, 100],
            ["2", "4", "B", "B", 200, 250],
        ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 5, 3, 6], DefaultCellStyles.ChangedCell);
        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }
}

