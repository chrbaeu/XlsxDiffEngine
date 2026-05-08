namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderBasicTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] oldFileContent = ExcelTestData.StandardOld();

    private readonly object?[][] newFileContent = ExcelTestData.StandardNew();

    [Test]
    public async Task Diff_WithEmptyWorksheets()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(Array.Empty<object?[]>());
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(Array.Empty<object?[]>());
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage(Array.Empty<object?[]>());
        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public async Task Diff_FullAgainstEmptyWorksheets()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(Array.Empty<object?[]>());
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
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", null, 1, null],
            ["B", null, 2, null],
            ["C", null, 3, null],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 4], DefaultCellStyles.RemovedRow);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_EmptyAgainstFullWorksheets()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(Array.Empty<object?[]>());
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
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
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            [null, "A", null, 1],
            [null, "B", null, 2],
            [null, "C", null, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 4], DefaultCellStyles.AddedRow);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithRecalculation()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Formula = "=10-9";
        newExcelPackage.Workbook.Worksheets[0].Cells[3, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[3, 2].Formula = "=10-6";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                .RecalculateFormulas()
                )
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithEmptyWorksheet()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = new();
        oldExcelPackage.Workbook.Worksheets.Add("Table");
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            [null, "A", null, 1],
            [null, "B", null, 4],
            [null, "C", null, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 4], DefaultCellStyles.AddedRow);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }
}
