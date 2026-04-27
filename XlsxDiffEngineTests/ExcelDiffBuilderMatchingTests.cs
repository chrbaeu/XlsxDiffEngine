namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderMatchingTests
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
    public void Diff_WithHighlighting()
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
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithHighlightingAndKeyColumn()
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
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithHighlightingAndKeyColumnAndInsertAndDelete()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[4, 1].Value = "D";
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
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            [null, "C", null, 3],
            ["D", null, 3, null],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 1, 4, 4], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 5, 4], DefaultCellStyles.RemovedRow);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithColumnsToCompare()
    {
        // Arrange
        object?[][] oldFile = [
            ["Title", "Value", "Other"],
            ["A", 1, "x"],
            ["B", 2, "y"],
        ];
        object?[][] newFile = [
            ["Title", "Value", "Other"],
            ["A", 1, "x"],
            ["C", 4, "z"],
        ];
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using var result = new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetColumnsToCompare("Value")
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value", "Other", "Other"],
            ["A", "A", 1, 1, "x", "x"],
            ["B", "C", 2, 4, "y", "z"],
        ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithSort()
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
            .SetColumnsToSortBy("Value")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["C", "C", 3, 3],
            ["B", "B", 2, 4],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithOldValueFallback()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 1].Value = "D";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetKeyColumns("Title")
            .SetColumnsToFillWithOldValueIfNoNewValueExists("Title")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            [null, "D", null, 3],
            ["C", "C", 3, null],
            ]);

        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 1, 4, 4], DefaultCellStyles.AddedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 1, 5, 4], DefaultCellStyles.RemovedRow);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 2], DefaultCellStyles.FallbackValue);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithIgnoreCase()
    {
        // Arrange
        object?[][] oldFile = [
            ["Title", "Value"],
            ["A", 1],
            ["b", 2],
        ];
        object?[][] newFile = [
            ["title", "value"],
            ["A", 1],
            ["B", 2],
        ];
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx"))
            .IgnoreCase(true)
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
                ["title", "title", "value", "value"],
                ["A", "A", 1, 1],
                ["b", "B", 2, 2],
            ]);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithoutIgnoreCase()
    {
        // Arrange
        object?[][] oldFile = [
            ["Title", "Value"],
            ["A", 1],
            ["b", 2],
        ];
        object?[][] newFile = [
            ["Title", "Value"],
            ["A", 1],
            ["B", 2],
        ];
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx"))
            .IgnoreCase(false)
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
                ["Title", "Title", "Value", "Value"],
                ["A", "A", 1, 1],
                ["b", "B", 2, 2],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedCell);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithIgnoreHeaderCaseAndCaseSensitiveData()
    {
        // Arrange
        object[][] oldFile = [
            ["Title", "Value"],
            ["b", 2],
        ];
        object[][] newFile = [
            ["title", "value"],
            ["B", 2],
        ];
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx"))
            .IgnoreHeaderCase(true)
            .IgnoreDataCase(false)
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
                ["title", "title", "value", "value"],
                ["b", "B", 2, 2],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 2, 2], DefaultCellStyles.ChangedCell);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithCaseSensitiveHeadersAndIgnoreDataCase()
    {
        // Arrange
        object[][] oldFile = [
            ["Title", "Value"],
            ["b", 2],
        ];
        object[][] newFile = [
            ["title", "Value"],
            ["B", 2],
        ];
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx"))
            .IgnoreHeaderCase(false)
            .IgnoreDataCase(true)
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
                ["title", "title", "Value", "Value", "Title", "Title"],
                [null, "B", 2, 2, "b", null],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 2, 2], DefaultCellStyles.ChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 5, 2, 6], DefaultCellStyles.ChangedCell);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithIgnoreDataCase_KeyColumnsAreMatchedCaseInsensitively()
    {
        // Arrange
        object[][] oldFile = [
            ["Title", "Value"],
            ["b", 2],
        ];
        object[][] newFile = [
            ["Title", "Value"],
            ["B", 2],
        ];
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx"))
            .SetKeyColumns("Title")
            .IgnoreDataCase(true)
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
                ["Title", "Title", "Value", "Value"],
                ["b", "B", 2, 2],
            ]);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }
}
