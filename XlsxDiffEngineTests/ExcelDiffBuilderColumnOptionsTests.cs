namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderColumnOptionsTests
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
    public void Diff_WithIgnoreColumnsNotInBoth()
    {
        // Arrange
        object?[][] oldFile = [
            ["Title", "Value", "OldOnly"],
            ["A", 1, "x"],
            ["B", 2, "y"],
        ];
        object?[][] newFile = [
            ["Title", "Value", "NewOnly"],
            ["A", 1, "foo"],
            ["B", 2, "bar"],
        ];
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
            .IgnoreColumnsNotInBoth(true)
            .Build();

        // Assert
        using var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 2],
        ]);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithShowHideOldColumnsAndShowColumns()
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
            .HideOldColumns()
            .ShowColumns("Value")
            .Build();

        // Assert
        var worksheet = result.Workbook.Worksheets[0];
        await Assert.That(worksheet.Column(1).Hidden).IsTrue();
        await Assert.That(worksheet.Column(2).Hidden).IsFalse();
        await Assert.That(worksheet.Column(3).Hidden).IsFalse();
        await Assert.That(worksheet.Column(4).Hidden).IsFalse();
    }

    [Test]
    public async Task Diff_WithHideColumns()
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
            .HideColumns("Title")
            .Build();

        // Assert
        var worksheet = result.Workbook.Worksheets[0];
        await Assert.That(worksheet.Column(1).Hidden).IsTrue();
        await Assert.That(worksheet.Column(2).Hidden).IsTrue();
        await Assert.That(worksheet.Column(3).Hidden).IsFalse();
        await Assert.That(worksheet.Column(4).Hidden).IsFalse();
    }

    [Test]
    public async Task Diff_WithSetColumnSizes()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();
        double[] sizes = [12.5, 20.0];

        // Act
        using var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
            )
            .SetColumnSizes(sizes)
            .Build();

        // Assert
        var worksheet = result.Workbook.Worksheets[0];
        await Assert.That(worksheet.Column(1).Width).IsEqualTo(sizes[0]);
        await Assert.That(worksheet.Column(2).Width).IsEqualTo(sizes[1]);
    }

    [Test]
    public async Task Diff_WithSetColumnSizeByIndex()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();
        int columnIndex = 2;
        double width = 33.3;

        // Act
        using var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
            )
            .SetColumnSize(columnIndex, width)
            .Build();

        // Assert
        var worksheet = result.Workbook.Worksheets[0];
        await Assert.That(worksheet.Column(columnIndex).Width).IsEqualTo(width);
    }

    [Test]
    public async Task Diff_WithSetColumnSizeByName()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();
        double width = 44.4;

        // Act
        using var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
            )
            .SetColumnSize("Value", width)
            .Build();

        // Assert
        var worksheet = result.Workbook.Worksheets[0];
        await Assert.That(worksheet.Column(3).Width).IsEqualTo(width);
    }
}
