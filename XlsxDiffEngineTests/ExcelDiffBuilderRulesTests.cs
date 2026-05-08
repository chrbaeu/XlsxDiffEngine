namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderRulesTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] content = ExcelTestData.NumericRuleBase();

    private readonly object?[][] transformedContent = ExcelTestData.NumericRuleBase()
        .Select((row, index) => index == 0 ? row : [row[0], 120.00])
        .ToArray();

    [Test]
    public async Task Diff_WithNumberFormatRule()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        newExcelPackage.Workbook.Worksheets[0].Cells[3, 2].Value = 100.04;
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 2].Value = 100.06;
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddModificationRules([
                new("Value", ModificationKind.NumberFormat, "0.0", DataKind.All),
                ])
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.0, 100.0],
            ["B", "B", 100.0, 100.0],
            ["C", "C", 100.0, 100.1],
            ["D", "D", 100.0, 100.0],
            ]);
        expectedResult.Workbook.Worksheets[0].Cells[2, 3, 5, 4].Style.Numberformat.Format = "0.0";
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 4, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithNumberFormatRuleAndTextCompare()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        newExcelPackage.Workbook.Worksheets[0].Cells[3, 2].Value = 100.04;
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 2].Value = 100.06;
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddModificationRules([
                new("Value", ModificationKind.NumberFormat, "0.0", DataKind.All),
                ])
            .SetColumnsToTextCompareOnly("Value")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.0, 100.0],
            ["B", "B", 100.0, 100.0],
            ["C", "C", 100.0, 100.1],
            ["D", "D", 100.0, 100.0],
            ]);
        expectedResult.Workbook.Worksheets[0].Cells[2, 3, 5, 4].Style.Numberformat.Format = "0.0";
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithMultiplyRule()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(transformedContent);
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 2].Value = 120.05;
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddModificationRules([
                new("Value", ModificationKind.Multiply, "1.2", DataKind.Old),
                ])
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 120.00, 120.00],
            ["B", "B", 120.00, 120.00],
            ["C", "C", 120.00, 120.05],
            ["D", "D", 120.00, 120.00],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithFormulaRule_AllCells()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(transformedContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Value = 0;
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 2].Value = 120.05;
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddModificationRules([
                new("Value", ModificationKind.Formula, "={#}*1.2", DataKind.Old),
                ])
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 0, 0],
            ["B", "B", 120.00, 120.00],
            ["C", "C", 120.00, 120.05],
            ["D", "D", 120.00, 120.00],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithFormulaRule_NonEmptyCells()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(transformedContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 2].Value = 120.05;
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddModificationRules([
                new("Value", ModificationKind.Formula, "={#}*1.2", DataKind.OldNonEmpty),
                ])
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", null, null],
            ["B", "B", 120.00, 120.00],
            ["C", "C", 120.00, 120.05],
            ["D", "D", 120.00, 120.00],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithRegexReplaceRule()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content);
        newExcelPackage.Workbook.Worksheets[0].Cells[4, 2].Value = 200.00;
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddModificationRules([
                new("Title", ModificationKind.RegexReplace, "B", DataKind.All, "X"),
                ])
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 100.00, 100.00],
            ["X", "X", 100.00, 100.00],
            ["C", "C", 100.00, 200.00],
            ["D", "D", 100.00, 100.00],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[4, 3, 4, 4], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public async Task Diff_WithCaseSensitiveHeaderRules_DoesNotApplyOnDifferentHeaderCase()
    {
        // Arrange
        object[][] oldFile = [
            ["Title", "Value"],
            ["A", 100.00],
        ];
        object[][] newFile = [
            ["title", "Value"],
            ["A", 100.00],
        ];
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .IgnoreHeaderCase(false)
            .AddModificationRules([
                new("title", ModificationKind.RegexReplace, "A", DataKind.All, "X"),
                ])
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["title", "title", "Value", "Value", "Title", "Title"],
            [null, "X", 100.00, 100.00, "A", null],
        ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 2, 2], DefaultCellStyles.ChangedCell);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 5, 2, 6], DefaultCellStyles.ChangedCell);

        await ExcelTestHelper.AssertExcelPackagesIdentical(result, expectedResult, true);
    }

}
