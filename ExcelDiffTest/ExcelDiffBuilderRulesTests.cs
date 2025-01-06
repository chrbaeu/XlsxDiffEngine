namespace ExcelDiffTest;

internal class ExcelDiffBuilderRulesTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] content1 = [
        ["Title", "Value"],
        ["A", 100.00],
        ["B", 100.00],
        ["C", 100.00],
        ["D", 100.00],
    ];

    private readonly object?[][] content2 = [
        ["Title", "Value"],
        ["A", 120.00],
        ["B", 120.00],
        ["C", 120.05],
        ["D", 120.00],
    ];

    [Test]
    public void Diff_WithNumberFormatRule()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
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

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithNumberFormatRuleAndTextCompare()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
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

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithMultiplyRule()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content2);
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

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithFormulaRule_AllCells()
    {
        // Arrange
        content1[1][1] = null;
        content2[1][1] = 0;
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content2);
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

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithFormulaRule_NonEmptyCells()
    {
        // Arrange
        content1[1][1] = null;
        content2[1][1] = null;
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content2);
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

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithRegexReplaceRule()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(content1);
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

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

}
