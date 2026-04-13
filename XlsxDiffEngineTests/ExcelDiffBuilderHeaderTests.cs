namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderHeaderTests
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
    public void Diff_WithColumnHeaderPostfix()
    {
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetNewHeaderColumnPostfix("New")
            .SetOldHeaderColumnPostfix("Old")
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["TitleOld", "TitleNew", "ValueOld", "ValueNew"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithCustomHeaderRows()
    {
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetHeader("Custom1", "Custom2")
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Custom1", null, null, null],
            ["Custom2", null, null, null],
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
        ]);
        expectedResult.Workbook.Worksheets[0].Cells[1, 1, 1, 4].Style.Font.Bold = false;
        expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 4].Style.Font.Bold = true;
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 3, 5, 4], DefaultCellStyles.ChangedCell);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_WithCustomHeaderRowsAndColumns()
    {
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetHeader([
                ["Top1", "Top2", "Top3", "Top4"],
                ["Sub1", "Sub2", "Sub3", "Sub4"]
            ])
            .Build();

        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Top1", "Top2", "Top3", "Top4"],
            ["Sub1", "Sub2", "Sub3", "Sub4"],
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
        ]);
        expectedResult.Workbook.Worksheets[0].Cells[1, 1, 1, 4].Style.Font.Bold = false;
        expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 4].Style.Font.Bold = true;
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[5, 3, 5, 4], DefaultCellStyles.ChangedCell);
        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }
}
