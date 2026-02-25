namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderTypesTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] oldFileContent = [
        ["Title", "Value"],
        ["A", null],
        ["B", null],
    ];

    private readonly object?[][] newFileContent = [
        ["Title", "Value"],
        ["A", null],
        ["B", null],
    ];

    [Test]
    public void Diff_DateTime()
    {
        // Arrange
        var a = oldFileContent[1][1] = newFileContent[1][1] = oldFileContent[2][1] = new DateTime(2026, 1, 25);
        var b = newFileContent[2][1] = new DateTime(2026, 1, 26);
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

        result.SaveAs(@"C:\Users\CBaeu\Desktop\Result.xlsx");

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", a, a],
            ["B", "B", a, b],
            ]);

        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_Bool()
    {
        // Arrange
        var a = oldFileContent[1][1] = newFileContent[1][1] = oldFileContent[2][1] = true;
        var b = newFileContent[2][1] = false;
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
            ["A", "A", a, a],
            ["B", "B", a, b],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }


    [Test]
    public void Diff_NumbersAsString()
    {
        // Arrange
        var a = oldFileContent[1][1] = oldFileContent[2][1] = "55";
        var b = newFileContent[1][1] = "56";
        var c = newFileContent[2][1] = "57";
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
            .AddValueChangedMarker(0, 1.5, DefaultCellStyles.YellowValueChangedMarker)
            .SetColumnsToCompareAsNumbers("Value")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", a, b],
            ["B", "B", a, c],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.YellowValueChangedMarker);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

}
