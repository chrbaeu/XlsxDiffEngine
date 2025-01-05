namespace ExcelDiffTest;

internal class ExcelDiffBuilderMergingTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object?[][] dataTab1 = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3],
    ];

    private readonly object?[][] dataTab2 = [
        ["Title", "Value"],
        ["D", 4],
        ["E", 5],
        ["F", 6],
    ];

    [Test]
    public void Diff_MultipleWorksheets_WithoutMerging()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        oldExcelPackage.AddWorksheet(dataTab2, "Tab2");
        using var oldFileStream = oldExcelPackage.ToMemoryStream();

        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        ExcelWorksheet ws = newExcelPackage.AddWorksheet(dataTab2, "Tab2");
        ws.Cells[3, 2].Value = 8;
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
            ["B", "B", 2, 2],
            ["C", "C", 3, 3],
        ], "Tab1");
        expectedResult.AddWorksheet([
            ["Title", "Title", "Value", "Value"],
            ["D", "D", 4, 4],
            ["E", "E", 5, 8],
            ["F", "F", 6, 6],
        ], "Tab2");
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[1].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_MultipleWorksheets_WithMerging()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        oldExcelPackage.AddWorksheet(dataTab2, "Tab2");
        using var oldFileStream = oldExcelPackage.ToMemoryStream();

        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1, "Tab1");
        ExcelWorksheet ws = newExcelPackage.AddWorksheet(dataTab2, "Tab2");
        ws.Cells[3, 2].Value = 8;
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                .SetMergedWorksheetName("Merged")
                )
            .MergeWorksheets()
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 2],
            ["C", "C", 3, 3],
            ["D", "D", 4, 4],
            ["E", "E", 5, 8],
            ["F", "F", 6, 6],
            ], "Merged");
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 3, 6, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_MultipleDocuments_WithMerging()
    {
        // Arrange
        using ExcelPackage oldExcelPackage1 = ExcelTestHelper.ConvertToExcelPackage(dataTab1);
        using ExcelPackage oldExcelPackage2 = ExcelTestHelper.ConvertToExcelPackage(dataTab2);
        using var oldFileStream1 = oldExcelPackage1.ToMemoryStream();
        using var oldFileStream2 = oldExcelPackage2.ToMemoryStream();

        using ExcelPackage newExcelPackage1 = ExcelTestHelper.ConvertToExcelPackage(dataTab1);
        using ExcelPackage newExcelPackage2 = ExcelTestHelper.ConvertToExcelPackage(dataTab2);
        newExcelPackage2.Workbook.Worksheets[0].Cells[3, 2].Value = 8;
        using var newFileStream1 = newExcelPackage1.ToMemoryStream();
        using var newFileStream2 = newExcelPackage2.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream1, "OldFile1.xlsx")
                .SetNewFile(newFileStream1, "NewFile1.xlsx")
                )
            .AddFiles(x => x
                .SetOldFile(oldFileStream2, "OldFile2.xlsx")
                .SetNewFile(newFileStream2, "NewFile2.xlsx")
                )
            .MergeDocuments()
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 2],
            ["C", "C", 3, 3],
            ["D", "D", 4, 4],
            ["E", "E", 5, 8],
            ["F", "F", 6, 6],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[6, 3, 6, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

    [Test]
    public void Diff_MultipleDocuments_WithoutMerging()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(dataTab1);
        using var oldFileStream1 = oldExcelPackage.ToMemoryStream();
        using var oldFileStream2 = oldExcelPackage.ToMemoryStream();
        using var newFileStream1 = oldExcelPackage.ToMemoryStream();
        using var newFileStream2 = oldExcelPackage.ToMemoryStream();

        // Act
        ExcelDiffBuilder builder = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream1, "OldFile1.xlsx")
                .SetNewFile(newFileStream1, "NewFile1.xlsx")
                )
            .AddFiles(x => x
                .SetOldFile(oldFileStream2, "OldFile2.xlsx")
                .SetNewFile(newFileStream2, "NewFile2.xlsx")
                );

        // Assert
        Assert.Throws<InvalidOperationException>(() =>
        {
            builder.Build();
        });
    }

}
