namespace XlsxDiffEngineTests;

internal class ExcelDiffBuilderAdditionalColumnsTests
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
    public void Diff_WithRowNumber()
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
            .AddRowNumberAsColumn("Row")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Row", "Row", "Title", "Title", "Value", "Value"],
            [1, 1, "A", "A", 1, 1],
            [2, 2, "B", "B", 2, 4],
            [3, 3, "C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithAdditionalMetadataColumns()
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
            .AddRowNumberAsColumn("Row")
            .AddWorksheetNameAsColumn("Worksheet")
            .AddDocumentNameAsColumn("Document")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Row", "Row", "Worksheet", "Worksheet", "Document", "Document", "Title", "Title", "Value", "Value"],
            [1, 1, "Table", "Table", "OldFile.xlsx", "NewFile.xlsx", "A", "A", 1, 1],
            [2, 2, "Table", "Table", "OldFile.xlsx", "NewFile.xlsx", "B", "B", 2, 4],
            [3, 3, "Table", "Table", "OldFile.xlsx", "NewFile.xlsx", "C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithCustomDocumentNameColumn()
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
                .SetDocumentName("ChangedDocumentName")
                )
            .AddDocumentNameAsColumn("Document")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Document", "Document", "Title", "Title", "Value", "Value"],
            ["ChangedDocumentName", "ChangedDocumentName", "A", "A", 1, 1],
            ["ChangedDocumentName", "ChangedDocumentName", "B", "B", 2, 4],
            ["ChangedDocumentName", "ChangedDocumentName", "C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithMergedWorksheetNameColumn()
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
                .SetMergedWorksheetName("Test")
                )
            .MergeWorksheets()
            .AddMergedWorksheetNameAsColumn("MergedWorksheet")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["MergedWorksheet", "MergedWorksheet", "Title", "Title", "Value", "Value"],
            ["Test", "Test", "A", "A", 1, 1],
            ["Test", "Test", "B", "B", 2, 4],
            ["Test", "Test", "C", "C", 3, 3],
            ], "Test");

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }
}
