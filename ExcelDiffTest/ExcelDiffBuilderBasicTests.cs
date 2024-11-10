using ExcelDiffEngine;
using System.Drawing;

namespace ExcelDiffTest;

internal class ExcelDiffBuilderBasicTests
{
    private readonly ExcelDiffBuilder excelDiffBuilder = new();

    private readonly object[][] oldFileContent = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 2],
        ["C", 3]
    ];

    private readonly object[][] newFileContent = [
        ["Title", "Value"],
        ["A", 1],
        ["B", 4],
        ["C", 3]
    ];

    [Test]
    public async Task Diff_WithHighlighting()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithHighlightingAndKeyColumn()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetKeyColumns("Title")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 1, 3, 2], DefaultCellStyles.ChangedRowKeyColumns);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithRecalculation()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2].Formula = "=10-9";
        newExcelPackage.Workbook.Worksheets[0].Cells[3, 2].Value = null;
        newExcelPackage.Workbook.Worksheets[0].Cells[3, 2].Formula = "=10-6";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                .RecalculateFormulas()
                )
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[3, 3, 3, 4], DefaultCellStyles.ChangedCell);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithHighlightingAndKeyColumnAndInsertAndDelete()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[4, 1].Value = "D";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetKeyColumns("Title")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
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
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

    [Test]
    public async Task Diff_WithoutUnchangedRows()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .IgnoreUnchangedRows()
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["B", "B", 2, 4],
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithColumnHeaderPostfix()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetNewHeaderColumnPostfix("New")
            .SetOldHeaderColumnPostfix("Old")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["TitleOld", "TitleNew", "ValueOld", "ValueNew"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithOmittedColumn()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetColumnsToOmit("Title")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Value", "Value"],
            [1, 1],
            [2, 4],
            [3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }


    [Test]
    public async Task Diff_WithRowNumber()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddRowNumberAsColumn("Row")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Row", "Row", "Title", "Title", "Value", "Value"],
            [1, 1, "A", "A", 1, 1],
            [2, 2, "B", "B", 2, 4],
            [3, 3, "C", "C", 3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithAdditionColumns()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .AddRowNumberAsColumn("Row")
            .AddWorksheetNameAsColumn("Worksheet")
            .AddDocumentNameAsColumn("Document")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Row", "Row", "Worksheet", "Worksheet", "Document", "Document", "Title", "Title", "Value", "Value"],
            [1, 1, "Table", "Table", "OldFile.xlsx", "NewFile.xlsx", "A", "A", 1, 1],
            [2, 2, "Table", "Table", "OldFile.xlsx", "NewFile.xlsx", "B", "B", 2, 4],
            [3, 3, "Table", "Table", "OldFile.xlsx", "NewFile.xlsx", "C", "C", 3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithChangedDocumentName()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                .SetDocumentName("ChangedDocumentName")
                )
            .AddDocumentNameAsColumn("Document")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Document", "Document", "Title", "Title", "Value", "Value"],
            ["ChangedDocumentName", "ChangedDocumentName", "A", "A", 1, 1],
            ["ChangedDocumentName", "ChangedDocumentName", "B", "B", 2, 4],
            ["ChangedDocumentName", "ChangedDocumentName", "C", "C", 3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithMergedWorksheetName()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                .SetMergedWorksheetName("Test")
                )
            .MergeWorkSheets()
            .AddMergedWorksheetNameAsColumn("MergedWorksheet")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["MergedWorksheet", "MergedWorksheet", "Title", "Title", "Value", "Value"],
            ["Test", "Test", "A", "A", 1, 1],
            ["Test", "Test", "B", "B", 2, 4],
            ["Test", "Test", "C", "C", 3, 3]
            ], "Test");
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithNumberFormatCopyCellFormatTrue()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellFormat(true)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        expectedResult.Workbook.Worksheets[0].Cells[2, 3, 4, 4].Style.Numberformat.Format = "0.00";
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithNumberFormatAndCopyCellFormatFalse()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellFormat(false)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithStyleAndCopyCellStyleTrue()
    {
        // Arrange
        CellStyle cellStyle = new()
        {
            Bold = true,
            Italic = true,
            Underline = true,
            FontColor = Color.Red,
            BackgroundColor = Color.Blue
        };
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellStyle(true)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 4], cellStyle);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithStyleAndCopyCellStyleFalse()
    {
        // Arrange
        CellStyle cellStyle = new()
        {
            Bold = true,
            Italic = true,
            Underline = true,
            FontColor = Color.Red,
            BackgroundColor = Color.Blue
        };
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellStyle(false)
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult)).IsTrue();
    }

    [Test]
    public async Task Diff_WithSkippedRows()
    {
        // Arrange
        using var oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using var newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        var result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .SetSkipRowRule((dataSource, row) => dataSource.GetCellText("Title", row) == "A")
            .Build();

        // Assert
        var expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3]
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 3, 2, 4], DefaultCellStyles.ChangedCell);
        await Assert.That(ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true)).IsTrue();
    }

}
