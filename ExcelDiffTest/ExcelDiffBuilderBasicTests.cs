using System.Drawing;

namespace ExcelDiffTest;

internal class ExcelDiffBuilderBasicTests
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
    public void Diff_WithRecalculation()
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
    public void Diff_WithoutUnchangedRows()
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
            .IgnoreUnchangedRows()
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["B", "B", 2, 4],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithColumnHeaderPostfix()
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
            .SetNewHeaderColumnPostfix("New")
            .SetOldHeaderColumnPostfix("Old")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["TitleOld", "TitleNew", "ValueOld", "ValueNew"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithOmittedColumn()
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
            .SetColumnsToOmit("Title")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Value", "Value"],
            [1, 1],
            [2, 4],
            [3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }


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
    public void Diff_WithAdditionColumns()
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
    public void Diff_WithChangedDocumentName()
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
    public void Diff_WithMergedWorksheetName()
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

    [Test]
    public void Diff_WithNumberFormatCopyCellFormatTrue()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellFormat(true)
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        expectedResult.Workbook.Worksheets[0].Cells[2, 3, 4, 4].Style.Numberformat.Format = "0.00";

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithNumberFormatAndCopyCellFormatFalse()
    {
        // Arrange
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        oldExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        newExcelPackage.Workbook.Worksheets[0].Cells[2, 2, 4, 2].Style.Numberformat.Format = "0.00";
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellFormat(false)
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithStyleAndCopyCellStyleTrue()
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
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellStyle(true)
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 1, 4, 4], cellStyle);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithStyleAndCopyCellStyleFalse()
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
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFileContent);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFileContent);
        ExcelHelper.SetCellStyle(oldExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        ExcelHelper.SetCellStyle(newExcelPackage.Workbook.Worksheets[0].Cells[2, 1, 4, 2], cellStyle);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();

        // Act
        using ExcelPackage result = excelDiffBuilder
            .AddFiles(x => x
                .SetOldFile(oldFileStream, "OldFile.xlsx")
                .SetNewFile(newFileStream, "NewFile.xlsx")
                )
            .CopyCellStyle(false)
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["A", "A", 1, 1],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult);
    }

    [Test]
    public void Diff_WithSkippedRows()
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
            .SetSkipRowRule((dataSource, row) => dataSource.GetCellText("Title", row) == "A")
            .Build();

        // Assert
        using ExcelPackage expectedResult = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Title", "Value", "Value"],
            ["B", "B", 2, 4],
            ["C", "C", 3, 3],
            ]);
        ExcelHelper.SetCellStyle(expectedResult.Workbook.Worksheets[0].Cells[2, 3, 2, 4], DefaultCellStyles.ChangedCell);

        ExcelTestHelper.CheckIfExcelPackagesIdentical(result, expectedResult, true);
    }

}
