namespace XlsxDiffEngineTests;

internal class ExcelDiffWriterTests
{
    [Test]
    public async Task WriteDiff_WithOffsetStartColumn_HidesOnlyOldDataColumnsRelativeToOffset()
    {
        // Arrange
        object[][] oldFile = [
            ["Title", "Value"],
            ["A", 1],
        ];
        object[][] newFile = [
            ["Title", "Value"],
            ["A", 2],
        ];
        using ExcelPackage oldExcelPackage = ExcelTestHelper.ConvertToExcelPackage(oldFile);
        using ExcelPackage newExcelPackage = ExcelTestHelper.ConvertToExcelPackage(newFile);
        using var oldFileStream = oldExcelPackage.ToMemoryStream();
        using var newFileStream = newExcelPackage.ToMemoryStream();
        using XlsxDataProvider oldDataProvider = new(new XlsxFileInfo(oldFileStream, "OldFile.xlsx"));
        using XlsxDataProvider newDataProvider = new(new XlsxFileInfo(newFileStream, "NewFile.xlsx"));
        IExcelDataSource oldDataSource = oldDataProvider.GetDataSources()[0];
        IExcelDataSource newDataSource = newDataProvider.GetDataSources()[0];
        ExcelDiffConfig config = new() { HideOldColumns = true };
        ExcelDiffWriter writer = new(oldDataSource, newDataSource, config);
        using ExcelPackage result = new();
        ExcelWorksheet worksheet = result.Workbook.Worksheets.Add("Result");

        // Act
        _ = writer.WriteDiff(worksheet, 1, 3);

        // Assert
        await Assert.That(worksheet.Column(1).Hidden).IsFalse();
        await Assert.That(worksheet.Column(2).Hidden).IsFalse();
        await Assert.That(worksheet.Column(3).Hidden).IsTrue();
        await Assert.That(worksheet.Column(4).Hidden).IsFalse();
        await Assert.That(worksheet.Column(5).Hidden).IsTrue();
        await Assert.That(worksheet.Column(6).Hidden).IsFalse();
    }
}
