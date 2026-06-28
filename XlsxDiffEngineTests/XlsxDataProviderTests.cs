using TUnit.Assertions.Enums;

namespace XlsxDiffEngineTests;

internal class XlsxDataProviderTests
{
    [Test]
    public async Task GetDataSources_SingleWorksheet_ReturnsSingleDataSource()
    {
        using ExcelPackage excelPackage = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Value"],
            ["A", 1],
            ["B", 2],
        ]);
        using XlsxDataProvider provider = CreateProvider(excelPackage);

        IReadOnlyList<IExcelDataSource> dataSources = provider.GetDataSources();
        string[] columnNames = [.. dataSources[0].GetColumnNames()];

        await Assert.That(dataSources.Count).IsEqualTo(1);
        await Assert.That(dataSources[0].Name).IsEqualTo("Table");
        await Assert.That(dataSources[0].DataRows).IsEqualTo(2);
        await Assert.That(columnNames).IsEquivalentTo(["Title", "Value"], CollectionOrdering.Matching);
        await Assert.That(dataSources[0].GetCellValue("Title", 1)).IsEqualTo("A");
        await Assert.That(Convert.ToDecimal(dataSources[0].GetCellValue("Value", 2))).IsEqualTo(2m);
    }

    [Test]
    public async Task GetDataSources_WithWorksheetNameColumn_ExposesWorksheetNames()
    {
        using ExcelPackage excelPackage = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Value"],
            ["A", 1],
            ["B", 2],
        ], "Products");
        using XlsxDataProvider provider = CreateProvider(excelPackage, new XlsxDataProviderConfig
        {
            WorksheetNameColumnName = "Worksheet"
        });

        IExcelDataSource dataSource = provider.GetDataSources()[0];
        string[] columnNames = [.. dataSource.GetColumnNames()];
        object?[] worksheetColumn = dataSource.GetColumn("Worksheet");

        await Assert.That(dataSource.Name).IsEqualTo("Products");
        await Assert.That(dataSource.DataRows).IsEqualTo(2);
        await Assert.That(columnNames).IsEquivalentTo(["Worksheet", "Title", "Value"], CollectionOrdering.Matching);
        await Assert.That(dataSource.GetCellValue("Worksheet", 1)).IsEqualTo("Products");
        await Assert.That(dataSource.GetCellValue("Worksheet", 2)).IsEqualTo("Products");
        await Assert.That(dataSource.GetCellText("Worksheet", 1)).IsEqualTo("Products");
        await Assert.That(worksheetColumn.Cast<string>()).IsEquivalentTo(["Products", "Products"], CollectionOrdering.Matching);
    }

    [Test]
    public async Task GetDataSources_WithWorksheetSection_UsesWorksheetSpecificRange()
    {
        using ExcelPackage excelPackage = ExcelTestHelper.ConvertToExcelPackage([
            ["ignore", "ignore", null],
            ["Meta", "Meta", null],
            ["Key", "Amount", null],
            ["A", 10, null],
            ["B", 20, null],
        ]);
        using XlsxDataProvider provider = CreateProvider(excelPackage, fileInfo: new XlsxFileInfo(excelPackage.ToMemoryStream())
        {
            WorksheetInfos = [new XlsxWorksheetInfo { Name = "Table", FromRow = 3, FromColumn = 1, ToRow = 5, ToColumn = 2 }]
        });

        IExcelDataSource dataSource = provider.GetDataSources()[0];
        string[] columnNames = [.. dataSource.GetColumnNames()];
        object?[] amountColumn = dataSource.GetColumn("Amount");

        await Assert.That(dataSource.Name).IsEqualTo("Table");
        await Assert.That(dataSource.DataRows).IsEqualTo(2);
        await Assert.That(columnNames).IsEquivalentTo(["Key", "Amount"], CollectionOrdering.Matching);
        await Assert.That(dataSource.GetCellValue("Key", 1)).IsEqualTo("A");
        await Assert.That(Convert.ToDecimal(dataSource.GetCellValue("Amount", 2))).IsEqualTo(20m);
        await Assert.That(amountColumn.Select(Convert.ToDecimal)).IsEquivalentTo([10m, 20m], CollectionOrdering.Matching);
        await Assert.That(dataSource.GetColumnNames().Contains("ignore")).IsFalse();
        await Assert.That(dataSource.GetColumnNames().Contains("Meta")).IsFalse();
    }

    [Test]
    public async Task GetDataSources_MergeWorksheets_ReturnsMergedWorksheet()
    {
        using ExcelPackage excelPackage = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Value"],
            ["A", 1],
            ["B", 2],
        ], "Tab1");
        excelPackage.AddWorksheet([
            ["Title", "Value"],
            ["C", 3],
            ["D", 4],
        ], "Tab2");
        using XlsxDataProvider provider = CreateProvider(excelPackage, new XlsxDataProviderConfig
        {
            MergeWorksheets = true
        }, new XlsxFileInfo(excelPackage.ToMemoryStream(), "Merged.xlsx")
        {
            MergedWorksheetName = "Merged"
        });

        IReadOnlyList<IExcelDataSource> firstResult = provider.GetDataSources();
        IReadOnlyList<IExcelDataSource> secondResult = provider.GetDataSources();
        object?[] titleColumn = firstResult[0].GetColumn("Title");
        object?[] valueColumn = firstResult[0].GetColumn("Value");

        await Assert.That(firstResult.Count).IsEqualTo(1);
        await Assert.That(secondResult.Count).IsEqualTo(1);
        await Assert.That(ReferenceEquals(firstResult[0], secondResult[0])).IsTrue();
        await Assert.That(firstResult[0].Name).IsEqualTo("Merged");
        await Assert.That(firstResult[0].DataRows).IsEqualTo(4);
        await Assert.That(firstResult[0].GetColumnNames()).IsEquivalentTo(["Title", "Value"], CollectionOrdering.Matching);
        await Assert.That(titleColumn.Cast<string>()).IsEquivalentTo(["A", "B", "C", "D"], CollectionOrdering.Matching);
        await Assert.That(valueColumn.Select(Convert.ToDecimal)).IsEquivalentTo([1m, 2m, 3m, 4m], CollectionOrdering.Matching);
    }

    [Test]
    public async Task GetDataSources_MergeDocumentsAndWorksheets_ReturnsSingleMergedDocumentDataSource()
    {
        using ExcelPackage excelPackage1 = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Value"],
            ["A", 1],
            ["B", 2],
        ], "Tab1");
        using ExcelPackage excelPackage2 = ExcelTestHelper.ConvertToExcelPackage([
            ["Title", "Value"],
            ["C", 3],
            ["D", 4],
        ], "Tab1");

        using XlsxDataProvider provider = new([
            new XlsxFileInfo(excelPackage1.ToMemoryStream(), "Document1.xlsx"),
            new XlsxFileInfo(excelPackage2.ToMemoryStream(), "Document2.xlsx")
        ], new XlsxDataProviderConfig
        {
            MergeDocuments = true,
            MergeWorksheets = true,
            MergedDocumentName = "AllData"
        });

        IReadOnlyList<IExcelDataSource> firstResult = provider.GetDataSources();
        IReadOnlyList<IExcelDataSource> secondResult = provider.GetDataSources();
        object?[] titleColumn = firstResult[0].GetColumn("Title");
        object?[] valueColumn = firstResult[0].GetColumn("Value");

        await Assert.That(firstResult.Count).IsEqualTo(1);
        await Assert.That(secondResult.Count).IsEqualTo(1);
        await Assert.That(ReferenceEquals(firstResult[0], secondResult[0])).IsTrue();
        await Assert.That(firstResult[0].Name).IsEqualTo("AllData");
        await Assert.That(firstResult[0].DataRows).IsEqualTo(4);
        await Assert.That(firstResult[0].GetColumnNames()).IsEquivalentTo(["Title", "Value"], CollectionOrdering.Matching);
        await Assert.That(titleColumn.Cast<string>()).IsEquivalentTo(["A", "B", "C", "D"], CollectionOrdering.Matching);
        await Assert.That(valueColumn.Select(Convert.ToDecimal)).IsEquivalentTo([1m, 2m, 3m, 4m], CollectionOrdering.Matching);
    }

    [Test]
    public async Task GetDataSources_SingleWorksheetWithTwoNamedSections_ReturnsTwoDataSources()
    {
        using ExcelPackage excelPackage = ExcelTestHelper.ConvertToExcelPackage([
            ["Id", "Value", null],
            [1, "A", null],
            [2, "B", null],
            [null, null, null],
            [null, "Code", "Amount"],
            [null, "X", 10],
            [null, "Y", 20],
        ], "Sheet1");
        using XlsxDataProvider provider = CreateProvider(excelPackage, fileInfo: new XlsxFileInfo(excelPackage.ToMemoryStream(), "Sections.xlsx")
        {
            WorksheetInfos =
            [
                new XlsxWorksheetInfo { Name = "Sheet1", FromRow = 1, FromColumn = 1, ToRow = 3, ToColumn = 2, AlternativeName = "Section1" },
                new XlsxWorksheetInfo { Name = "Sheet1", FromRow = 5, FromColumn = 2, ToRow = 7, ToColumn = 3, AlternativeName = "Section2" }
            ]
        });

        IReadOnlyList<IExcelDataSource> dataSources = provider.GetDataSources();
        IExcelDataSource section1 = dataSources.Single(x => x.Name == "Section1");
        IExcelDataSource section2 = dataSources.Single(x => x.Name == "Section2");
        string[] section1Columns = [.. section1.GetColumnNames()];
        string[] section2Columns = [.. section2.GetColumnNames()];
        object?[] section1Ids = section1.GetColumn("Id");
        object?[] section1Values = section1.GetColumn("Value");
        object?[] section2Codes = section2.GetColumn("Code");
        object?[] section2Amounts = section2.GetColumn("Amount");

        await Assert.That(dataSources.Count).IsEqualTo(2);
        await Assert.That(section1.Name).IsEqualTo("Section1");
        await Assert.That(section1.DataRows).IsEqualTo(2);
        await Assert.That(section1Columns).IsEquivalentTo(["Id", "Value"], CollectionOrdering.Matching);
        await Assert.That(section1Ids.Select(Convert.ToDecimal)).IsEquivalentTo([1m, 2m], CollectionOrdering.Matching);
        await Assert.That(section1Values.Cast<string>()).IsEquivalentTo(["A", "B"], CollectionOrdering.Matching);
        await Assert.That(section2.Name).IsEqualTo("Section2");
        await Assert.That(section2.DataRows).IsEqualTo(2);
        await Assert.That(section2Columns).IsEquivalentTo(["Code", "Amount"], CollectionOrdering.Matching);
        await Assert.That(section2Codes.Cast<string>()).IsEquivalentTo(["X", "Y"], CollectionOrdering.Matching);
        await Assert.That(section2Amounts.Select(Convert.ToDecimal)).IsEquivalentTo([10m, 20m], CollectionOrdering.Matching);
    }

    private static XlsxDataProvider CreateProvider(ExcelPackage excelPackage, XlsxDataProviderConfig? config = null, XlsxFileInfo? fileInfo = null)
    {
        return new(fileInfo ?? new XlsxFileInfo(excelPackage.ToMemoryStream(), "Test.xlsx"), config);
    }
}
