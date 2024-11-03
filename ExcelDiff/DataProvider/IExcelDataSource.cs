using OfficeOpenXml;

namespace ExcelDiffEngine;

public interface IExcelDataSource
{
    public string Name { get; }
    public int DataRows { get; }
    public IReadOnlyCollection<string> GetColumnNames();
    public ExcelRange? GetExcelRange(string columnName, int row);
    public object? GetCellValue(string columnName, int row);
    public string GetCellText(string columnName, int row);
    public IReadOnlyDictionary<string, object?> GetRow(int row);
    public object?[] GetColumn(string columnName);
}
