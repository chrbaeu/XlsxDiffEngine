using OfficeOpenXml;

namespace ExcelDiffEngine;

public interface IExcelDataSource
{
    public string Name { get; }
    public int DataRows { get; }
    public List<string> GetColumnNames();
    public ExcelRange? GetExcelRange(string columnName, int row);
    public object? GetCellValue(string columnName, int row);
    public string GetCellText(string columnName, int row);
    public Dictionary<string, object?> GetRow(int row);
    public object?[] GetColumn(string columnName);
}
