using OfficeOpenXml;
using System.Globalization;

namespace XlsxDiffEngine;

internal sealed class ExcelDataSource : IExcelDataSource
{
    private readonly ExcelWorksheet worksheet;
    private readonly ExcelDataSourceConfig config;
    private readonly ExcelAddress section;
    private readonly List<string> columnNames = [];
    private readonly Dictionary<string, int> columnDict;
    private readonly int dataRowsOffset;

    public string Name { get; }
    public int DataRows { get; }

    public ExcelDataSource(ExcelWorksheet worksheet, ExcelDataSourceConfig? config = null, ExcelAddress? section = null)
    {
        this.worksheet = worksheet;
        this.section = section ?? worksheet.Dimension;
        this.config = config ?? new ExcelDataSourceConfig();
        Name = worksheet.Name;
        DataRows = this.section.Rows - 1;
        columnDict = new(this.config.StringComparer);
        dataRowsOffset = this.section.Start.Row;
    }

    public IReadOnlyCollection<string> GetColumnNames()
    {
        if (columnNames.Count > 0)
        {
            return columnNames;
        }
        if (config.RowNumberColumnName is not null)
        {
            columnNames.Add(config.RowNumberColumnName);
            columnDict.Add(config.RowNumberColumnName, -3);
        }
        if (config.WorksheetNameColumnName is not null)
        {
            columnNames.Add(config.WorksheetNameColumnName);
            columnDict.Add(config.WorksheetNameColumnName, -2);
        }
        if (config.CustomColumnName is not null)
        {
            columnNames.Add(config.CustomColumnName);
            columnDict.Add(config.CustomColumnName, -1);
        }
        HashSet<string> columnsToIgnore = new(config.ColumnsToIgnore, config.StringComparer);
        for (int columnIndex = section.Start.Column; columnIndex <= section.End.Column; columnIndex++)
        {
            string columnName = worksheet.Cells[section.Start.Row, columnIndex].Text;
            if (columnDict.ContainsKey(columnName)) { continue; }
            if (columnsToIgnore.Contains(columnName)) { continue; }
            columnNames.Add(columnName);
            columnDict.Add(columnNames.Last(), columnIndex);
        }
        return columnNames;
    }

    public ExcelRange? GetExcelRange(string columnName, int row)
    {
        if (row < 1 || row > DataRows) { return null; }
        if (columnDict.TryGetValue(columnName, out int column))
        {
            return column < 0 ? null : worksheet.Cells[dataRowsOffset + row, column];
        }
        return null;
    }

    public object? GetCellValue(string columnName, int row)
    {
        if (row < 1 || row > DataRows) { return null; }
        if (columnDict.TryGetValue(columnName, out int column))
        {
            if (column == -3)
            {
                return row;
            }
            else if (column == -2)
            {
                return Name;
            }
            else if (column == -1)
            {
                return config.CustomColumnValue;
            }
            return worksheet.Cells[dataRowsOffset + row, column].Value;
        }
        return null;
    }

    public string GetCellText(string columnName, int row)
    {
        if (row < 1 || row > DataRows) { return ""; }
        if (columnDict.TryGetValue(columnName, out int column))
        {
            if (column == -3)
            {
                return row.ToString(CultureInfo.InvariantCulture);
            }
            else if (column == -2)
            {
                return Name;
            }
            else if (column == -1)
            {
                return config.CustomColumnValue?.ToString() ?? "";
            }
            return worksheet.Cells[dataRowsOffset + row, column].Text;
        }
        return "";
    }

    public IReadOnlyDictionary<string, object?> GetRow(int row)
    {
        Dictionary<string, object?> rowValues = [];
        foreach (string columnName in GetColumnNames())
        {
            rowValues[columnName] = GetCellValue(columnName, row);
        }
        return rowValues;
    }

    public object?[] GetColumn(string columnName)
    {
        if (columnDict.TryGetValue(columnName, out int column))
        {
            if (column == -3)
            {
                return Enumerable.Range(1, DataRows).OfType<object?>().ToArray();
            }
            else if (column == -2)
            {
                return Enumerable.Range(1, DataRows).Select(x => (object?)Name).ToArray();
            }
            else if (column == -1)
            {
                return Enumerable.Range(1, DataRows).Select(x => config.CustomColumnValue).ToArray();
            }
            object?[] cellValues = worksheet.Cells[section.Start.Row + 1, column, section.End.Row, column].GetValue<object?[]>();
            return cellValues;
        }
        return Enumerable.Range(1, DataRows).Select(x => (object?)null).ToArray(); ;
    }

}
