using OfficeOpenXml;
using System.Globalization;

namespace XlsxDiffEngine;

internal sealed class MergedExcelDataSource : IExcelDataSource
{
    private readonly ICollection<IExcelDataSource> excelDataSources;
    private readonly ExcelDataSourceConfig config;
    private readonly List<string> columnNames = [];
    private readonly Dictionary<string, int> columnDict;
    public string Name { get; }
    public int DataRows { get; }

    public MergedExcelDataSource(string name, ICollection<IExcelDataSource> excelDataSources, ExcelDataSourceConfig config)
    {
        this.excelDataSources = excelDataSources;
        this.config = config ?? new ExcelDataSourceConfig();
        Name = name;
        DataRows = excelDataSources.Sum(x => x.DataRows);
        columnDict = new(this.config.StringComparer);
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
        if (config.MergedWorksheetNameColumnName is not null)
        {
            columnNames.Add(config.MergedWorksheetNameColumnName);
            columnDict.Add(config.MergedWorksheetNameColumnName, -2);
        }
        if (config.CustomColumnName is not null)
        {
            columnNames.Add(config.CustomColumnName);
            columnDict.Add(config.CustomColumnName, -1);
        }
        HashSet<string> columnsToIgnore = new(config.ColumnsToIgnore, config.StringComparer);
        foreach (string columnName in excelDataSources.SelectMany(x => x.GetColumnNames()))
        {
            if (columnDict.ContainsKey(columnName)) { continue; }
            if (columnsToIgnore.Contains(columnName)) { continue; }
            columnNames.Add(columnName);
            columnDict.Add(columnName, 0);
        }
        return columnNames;
    }

    public ExcelRange? GetExcelRange(string columnName, int row)
    {
        if (row < 1 || row > DataRows) { return null; }
        if (columnDict.TryGetValue(columnName, out int column))
        {
            if (column < 0)
            {
                return null;
            }
            foreach (IExcelDataSource dataSource in excelDataSources)
            {
                if (row <= dataSource.DataRows)
                {
                    return dataSource.GetExcelRange(columnName, row);
                }
                row -= dataSource.DataRows;
            }
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
            foreach (IExcelDataSource dataSource in excelDataSources)
            {
                if (row <= dataSource.DataRows)
                {
                    return dataSource.GetCellValue(columnName, row);
                }
                row -= dataSource.DataRows;
            }
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
            foreach (IExcelDataSource dataSource in excelDataSources)
            {
                if (row <= dataSource.DataRows)
                {
                    return dataSource.GetCellText(columnName, row);
                }
                row -= dataSource.DataRows;
            }
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
            object?[] cellValues = excelDataSources.SelectMany(x => x.GetColumn(columnName)).ToArray();
            return cellValues;
        }
        return Enumerable.Range(1, DataRows).Select(x => (object?)null).ToArray(); ;
    }
}