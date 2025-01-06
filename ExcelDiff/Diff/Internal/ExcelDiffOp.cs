using System.Data;
using System.Globalization;
using System.Text;

namespace ExcelDiffEngine;

internal sealed class ExcelDiffOp
{
    private readonly IExcelDataSource oldDataSource;
    private readonly IExcelDataSource newDataSource;
    private readonly ExcelDiffConfig config;
    private readonly StringComparer stringComparer;
    private readonly StringBuilder stringBuilder = new();

    public IReadOnlyList<string> MergedColumnNames { get; }

    public ExcelDiffOp(IExcelDataSource oldDataSource, IExcelDataSource newDataSource, ExcelDiffConfig config)
    {
        this.oldDataSource = oldDataSource;
        this.newDataSource = newDataSource;
        this.config = config;
        stringComparer = config.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        MergedColumnNames = newDataSource.GetColumnNames().Union(oldDataSource.GetColumnNames(), stringComparer).ToList().AsReadOnly();
    }

    public List<(int? oldRow, int? newRow)> GetMergedRows()
    {
        List<DataKey> oldDataKeys = GetDataKeys(oldDataSource);
        List<DataKey> newDataKeys = GetDataKeys(newDataSource);
        var oldKeyDict = oldDataKeys.ToDictionary(x => x.PrimaryKey, stringComparer);
        var oldSecondaryKeyDict = oldDataKeys.ToDictionary(x => x.SecondaryKey, stringComparer);
        var newKeyDict = newDataKeys.ToDictionary(x => x.PrimaryKey, stringComparer);
        var usedDataKeys = newDataKeys.Select(x => x.PrimaryKey).ToHashSet(stringComparer);
        List<(int? oldRow, int? newRow)> diff = [];
        foreach (string dataKey in GetCombinedKeyList(oldDataKeys, newDataKeys))
        {
            _ = oldKeyDict.TryGetValue(dataKey, out DataKey? oldRow);
            _ = newKeyDict.TryGetValue(dataKey, out DataKey? newRow);
            if (newRow is null)
            {
                if (oldRow is not null && !usedDataKeys.Contains(dataKey))
                {
                    diff.Add((oldRow.RowNumber, null));
                }
            }
            else if (oldRow is not null)
            {
                diff.Add((oldRow.RowNumber, newRow.RowNumber));
            }
            else if (!usedDataKeys.Contains(dataKey) && oldSecondaryKeyDict.TryGetValue(newRow.SecondaryKey, out oldRow))
            {
                diff.Add((oldRow.RowNumber, newRow.RowNumber));
                _ = usedDataKeys.Add(oldRow.PrimaryKey);
            }
            else
            {
                diff.Add((null, newRow.RowNumber));
            }
        }
        return diff;
    }

    private static List<string> GetCombinedKeyList(List<DataKey> oldDataKeys, List<DataKey> newDataKeys)
    {
        var groupKeys = newDataKeys
            .Select(item => item.GroupKey)
            .Union(oldDataKeys.Select(item => item.GroupKey))
            .ToList();
        var combinedKeyList = groupKeys
            .SelectMany(group => newDataKeys
                .Where(item => item.GroupKey == group)
                .Select(x => x.PrimaryKey)
                .Union(oldDataKeys.Where(item => item.GroupKey == group).Select(x => x.PrimaryKey)))
            .Distinct()
            .ToList();
        return combinedKeyList;
    }

    private List<DataKey> GetDataKeys(IExcelDataSource dataSource)
    {
        List<DataKey> dataKeys = [];
        HashSet<string> primaryKeySet = [];
        for (int row = 1; row <= dataSource.DataRows; row++)
        {
            if (config.SkipRowRule is not null && config.SkipRowRule(dataSource, row)) { continue; }
            var dataKey = new DataKey(
                GetKey(dataSource, row, config.KeyColumns),
                GetKey(dataSource, row, config.SecondaryKeyColumns),
                config.GroupKeyColumns.Count > 0 ? GetKey(dataSource, row, config.GroupKeyColumns) : "",
                row);
            if (primaryKeySet.Contains(dataKey.PrimaryKey))
            {
                int i = 1;
                for (; primaryKeySet.Contains($"@+{i}{dataKey.PrimaryKey}"); i++) { }
                dataKey = dataKey with { PrimaryKey = $"@+{i}{dataKey.PrimaryKey}" };
            }
            primaryKeySet.Add(dataKey.PrimaryKey);
            dataKeys.Add(dataKey);
        }
        return dataKeys;
    }

    private string GetKey(IExcelDataSource dataSource, int row, IReadOnlyCollection<string> keyColumnNames)
    {
        if (keyColumnNames.Count == 0) { return row.ToString(CultureInfo.InvariantCulture); }
        if (keyColumnNames.Count == 1) { return dataSource.GetCellText(keyColumnNames.First(), row); }
        _ = stringBuilder.Clear();
        foreach (string columnName in keyColumnNames)
        {
            _ = stringBuilder
                .Append('@')
                .Append(columnName)
                .Append(':')
                .Append(dataSource.GetCellText(columnName, row));
        }
        return stringBuilder.ToString();
    }

    private sealed record DataKey(string PrimaryKey, string SecondaryKey, string GroupKey, int RowNumber);
}
