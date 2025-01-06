using OfficeOpenXml;

namespace ExcelDiffEngine;

/// <summary>
/// A class for writing the differences between two <see cref="IExcelDataSource"/> into an <see cref="ExcelWorksheet"/>,
/// including options for comparing, styling, and marking changes.
/// </summary>
public sealed class ExcelDiffWriter
{
    private readonly IExcelDataSource oldDataSource;
    private readonly IExcelDataSource newDataSource;
    private readonly ExcelDiffConfig config;
    private readonly StringComparer stringComparer;
    private readonly ExcelDiffOp excelDiffOp;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelDiffWriter"/> class with the specified data sources and configuration.
    /// </summary>
    /// <param name="oldDataSource">The data source representing the old state of the data.</param>
    /// <param name="newDataSource">The data source representing the new state of the data.</param>
    /// <param name="config">Configuration options for comparing and styling differences.</param>
    public ExcelDiffWriter(IExcelDataSource oldDataSource, IExcelDataSource newDataSource, ExcelDiffConfig config)
    {
        ArgumentNullThrowHelper.ThrowIfNull(oldDataSource, nameof(oldDataSource));
        ArgumentNullThrowHelper.ThrowIfNull(newDataSource, nameof(newDataSource));
        ArgumentNullThrowHelper.ThrowIfNull(config, nameof(config));
        this.oldDataSource = oldDataSource;
        this.newDataSource = newDataSource;
        this.config = config;
        stringComparer = config.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        excelDiffOp = new(oldDataSource, newDataSource, config);
    }

    /// <summary>
    /// Writes the differences between the old and new data sources into the specified worksheet, starting at the specified row and column.
    /// </summary>
    /// <param name="worksheet">The worksheet where differences will be written.</param>
    /// <param name="row">The starting row for writing data. Default is 1.</param>
    /// <param name="column">The starting column for writing data. Default is 1.</param>
    /// <returns>A tuple indicating the end row and column where data was written.</returns>
    public (int endRow, int endColumn) WriteDiff(ExcelWorksheet worksheet, int row = 1, int column = 1)
    {
        ArgumentNullThrowHelper.ThrowIfNull(worksheet, nameof(worksheet));
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfLessThan(row, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
#endif
        int headerEndColumns = WriteHeader(worksheet, row, column);
        int dataEndRow = WriteData(worksheet, row + 1, column);
        return (dataEndRow, headerEndColumns);
    }

    private int WriteData(ExcelWorksheet worksheet, int row, int startColumn)
    {
        ModificationRuleHandler ruleHandler = new(config.ModificationRules, config.IgnoreCase);
        foreach ((int? oldRow, int? newRow) in excelDiffOp.GetMergedRows())
        {
            int column = startColumn;
            bool isChanged = false;
            List<ExcelRange> keyCells = [];
            foreach (string columnName in excelDiffOp.MergedColumnNames)
            {
                if (config.ColumnsToOmit.Contains(columnName, stringComparer))
                {
                    continue;
                }
                ExcelRange? oldDstCell = worksheet.Cells[row, column];
                object? oldValue = SetCell(oldDstCell, columnName, oldRow, ruleHandler, DataKind.Old);
                string oldText = config.AddOldValueAsComment ? oldDstCell?.Text ?? "" : "";
                if (config.ShowOldDataColumn) { column++; } else { oldDstCell = null; }
                ExcelRange newDstCell = worksheet.Cells[row, column];
                object? newValue = SetCell(newDstCell, columnName, newRow, ruleHandler, DataKind.New);
                if (config.AddOldValueAsComment && oldValue?.ToString() != newValue?.ToString())
                {
                    string? comment = config.OldValueCommentPrefix is { } prefix ? prefix + oldText : oldText;
                    _ = newDstCell.AddComment(comment ?? "");
                }
                column++;
                isChanged |= GetAndHandleChangedState(columnName, oldDstCell, oldValue, newDstCell, newValue);
                if (config.KeyColumns.Contains(columnName, stringComparer))
                {
                    if (oldDstCell is not null) { keyCells.Add(oldDstCell); }
                    keyCells.Add(newDstCell);
                }
            }
            if (isChanged && config.ChangedRowKeyColumnsStyle is not null)
            {
                foreach (ExcelRange keyCell in keyCells)
                {
                    ExcelHelper.SetCellStyle(keyCell, config.ChangedRowKeyColumnsStyle);
                }
            }
            if (oldRow is null) { ExcelHelper.SetCellStyle(worksheet.Cells[row, startColumn, row, column - 1], config.AddedRowStyle); }
            if (newRow is null) { ExcelHelper.SetCellStyle(worksheet.Cells[row, startColumn, row, column - 1], config.RemovedRowStyle); }
            if (config.IgnoreUnchangedRows && !isChanged)
            {
                worksheet.Cells[row, startColumn, row, column - 1].Clear();
                continue;
            }
            row++;
        }
        return row - 1;
    }

    private bool GetAndHandleChangedState(string columnName, ExcelRange? oldDstCell, object? oldValue, ExcelRange newDstCell, object? newValue)
    {
        if (config.ColumnsToCompare is not null && !config.ColumnsToCompare.Contains(columnName, stringComparer))
        {
            return false;
        }
        if (config.ColumnsToIgnore is not null && config.ColumnsToIgnore.Contains(columnName, stringComparer))
        {
            return false;
        }
        CellStyle? cellStyle = null;
        if (config.ColumnsToTextCompareOnly.Contains(columnName, stringComparer))
        {
            if ((oldDstCell is null && oldValue?.ToString() != newValue?.ToString())
                || (oldDstCell is not null && oldDstCell.Text != newDstCell.Text))
            {
                cellStyle = config.ChangedCellStyle;
            }
        }
        else if (oldValue is double oldNumber && newValue is double newNumber
            && config.ValueChangedMarkers.Count > 0)
        {
            double pDiff = Math.Abs((oldNumber - newNumber) / ((oldNumber + newNumber) / 2.0));
            double aDiff = Math.Abs(oldNumber - newNumber);
            foreach (ValueChangedMarker valueChangedMarker in config.ValueChangedMarkers)
            {
                if (pDiff > valueChangedMarker.MinDeviationInPercent && aDiff > valueChangedMarker.MinDeviationAbsolute)
                {
                    cellStyle = valueChangedMarker.CellStyle;
                }
            }
        }
        else if (oldValue?.ToString() != newValue?.ToString())
        {
            cellStyle = config.ChangedCellStyle;
        }
        if (cellStyle is not null)
        {
            ExcelHelper.SetCellStyle(oldDstCell, cellStyle);
            ExcelHelper.SetCellStyle(newDstCell, cellStyle);
            return true;
        }
        return false;
    }

    private static (object? Value, ExcelRange? SrcCell) GetValueAndCell(string columnName, int? oldRow, IExcelDataSource excelDataSource)
    {
        if (oldRow is null) { return (null, null); }
        ExcelRange? srcCell = excelDataSource.GetExcelRange(columnName, oldRow.Value);
        object? value = srcCell?.Value ?? excelDataSource.GetCellValue(columnName, oldRow.Value);
        return (value, srcCell);
    }

    private object? SetCell(ExcelRange dstCell, string columnName, int? oldRow, ModificationRuleHandler ruleHandler, DataKind dataKind)
    {
        (object? value, ExcelRange? srcCell) = GetValueAndCell(columnName, oldRow, dataKind.HasFlag(DataKind.Old) ? oldDataSource : newDataSource);
        if (!dataKind.HasFlag(DataKind.Old) || config.ShowOldDataColumn)
        {
            dstCell.Value = value;
            if (config.CopyCellStyle) { ExcelHelper.CopyCellStyle(dstCell, srcCell); }
            if (config.CopyCellFormat) { ExcelHelper.CopyCellFormat(dstCell, srcCell); }
            ruleHandler.ApplyRules(dstCell, columnName, dataKind);
            return dstCell.Value;
        }
        return value;
    }

    private int WriteHeader(ExcelWorksheet worksheet, int startRow, int startColumn)
    {
        int column = startColumn;
        foreach (string columnName in excelDiffOp.MergedColumnNames)
        {
            if (config.ColumnsToOmit.Contains(columnName, stringComparer))
            {
                continue;
            }
            if (config.ShowOldDataColumn)
            {
                worksheet.Cells[startRow, column].Value = config.OldHeaderColumnPostfix is { } oldPostfix ? columnName + oldPostfix : columnName;
                if (config.OldHeaderColumnComment is not null)
                {
                    _ = worksheet.Cells[startRow, column].AddComment(config.OldHeaderColumnComment);
                }
                column++;
            }
            worksheet.Cells[startRow, column].Value = config.NewHeaderColumnPostfix is { } newPostfix ? columnName + newPostfix : columnName;
            if (config.NewHeaderColumnComment is not null)
            {
                _ = worksheet.Cells[startRow, column].AddComment(config.NewHeaderColumnComment);
            }
            column++;
        }
        column--;
        ExcelHelper.SetCellStyle(worksheet.Cells[startRow, startColumn, startRow, column], config.HeaderStyle);
        return column;
    }

}
