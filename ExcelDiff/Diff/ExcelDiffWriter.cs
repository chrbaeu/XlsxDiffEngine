using OfficeOpenXml;

namespace ExcelDiffEngine;

public sealed class ExcelDiffWriter
{
    private readonly IExcelDataSource oldDataSource;
    private readonly IExcelDataSource newDataSource;
    private readonly ExcelDiffConfig config;
    private readonly StringComparer stringComparer;
    private readonly ExcelDiffOp excelDiffOp;

    public ExcelDiffWriter(IExcelDataSource oldDataSource, IExcelDataSource newDataSource, ExcelDiffConfig config)
    {
        this.oldDataSource = oldDataSource;
        this.newDataSource = newDataSource;
        this.config = config;
        stringComparer = config.IgnoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        excelDiffOp = new(oldDataSource, newDataSource, config);
    }

    public (int endRow, int endColumn) WriteDiff(ExcelWorksheet worksheet, int row = 1, int column = 1)
    {
        int headerEndColumns = WriteHeader(worksheet, row, column);
        int dataEndRow = WriteData(worksheet, row + 1, column);
        return (dataEndRow, headerEndColumns);
    }

    private int WriteData(ExcelWorksheet worksheet, int row, int startColumn)
    {
        ModificationRuleHandler ruleHandler = new(config.ModificationRules, config.IgnoreCase);
        foreach ((var oldRow, var newRow) in excelDiffOp.GetMergedRows())
        {
            var column = startColumn;
            var isChanged = false;
            List<ExcelRange> keyCells = [];
            foreach (var columnName in excelDiffOp.MergedColumnNames)
            {
                if (config.ColumnsToOmit.Contains(columnName, stringComparer))
                {
                    continue;
                }
                var oldDstCell = worksheet.Cells[row, column];
                var oldValue = SetCell(oldDstCell, columnName, oldRow, ruleHandler, DataKind.Old);
                if (config.ShowOldDataColumn) { column++; } else { oldDstCell = null; }
                var newDstCell = worksheet.Cells[row, column];
                var newValue = SetCell(newDstCell, columnName, newRow, ruleHandler, DataKind.New);
                if (config.AddOldValueAsComment && oldValue?.ToString() != newValue?.ToString())
                {
                    var comment = config.OldValueCommentPrefix is { } prefix ? prefix + oldValue?.ToString() : oldValue?.ToString();
                    newDstCell.AddComment(comment ?? "");
                }
                column++;
                isChanged |= GetAndHandleChangedState(columnName, oldDstCell, oldValue, newDstCell, newValue);
                if (config.KeyColumns.Contains(columnName, stringComparer))
                {
                    if (oldDstCell is not null) { keyCells.Add(oldDstCell); }
                    keyCells.Add(newDstCell);
                }
            }
            if (isChanged && config.ChangedRowKeyColumsStyle is not null)
            {
                foreach (var keyCell in keyCells)
                {
                    ExcelHelper.SetCellStyle(keyCell, config.ChangedRowKeyColumsStyle);
                }
            }
            if (oldRow is null) { ExcelHelper.SetCellStyle(worksheet.Cells[row, startColumn, row, column - 1], config.AddedRowStyle); }
            if (newRow is null) { ExcelHelper.SetCellStyle(worksheet.Cells[row, startColumn, row, column - 1], config.RemovedRowStyle); }
            if (config.IgnoreUnchangedColumns && !isChanged)
            {
                worksheet.Cells[row, startColumn, row, column - 1].Clear();
                continue;
            }
            row++;
        }
        return row - 1;
    }

    public bool GetAndHandleChangedState(string columnName, ExcelRange? oldDstCell, object? oldValue, ExcelRange newDstCell, object? newValue)
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
        if (oldValue is double oldNumber && newValue is double newNumber
            && !config.ColumnsToTextCompareOnly.Contains(columnName, stringComparer)
            && config.ValueChangedMarkers.Count > 0)
        {
            var pDiff = Math.Abs((oldNumber - newNumber) / ((oldNumber + newNumber) / 2.0));
            var aDiff = Math.Abs(oldNumber - newNumber);
            foreach (var valueChangedMarker in config.ValueChangedMarkers)
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

    private (object? Value, ExcelRange? SrcCell) GetValueAndCell(string columnName, int? oldRow, IExcelDataSource excelDataSource)
    {
        if (oldRow is null) { return (null, null); }
        var srcCell = excelDataSource.GetExcelRange(columnName, oldRow.Value);
        var value = srcCell?.Value ?? excelDataSource.GetCellValue(columnName, oldRow.Value);
        return (value, srcCell);
    }

    private object? SetCell(ExcelRange dstCell, string columnName, int? oldRow, ModificationRuleHandler ruleHandler, DataKind dataKind)
    {
        var (value, srcCell) = GetValueAndCell(columnName, oldRow, dataKind == DataKind.Old ? oldDataSource : newDataSource);
        if (dataKind != DataKind.Old || config.ShowOldDataColumn)
        {
            dstCell.Value = value;
            if (config.CopyCellStyle) { ExcelHelper.CopyCellStyle(dstCell, srcCell); }
            if (config.CopyCellFormat) { ExcelHelper.CopyCellFormat(dstCell, srcCell); }
            ruleHandler.ApplyRules(dstCell, columnName, DataKind.Old);
            return dstCell.Value;
        }
        return value;
    }

    private int WriteHeader(ExcelWorksheet worksheet, int startRow, int startColumn)
    {
        int column = startColumn;
        foreach (var columnName in excelDiffOp.MergedColumnNames)
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
                    worksheet.Cells[startRow, column].AddComment(config.OldHeaderColumnComment);
                }
                column++;
            }
            worksheet.Cells[startRow, column].Value = config.NewHeaderColumnComment is { } newPostfix ? columnName + newPostfix : columnName;
            if (config.NewHeaderColumnComment is not null)
            {
                worksheet.Cells[startRow, column].AddComment(config.NewHeaderColumnComment);
            }
            column++;
        }
        column--;
        ExcelHelper.SetCellStyle(worksheet.Cells[startRow, startColumn, startRow, column], config.HeaderStyle);
        return column;
    }

}
