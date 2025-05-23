using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace XlsxDiffTool.Services;

public class InvalidScriptStatementException : Exception
{
    public InvalidScriptStatementException(int lineNumber, string statement)
        : base($"The statement ‘{statement}’ in line {lineNumber} is invalid.") { }
    public InvalidScriptStatementException(int lineNumber, string command, string message)
        : base($"Executing command ‘{command}’ in line {lineNumber} failed: {message}") { }
}

public partial class ExcelScriptExecutor
{

    [GeneratedRegex(@"^(?<command>\w+)\[(?<arg>[^\]]+)\](?:=(?<value>.*))?$", RegexOptions.IgnoreCase)]
    private static partial Regex ParseInstructionRegex { get; }

    [GeneratedRegex(@"^[A-Z]+\d+$", RegexOptions.IgnoreCase)]
    private static partial Regex IsValidCellAddressRegex { get; }

    [GeneratedRegex(@"^[A-Z]+\d+(:[A-Z]+\d+)?$", RegexOptions.IgnoreCase)]
    private static partial Regex IsValidCellRangeRegex { get; }

    [GeneratedRegex(@"^[A-Z]+$", RegexOptions.IgnoreCase)]
    private static partial Regex IsValidColumnLetterRegex { get; }

    private sealed record class Statement(int LineNumber, string Command, string Arg, string Value);

    /// <summary>
    /// Executes a script on an ExcelPackage.
    /// </summary>
    /// <param name="package">The ExcelPackage to run the script on.</param>
    /// <param name="script">The script to execute.</param>
    public static void ExecuteScript(ExcelPackage package, string script)
    {
        List<ExcelWorksheet> currentWorksheets = [.. package.Workbook.Worksheets];

        var scriptData = script.Split(["\r\n", "\n", "\r"], StringSplitOptions.TrimEntries)
            .Select((statement, index) => (Statement: statement, LineNumber: index + 1))
            .ToList();

        foreach (var item in scriptData)
        {
            if (string.IsNullOrEmpty(item.Statement)) { continue; }

            var match = ParseInstructionRegex.Match(item.Statement);
            if (!match.Success)
            {
                throw new InvalidScriptStatementException(item.LineNumber, item.Statement);
            }

            Statement statement = new(
                item.LineNumber,
                match.Groups["command"].Value.ToUpperInvariant().Trim(),
                match.Groups["arg"].Value.ToUpperInvariant().Trim(),
                match.Groups["value"].Value.Trim()
            );

            switch (statement.Command)
            {
                case "WORKSHEET":
                    ProcessWorksheetCommand(statement, currentWorksheets, package);
                    break;
                case "SET":
                    ProcessSetCommand(statement, currentWorksheets);
                    break;
                case "SWAP":
                    ProcessSwapCommand(statement, currentWorksheets);
                    break;
                case "REMROW":
                    ProcessRemoveRowCommand(statement, currentWorksheets);
                    break;
                case "REMCOL":
                    ProcessRemoveColumnCommand(statement, currentWorksheets);
                    break;
                case "ADDROW":
                    ProcessAddRowCommand(statement, currentWorksheets);
                    break;
                case "ADDCOL":
                    ProcessAddColumnCommand(statement, currentWorksheets);
                    break;
                default:
                    throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, "The command name is not valid.");
            }
        }
    }

    private static void ProcessWorksheetCommand(Statement statement, List<ExcelWorksheet> currentWorksheets, ExcelPackage package)
    {
        if (statement.Arg.Equals("ALL", StringComparison.OrdinalIgnoreCase))
        {
            currentWorksheets.Clear();
            currentWorksheets.AddRange(package.Workbook.Worksheets);
        }
        else if (int.TryParse(statement.Arg, out int wsNumber))
        {
            if (wsNumber < 1 || wsNumber > package.Workbook.Worksheets.Count)
            {
                throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"No worksheet with index {wsNumber} found.");
            }
            currentWorksheets.Clear();
            currentWorksheets.Add(package.Workbook.Worksheets[wsNumber]);
        }
        else
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Argument '{statement.Arg}' is not a valid worksheet index.");
        }
    }

    private static void ProcessSetCommand(Statement statement, List<ExcelWorksheet> currentWorksheets)
    {
        object setValue = statement.Value.StartsWith('\"') && statement.Value.EndsWith('\"')
            ? statement.Value[1..^1]
            : int.TryParse(statement.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var intValue) ? intValue
            : double.TryParse(statement.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var doubleValue) ? doubleValue
            : throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Value '{statement.Value}' is not a valid value.");
        var cellRange = statement.Arg;
        if (!IsValidCellRangeRegex.IsMatch(cellRange))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Cell range '{cellRange}' is not valid.");
        }
        foreach (var ws in currentWorksheets)
        {
            ws.Cells[cellRange].Value = setValue;
        }
    }

    private static void ProcessSwapCommand(Statement statement, List<ExcelWorksheet> currentWorksheets)
    {
        var cellAddress1 = statement.Arg;
        if (!IsValidCellAddressRegex.IsMatch(cellAddress1))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Cell address '{cellAddress1}' is not valid.");
        }
        var cellAddress2 = statement.Value;
        if (cellAddress2.StartsWith('[') && cellAddress2.EndsWith(']'))
        {
            cellAddress2 = cellAddress2[1..^1];
        }
        else
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Value '{statement.Value}' is not a valid cell address format.");
        }
        if (!IsValidCellAddressRegex.IsMatch(cellAddress2))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Cell address '{cellAddress2}' is not valid.");
        }
        foreach (var ws in currentWorksheets)
        {
            var cell1 = ws.Cells[cellAddress1];
            var cell2 = ws.Cells[cellAddress2];
            var tempValue = cell1.Value;
            var tempFormula = cell1.Formula;
            if (string.IsNullOrEmpty(cell2.Formula))
            {
                cell1.Value = cell2.Value;
            }
            else
            {
                cell1.Formula = cell2.Formula;
            }
            if (string.IsNullOrEmpty(tempFormula))
            {
                cell2.Value = tempValue;
            }
            else
            {
                cell2.Formula = tempFormula;
            }
            SwapCellStyles(cell1, cell2);
        }
    }

    private static void ProcessRemoveRowCommand(Statement statement, List<ExcelWorksheet> currentWorksheets)
    {
        if (!int.TryParse(statement.Arg, out int rowNumber))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Argument '{statement.Arg}' is not a valid row index.");
        }
        foreach (var ws in currentWorksheets)
        {
            ws.DeleteRow(rowNumber);
        }
    }

    private static void ProcessRemoveColumnCommand(Statement statement, List<ExcelWorksheet> currentWorksheets)
    {
        if (!IsValidColumnLetterRegex.IsMatch(statement.Arg))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Argument '{statement.Arg}' is not a valid column letter.");
        }
        var colNumber = ColumnLetterToNumber(statement.Arg);
        foreach (var ws in currentWorksheets)
        {
            ws.DeleteColumn(colNumber);
        }
    }

    private static void ProcessAddRowCommand(Statement statement, List<ExcelWorksheet> currentWorksheets)
    {
        if (!int.TryParse(statement.Arg, out int rowNumber))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Argument '{statement.Arg}' is not a valid row index.");
        }
        foreach (var ws in currentWorksheets)
        {
            ws.InsertRow(rowNumber, 1);
        }
    }

    private static void ProcessAddColumnCommand(Statement statement, List<ExcelWorksheet> currentWorksheets)
    {
        if (!IsValidColumnLetterRegex.IsMatch(statement.Arg))
        {
            throw new InvalidScriptStatementException(statement.LineNumber, statement.Command, $"Argument '{statement.Arg}' is not a valid column letter.");
        }
        int colNumber = ColumnLetterToNumber(statement.Arg);
        foreach (var ws in currentWorksheets)
        {
            ws.InsertColumn(colNumber, 1);
        }
    }

    private static int ColumnLetterToNumber(string columnLetter)
    {
        int sum = 0;
        foreach (var c in columnLetter)
        {
            sum = sum * 26 + (c - 'A' + 1);
        }
        return sum;
    }

    private static void SwapCellStyles(ExcelRange cell1, ExcelRange cell2)
    {
        (cell2.Style.Numberformat.Format, cell1.Style.Numberformat.Format) = (cell1.Style.Numberformat.Format, cell2.Style.Numberformat.Format);
        (cell2.Style.Font.Name, cell1.Style.Font.Name) = (cell1.Style.Font.Name, cell2.Style.Font.Name);
        (cell2.Style.Font.Size, cell1.Style.Font.Size) = (cell1.Style.Font.Size, cell2.Style.Font.Size);
        (cell2.Style.Font.Bold, cell1.Style.Font.Bold) = (cell1.Style.Font.Bold, cell2.Style.Font.Bold);
        (cell2.Style.Font.Italic, cell1.Style.Font.Italic) = (cell1.Style.Font.Italic, cell2.Style.Font.Italic);
        (cell2.Style.Font.UnderLine, cell1.Style.Font.UnderLine) = (cell1.Style.Font.UnderLine, cell2.Style.Font.UnderLine);
        (cell2.Style.Font.UnderLineType, cell1.Style.Font.UnderLineType) = (cell1.Style.Font.UnderLineType, cell2.Style.Font.UnderLineType);

        var tempFill = cell1.Style.Fill;
        cell1.Style.Fill.PatternType = cell2.Style.Fill.PatternType;
        if (cell1.Style.Fill.PatternType != ExcelFillStyle.None)
        {
            TransferColor(cell1.Style.Fill.BackgroundColor, cell2.Style.Fill.BackgroundColor);
        }
        cell2.Style.Fill.PatternType = tempFill.PatternType;
        if (cell2.Style.Fill.PatternType != ExcelFillStyle.None)
        {
            TransferColor(cell2.Style.Fill.BackgroundColor, tempFill.BackgroundColor);
        }

        var tempFontColor = cell1.Style.Font.Color;
        TransferColor(cell1.Style.Font.Color, cell2.Style.Font.Color);
        TransferColor(cell1.Style.Font.Color, tempFontColor);
    }

    private static void TransferColor(ExcelColor dstExcelColor, ExcelColor srcExcelColor)
    {
        if (!string.IsNullOrEmpty(srcExcelColor.Rgb))
        {
            dstExcelColor.SetColor(Color.FromArgb(int.Parse(srcExcelColor.Rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture)));
        }
        else if (srcExcelColor.Theme != null)
        {
            dstExcelColor.SetColor(srcExcelColor.Theme.Value);
            dstExcelColor.Tint = srcExcelColor.Tint;
        }
        else if (srcExcelColor.Indexed != int.MinValue)
        {
            dstExcelColor.SetColor((ExcelIndexedColor)srcExcelColor.Indexed);
        }
    }
}
