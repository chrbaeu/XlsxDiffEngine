using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

namespace ExcelDiffEngine;

internal static class ExcelHelper
{
    public static void SetCellStyle(ExcelRange? cells, CellStyle? cellStyle)
    {
        if (cells is null || cellStyle is null) { return; }
        if (cellStyle.FontColor is Color fontColor)
        {
            cells.Style.Font.Color.SetColor(fontColor);
        }
        if (cellStyle.BackgroundColor is Color backgroundColor)
        {
            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cells.Style.Fill.BackgroundColor.SetColor(backgroundColor);
        }
        if (cellStyle.Bold is bool bold)
        {
            cells.Style.Font.Bold = bold;
        }
        if (cellStyle.Italic is bool italic)
        {
            cells.Style.Font.Italic = italic;
        }
        if (cellStyle.Underline is bool underline)
        {
            cells.Style.Font.UnderLine = underline;
        }
    }

    public static void CopyCellStyle(ExcelRange dstCell, ExcelRange? srcCell)
    {
        if (srcCell is not ExcelRange sourceCell) { return; }
        dstCell.Style.Font.Bold = sourceCell.Style.Font.Bold;
        dstCell.Style.Font.Italic = sourceCell.Style.Font.Italic;
        dstCell.Style.Font.UnderLine = sourceCell.Style.Font.UnderLine;
        dstCell.Style.Font.Size = sourceCell.Style.Font.Size;
        TransferColor(dstCell.Style.Font.Color, sourceCell.Style.Font.Color);
        dstCell.Style.Fill.PatternType = sourceCell.Style.Fill.PatternType;
        if (sourceCell.Style.Fill.PatternType != ExcelFillStyle.None)
        {
            TransferColor(dstCell.Style.Fill.BackgroundColor, sourceCell.Style.Fill.BackgroundColor);
        }
    }

    public static void CopyCellFormat(ExcelRange dstCell, ExcelRange? srcCell)
    {
        if (srcCell is not ExcelRange sourceCell) { return; }
        dstCell.Style.Numberformat.Format = sourceCell.Style.Numberformat.Format;
    }

    private static void TransferColor(ExcelColor dstExcelColor, ExcelColor srcExcelColor)
    {
        if (!string.IsNullOrEmpty(srcExcelColor.Rgb))
        {
            dstExcelColor.SetColor(Color.FromArgb(int.Parse(srcExcelColor.Rgb, NumberStyles.HexNumber)));
        }
        else if (srcExcelColor.Theme != null)
        {
            dstExcelColor.SetColor(srcExcelColor.Theme.Value);
            dstExcelColor.Tint = srcExcelColor.Tint;
        }
        else
        {
            dstExcelColor.SetColor((ExcelIndexedColor)srcExcelColor.Indexed);
        }
    }

}
