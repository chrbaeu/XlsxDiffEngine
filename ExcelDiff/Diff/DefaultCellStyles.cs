using System.Drawing;

namespace ExcelDiffEngine;

public static class DefaultCellStyles
{
    public static readonly CellStyle None = new();
    public static readonly CellStyle Header = new() { Bold = true };
    public static readonly CellStyle AddedRow = new() { BackgroundColor = Color.FromArgb(173, 254, 173) };
    public static readonly CellStyle RemovedRow = new() { BackgroundColor = Color.LightGray, FontColor = Color.DarkGray };
    public static readonly CellStyle ChangedCell = new() { BackgroundColor = Color.FromArgb(255, 178, 101) };
    public static readonly CellStyle ChangedRowKeyColumns = new() { BackgroundColor = Color.FromArgb(150, 175, 255) };
    public static readonly CellStyle YellowValueChangedMarker = new() { BackgroundColor = Color.FromArgb(254, 255, 101) };
    public static readonly CellStyle OrangeValueChangedMarker = new() { BackgroundColor = Color.FromArgb(255, 178, 101) };
    public static readonly CellStyle RedValueChangedMarker = new() { BackgroundColor = Color.FromArgb(255, 99, 99) };
}