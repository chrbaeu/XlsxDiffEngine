using System.Drawing;

namespace ExcelDiffEngine;

public sealed record class CellStyle
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    public Color? FontColor { get; init; }
    public Color? BackgroundColor { get; init; }
}