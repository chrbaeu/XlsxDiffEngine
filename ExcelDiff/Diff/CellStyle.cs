using System.Drawing;

namespace ExcelDiffEngine;

public record class CellStyle
{
    public Color? FontColor { get; init; }
    public Color? BackgroundColor { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
}