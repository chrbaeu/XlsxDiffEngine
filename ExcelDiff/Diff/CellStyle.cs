using System.Drawing;

namespace ExcelDiffEngine;

/// <summary>
/// Represents styling options for a cell in an Excel sheet, including font styles and colors.
/// </summary>
public sealed record class CellStyle
{
    /// <summary>
    /// Specifies whether the font is bold. If null, the bold style is unspecified.
    /// </summary>
    public bool? Bold { get; init; }

    /// <summary>
    /// Specifies whether the font is italicized. If null, the italic style is unspecified.
    /// </summary>
    public bool? Italic { get; init; }

    /// <summary>
    /// Specifies whether the font is underlined. If null, the underline style is unspecified.
    /// </summary>
    public bool? Underline { get; init; }

    /// <summary>
    /// The color of the font. If null, the font color is unspecified.
    /// </summary>
    public Color? FontColor { get; init; }

    /// <summary>
    /// The background color of the cell. If null, the background color is unspecified.
    /// </summary>
    public Color? BackgroundColor { get; init; }
}
