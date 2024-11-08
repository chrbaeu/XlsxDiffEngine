using System.Drawing;

namespace ExcelDiffEngine;

/// <summary>
/// Provides a set of predefined <see cref="CellStyle"/> objects representing common cell styles 
/// used for various data states or types in Excel comparisons.
/// </summary>
public static class DefaultCellStyles
{
    /// <summary>
    /// Represents a default cell with no specific formatting and style.
    /// </summary>
    public static readonly CellStyle None = new();

    /// <summary>
    /// Represents a header cell style with bold text.
    /// </summary>
    public static readonly CellStyle Header = new() { Bold = true };

    /// <summary>
    /// Represents the style for newly added rows, with a light green background.
    /// </summary>
    public static readonly CellStyle AddedRow = new() { BackgroundColor = Color.FromArgb(173, 254, 173) };

    /// <summary>
    /// Represents the style for removed rows, with a light gray background and dark gray font color.
    /// </summary>
    public static readonly CellStyle RemovedRow = new() { BackgroundColor = Color.LightGray, FontColor = Color.DarkGray };

    /// <summary>
    /// Represents the style for cells with changed values, with a light orange background.
    /// </summary>
    public static readonly CellStyle ChangedCell = new() { BackgroundColor = Color.FromArgb(255, 178, 101) };

    /// <summary>
    /// Represents the style for key columns in rows with changes, with a light blue background.
    /// </summary>
    public static readonly CellStyle ChangedRowKeyColumns = new() { BackgroundColor = Color.FromArgb(150, 175, 255) };

    /// <summary>
    /// Represents a yellow marker style for cells with changed values, with a yellow background.
    /// </summary>
    public static readonly CellStyle YellowValueChangedMarker = new() { BackgroundColor = Color.FromArgb(254, 255, 101) };

    /// <summary>
    /// Represents an orange marker style for cells with changed values, with a light orange background.
    /// </summary>
    public static readonly CellStyle OrangeValueChangedMarker = new() { BackgroundColor = Color.FromArgb(255, 178, 101) };

    /// <summary>
    /// Represents a red marker style for cells with significant value changes, with a light red background.
    /// </summary>
    public static readonly CellStyle RedValueChangedMarker = new() { BackgroundColor = Color.FromArgb(255, 99, 99) };
}
