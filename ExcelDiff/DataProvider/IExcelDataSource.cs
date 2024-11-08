using OfficeOpenXml;

namespace ExcelDiffEngine;

/// <summary>
/// Represents a data source for accessing <see cref="ExcelWorksheet"/> data. This interface provides methods 
/// for retrieving information about columns, rows, and cells within an <see cref="ExcelWorksheet"/>.
/// </summary>
public interface IExcelDataSource
{
    /// <summary>
    /// The name of the data source.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// The number of data rows available in the data source.
    /// </summary>
    public int DataRows { get; }

    /// <summary>
    /// Retrieves the names of all columns in the data source.
    /// </summary>
    /// <returns>A read-only collection of column names.</returns>
    public IReadOnlyCollection<string> GetColumnNames();

    /// <summary>
    /// Gets the <see cref="ExcelRange"/> for a specific cell in a given column and row.
    /// </summary>
    /// <param name="columnName">The name of the column containing the cell.</param>
    /// <param name="row">The row number of the cell.</param>
    /// <returns>The <see cref="ExcelRange"/> of the cell, or null if the cell is not found.</returns>
    public ExcelRange? GetExcelRange(string columnName, int row);

    /// <summary>
    /// Retrieves the value of a specific cell in a given column and row.
    /// </summary>
    /// <param name="columnName">The name of the column containing the cell.</param>
    /// <param name="row">The row number of the cell.</param>
    /// <returns>The cell’s value, or null if the cell has no value or does not exist.</returns>
    public object? GetCellValue(string columnName, int row);

    /// <summary>
    /// Retrieves the text content of a specific cell in a given column and row.
    /// </summary>
    /// <param name="columnName">The name of the column containing the cell.</param>
    /// <param name="row">The row number of the cell.</param>
    /// <returns>The text content of the cell as a <see cref="string"/>.</returns>
    public string GetCellText(string columnName, int row);

    /// <summary>
    /// Retrieves all cell values in a specific row as a dictionary, with column names as keys.
    /// </summary>
    /// <param name="row">The row number to retrieve.</param>
    /// <returns>A read-only dictionary mapping column names to their corresponding cell values.</returns>
    public IReadOnlyDictionary<string, object?> GetRow(int row);

    /// <summary>
    /// Retrieves all values of a specific column as an array.
    /// </summary>
    /// <param name="columnName">The name of the column to retrieve.</param>
    /// <returns>An array containing the values of the specified column.</returns>
    public object?[] GetColumn(string columnName);
}
