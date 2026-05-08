using OfficeOpenXml.Style;

namespace XlsxDiffEngineTests;

internal static class ExcelTestHelper
{
    public static ExcelPackage ConvertToExcelPackage(object?[][] data, string worksheetName = "Table")
    {
        ExcelPackage excelPackage = new();
        excelPackage.AddWorksheet(data, worksheetName);
        return excelPackage;
    }

    public static async Task AssertExcelPackagesIdentical(ExcelPackage excelPackageA, ExcelPackage excelPackageB, bool compareCellBackground = false)
    {
        await AssertEqual(
            excelPackageA.Workbook.Worksheets.Count,
            excelPackageB.Workbook.Worksheets.Count,
            "the number of worksheets is different");

        for (int i = 0; i < excelPackageA.Workbook.Worksheets.Count; i++)
        {
            ExcelWorksheet worksheetA = excelPackageA.Workbook.Worksheets[i];
            ExcelWorksheet worksheetB = excelPackageB.Workbook.Worksheets[i];

            await AssertEqual(
                worksheetA.Name,
                worksheetB.Name,
                $"worksheet names do not match at index {i + 1}");
            await AssertEqual(
                worksheetA.Dimension?.Rows,
                worksheetB.Dimension?.Rows,
                $"worksheet row count in '{worksheetA.Name}' does not match");
            await AssertEqual(
                worksheetA.Dimension?.Columns,
                worksheetB.Dimension?.Columns,
                $"worksheet column count in '{worksheetA.Name}' does not match");

            if (worksheetA.Dimension is null) { continue; }
            for (int row = 1; row <= worksheetA.Dimension.Rows; row++)
            {
                for (int col = 1; col <= worksheetA.Dimension.Columns; col++)
                {
                    string cellA = worksheetA.Cells[row, col].Text;
                    string cellB = worksheetB.Cells[row, col].Text;

                    await AssertEqual(
                        cellA,
                        cellB,
                        $"cell content in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");

                    ExcelNumberFormat formatA = worksheetA.Cells[row, col].Style.Numberformat;
                    ExcelNumberFormat formatB = worksheetB.Cells[row, col].Style.Numberformat;

                    using (Assert.Multiple())
                    {
                        await AssertEqual(
                            formatA.NumFmtID,
                            formatB.NumFmtID,
                            $"number format ID in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        await AssertEqual(
                            formatA.Format,
                            formatB.Format,
                            $"number format string in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                    }

                    ExcelStyle styleA = worksheetA.Cells[row, col].Style;
                    ExcelStyle styleB = worksheetB.Cells[row, col].Style;

                    string fontColorA = GetColorValue(styleA.Font.Color);
                    string fontColorB = GetColorValue(styleB.Font.Color);
                    using (Assert.Multiple())
                    {
                        await AssertEqual(
                            styleA.Font.Bold,
                            styleB.Font.Bold,
                            $"font bold flag in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        await AssertEqual(
                            styleA.Font.Italic,
                            styleB.Font.Italic,
                            $"font italic flag in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        await AssertEqual(
                            styleA.Font.UnderLine,
                            styleB.Font.UnderLine,
                            $"font underline flag in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        await AssertEqual(
                            styleA.Font.Size,
                            styleB.Font.Size,
                            $"font size in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        await AssertEqual(
                            styleA.Font.Strike,
                            styleB.Font.Strike,
                            $"font strike flag in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        await AssertEqual(
                            fontColorA,
                            fontColorB,
                            $"font color in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                    }

                    if (compareCellBackground)
                    {
                        string backgroundColorA = GetFillBackgroundColorValue(styleA.Fill.PatternType, styleA.Fill.BackgroundColor);
                        string backgroundColorB = GetFillBackgroundColorValue(styleB.Fill.PatternType, styleB.Fill.BackgroundColor);
                        using (Assert.Multiple())
                        {
                            await AssertEqual(
                                styleA.Fill.PatternType,
                                styleB.Fill.PatternType,
                                $"cell background pattern in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                            await AssertEqual(
                                backgroundColorA,
                                backgroundColorB,
                                $"cell background color in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]");
                        }
                    }
                }
            }
        }
    }

    private static Task AssertEqual<T>(T actual, T expected, string because)
    {
        if (EqualityComparer<T>.Default.Equals(actual, expected))
        {
            return Task.CompletedTask;
        }

        Assert.Fail($"{because}. Expected: {FormatValue(expected)}. Actual: {FormatValue(actual)}.");
        return Task.CompletedTask;
    }

    private static string FormatValue<T>(T value)
    {
        return value switch
        {
            null => "<null>",
            string text => $"'{text}'",
            _ => value.ToString() ?? "<null>",
        };
    }

    private static string GetColorValue(ExcelColor excelColor)
    {
        if (!string.IsNullOrEmpty(excelColor.Rgb))
        {
            return $"Rgb={excelColor.Rgb}";
        }

        if (excelColor.Theme != null)
        {
            return $"Theme={excelColor.Theme}; Tint={excelColor.Tint}";
        }

        if (excelColor.Indexed > 0)
        {
            return $"Indexed={excelColor.Indexed}";
        }

        return "<none>";
    }

    private static string GetFillBackgroundColorValue(ExcelFillStyle patternType, ExcelColor backgroundColor)
    {
        return patternType == ExcelFillStyle.None ? "<none>" : GetColorValue(backgroundColor);
    }

}
