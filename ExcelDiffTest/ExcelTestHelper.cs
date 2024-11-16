using OfficeOpenXml.Style;

namespace ExcelDiffTest;

internal static class ExcelTestHelper
{
    public static ExcelPackage ConvertToExcelPackage(object?[][] data, string worksheetName = "Table")
    {
        ExcelPackage excelPackage = new();
        excelPackage.AddWorksheet(data, worksheetName);
        return excelPackage;
    }

    public static void CheckIfExcelPackagesIdentical(ExcelPackage excelPackageA, ExcelPackage excelPackageB, bool compareCellBackground = false)
    {
        if (excelPackageA.Workbook.Worksheets.Count != excelPackageB.Workbook.Worksheets.Count)
        {
            throw new ArgumentException($"The number of worksheets is different: " +
                                         $"{excelPackageA.Workbook.Worksheets.Count} vs {excelPackageB.Workbook.Worksheets.Count}.");
        }

        for (int i = 0; i < excelPackageA.Workbook.Worksheets.Count; i++)
        {
            ExcelWorksheet worksheetA = excelPackageA.Workbook.Worksheets[i];
            ExcelWorksheet worksheetB = excelPackageB.Workbook.Worksheets[i];

            if (worksheetA.Name != worksheetB.Name)
            {
                throw new ArgumentException($"Worksheet names do not match: '{worksheetA.Name}' vs '{worksheetB.Name}'.");
            }

            if (worksheetA.Dimension?.Rows != worksheetB.Dimension?.Rows || worksheetA.Dimension?.Columns != worksheetB.Dimension?.Columns)
            {
                throw new ArgumentException($"Worksheet dimensions in '{worksheetA.Name}' do not match: " +
                                             $"{worksheetA.Dimension?.Rows}x{worksheetA.Dimension?.Columns} vs " +
                                             $"{worksheetB.Dimension?.Rows}x{worksheetB.Dimension?.Columns}.");
            }

            for (int row = 1; row <= worksheetA.Dimension.Rows; row++)
            {
                for (int col = 1; col <= worksheetA.Dimension.Columns; col++)
                {
                    string cellA = worksheetA.Cells[row, col].Text;
                    string cellB = worksheetB.Cells[row, col].Text;

                    if (cellA != cellB)
                    {
                        throw new ArgumentException($"Cell content in worksheet '{worksheetA.Name}' differs at position [{row}, {col}]: '{cellA}' vs '{cellB}'.");
                    }

                    ExcelNumberFormat formatA = worksheetA.Cells[row, col].Style.Numberformat;
                    ExcelNumberFormat formatB = worksheetB.Cells[row, col].Style.Numberformat;

                    if (formatA.NumFmtID != formatB.NumFmtID || formatA.Format != formatB.Format)
                    {
                        throw new ArgumentException($"Number format in worksheet '{worksheetA.Name}' differs at position [{row}, {col}].");
                    }

                    ExcelStyle styleA = worksheetA.Cells[row, col].Style;
                    ExcelStyle styleB = worksheetB.Cells[row, col].Style;

                    if (styleA.Font.Bold != styleB.Font.Bold || styleA.Font.Italic != styleB.Font.Italic
                        || styleA.Font.UnderLine != styleB.Font.UnderLine || styleA.Font.Size != styleB.Font.Size
                        || !IsSameColor(styleA.Font.Color, styleB.Font.Color))
                    {
                        throw new ArgumentException($"Font style in worksheet '{worksheetA.Name}' differs at position [{row}, {col}].");
                    }

                    if (compareCellBackground && (styleA.Fill.PatternType != styleB.Fill.PatternType
                                                  || !IsSameColor(styleA.Fill.BackgroundColor, styleB.Fill.BackgroundColor)))
                    {
                        throw new ArgumentException($"Cell background in worksheet '{worksheetA.Name}' differs at position [{row}, {col}].");
                    }
                }
            }
        }
    }

    private static bool IsSameColor(ExcelColor excelColorA, ExcelColor excelColorB)
    {
        if (!string.IsNullOrEmpty(excelColorA.Rgb))
        {
            return excelColorA.Rgb == excelColorB.Rgb;
        }
        else if (excelColorA.Theme != null)
        {
            return excelColorB.Theme == excelColorA.Theme && excelColorB.Tint == excelColorA.Tint;
        }
        else if (excelColorA.Indexed != int.MinValue)
        {
            return excelColorB.Indexed == excelColorA.Indexed;
        }
        return true;
    }

}
