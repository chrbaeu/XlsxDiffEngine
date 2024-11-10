using OfficeOpenXml;
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

    public static bool CheckIfExcelPackagesIdentical(ExcelPackage excelPackageA, ExcelPackage excelPackageB, bool compareCellBackground = false)
    {
        if (excelPackageA.Workbook.Worksheets.Count != excelPackageB.Workbook.Worksheets.Count)
        {
            return false;
        }
        for (int i = 0; i < excelPackageA.Workbook.Worksheets.Count; i++)
        {
            var worksheetA = excelPackageA.Workbook.Worksheets[i];
            var worksheetB = excelPackageB.Workbook.Worksheets[i];
            if (worksheetA.Name != worksheetB.Name)
            {
                return false;
            }
            if (worksheetA.Dimension?.Rows != worksheetB.Dimension?.Rows || worksheetA.Dimension?.Columns != worksheetB.Dimension?.Columns)
            {
                return false;
            }
            for (int row = 1; row <= worksheetA.Dimension.Rows; row++)
            {
                for (int col = 1; col <= worksheetA.Dimension.Columns; col++)
                {
                    var cellA = worksheetA.Cells[row, col].Text;
                    var cellB = worksheetB.Cells[row, col].Text;
                    if (cellA != cellB)
                    {
                        return false;
                    }
                    var formatA = worksheetA.Cells[row, col].Style.Numberformat;
                    var formatB = worksheetB.Cells[row, col].Style.Numberformat;
                    if (formatA.NumFmtID != formatB.NumFmtID || formatA.Format != formatB.Format)
                    {
                        return false;
                    }
                    var styleA = worksheetA.Cells[row, col].Style;
                    var styleB = worksheetB.Cells[row, col].Style;
                    if (styleA.Font.Bold != styleB.Font.Bold || styleA.Font.Italic != styleB.Font.Italic
                        || styleA.Font.UnderLine != styleB.Font.UnderLine || styleA.Font.Size != styleB.Font.Size
                        || !IsSameColor(styleA.Font.Color, styleB.Font.Color))
                    {
                        return false;
                    }
                    if (compareCellBackground && (styleA.Fill.PatternType != styleB.Fill.PatternType
                        || !IsSameColor(styleA.Fill.BackgroundColor, styleB.Fill.BackgroundColor)))
                    {
                        return false;
                    }
                }
            }
        }
        return true;
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
