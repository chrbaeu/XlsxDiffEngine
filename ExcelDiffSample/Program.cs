using OfficeOpenXml;

namespace ExcelDiffSample;

internal class Program
{
    internal static void Main()
    {
        Console.WriteLine("ExcelDiff Sample");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    }
}
