using OfficeOpenXml;

namespace ExcelDiffSample;

internal class Program
{
    static void Main()
    {
        Console.WriteLine("ExcelDiff Sample");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var list = new List<string>();
        list.Clear();
    }
}
