using OfficeOpenXml;

namespace ExcelDiffTest;

public class Setup
{
    [Before(Assembly)]
    public static async Task Initialize()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
}
