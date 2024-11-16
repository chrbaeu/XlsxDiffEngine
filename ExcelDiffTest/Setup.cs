namespace ExcelDiffTest;

internal class Setup
{
    [Before(Assembly)]
    public static void Initialize()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
}
