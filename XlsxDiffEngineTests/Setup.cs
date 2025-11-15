namespace XlsxDiffEngineTests;

internal class Setup
{
    [Before(Assembly)]
    public static void Initialize()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Christian Baeumlisberger");
    }
}
