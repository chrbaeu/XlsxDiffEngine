using ExcelDiffEngine;

namespace ExcelDiffTest;

public class SimpleExcelDocTest
{
    public string SimpleDocOld = Path.Combine("SimpleDoc", "SimpleDocNew.xlsx");
    public string SimpleDocNew = Path.Combine("SimpleDoc", "SimpleDocOld.xlsx");

    [Test]
    public void MyTest()
    {
        new ExcelDiffBuilder()
            .AddFiles(x => x
                .SetOldFile(SimpleDocOld)
                .SetNewFile(SimpleDocNew)
                )
            .Build("Vergleich.xlsx");
    }
}
