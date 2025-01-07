using OfficeOpenXml;
using XlsxDiffEngine;
using XlsxDiffTool.Models;

namespace XlsxDiffTool.Plugins;

public interface IPlugin
{
    public string Name { get; }
    public string Tooltip { get; }
    public void OnExcelPackageLoading(DiffConfigModel diffConfigModel, ExcelPackage excelPackage);
    public Task OnDiffCreation(DiffConfigModel diffConfigModel, ExcelDiffBuilder excelDiffBuilder);
    public void OnExcelPackageSaving(DiffConfigModel diffConfigModel, ExcelPackage excelPackage);
}
