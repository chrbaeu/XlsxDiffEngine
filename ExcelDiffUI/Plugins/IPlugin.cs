using ExcelDiffEngine;
using ExcelDiffUI.Models;
using OfficeOpenXml;

namespace ExcelDiffUI.Services;

public interface IPlugin
{
    public string Name { get; }
    public string Tooltip { get; }
    public void OnExcelPackageLoading(DiffConfigModel diffConfigModel, ExcelPackage excelPackage);
    public Task OnDiffCreation(DiffConfigModel diffConfigModel, ExcelDiffBuilder excelDiffBuilder);
    public void OnExcelPackageSaving(DiffConfigModel diffConfigModel, ExcelPackage excelPackage);
}
