using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using XlsxDiffTool.Common;
using XlsxDiffTool.Models;
using XlsxDiffTool.Services;

namespace XlsxDiffTool.ViewModels;

public partial class ColumnSelectorViewModel(ColumnInfoService columnInfoService) : ObservableObject, IViewModel
{
    [ObservableProperty]
    public partial string ColumnName { get; set; } = "";

    public ObservableCollection<ColumnInfoModel> Columns { get; } = columnInfoService.Columns;

    [RelayCommand]
    public void AddColumn()
    {
        if (ColumnName is string columnName and not "")
        {
            columnInfoService.AddManualColumn(columnName);
            ColumnName = "";
        }
    }
}
