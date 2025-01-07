using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using XlsxDiffTool.Common;
using XlsxDiffTool.Models;
using XlsxDiffTool.Services;
using System.Collections.ObjectModel;

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
