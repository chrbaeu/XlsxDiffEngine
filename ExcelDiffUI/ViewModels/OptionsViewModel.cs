using CommunityToolkit.Mvvm.ComponentModel;
using ExcelDiffUI.Common;
using ExcelDiffUI.Models;

namespace ExcelDiffUI.ViewModels;

public partial class OptionsViewModel(DiffConfigModel diffOptions) : ObservableObject, IViewModel
{
    public DiffConfigModel DiffOptions { get; } = diffOptions;
}
