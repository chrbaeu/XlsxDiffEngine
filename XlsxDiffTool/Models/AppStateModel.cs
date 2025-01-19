using CommunityToolkit.Mvvm.ComponentModel;

namespace XlsxDiffTool.Models;

public partial class AppStateModel : ObservableObject
{
    [ObservableProperty]
    public partial bool IsBusy { get; set; }
}
