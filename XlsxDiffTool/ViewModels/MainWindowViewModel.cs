using CommunityToolkit.Mvvm.ComponentModel;
using XlsxDiffTool.Common;

namespace XlsxDiffTool.ViewModels;

public sealed partial class MainWindowViewModel(
    MainViewModel mainViewModel,
    AppInfo appInfo,
    WindowStateModel windowState
    ) : ObservableObject, IViewModel
{
    public WindowStateModel WindowState { get; } = windowState;

    [ObservableProperty]
    public partial string Title { get; set; } = $"{appInfo.AppName} V{appInfo.Version.ToString(3)}";

    public MainViewModel MainViewModel { get; } = mainViewModel;

}
