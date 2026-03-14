using System.Diagnostics.CodeAnalysis;

namespace XlsxDiffTool.Common;

public enum DialogButton
{
    OK = 0,
    OKCancel = 1,
    YesNoCancel = 3,
    YesNo = 4
}

public enum DialogResult
{
    None = 0,
    OK = 1,
    Cancel = 2,
    Yes = 6,
    No = 7
}

public interface IDialogService
{
    public SynchronizationContext? SynchronizationContext { get; }

    public DialogResult ShowMessageBox(object? owner, string title, string text, DialogButton buttons = DialogButton.OK);

    public string ShowOpenFileDialog(object? owner, string? extensions = null, string? initialDirectory = null, string? fileName = null, string? title = null);
    public string[] ShowOpenFileMultiselectDialog(object? owner, string? extensions = null, string? initialDirectory = null, string? fileName = null, string? title = null);
    public string ShowSaveFileDialog(object? owner, string? extensions = null, string? initialDirectory = null, string? fileName = null, string? title = null);
    public string ShowOpenFolderDialog(object? owner, string? initialDirectory = null, string? title = null);

    public Task ShowDialogAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] TViewModel>(object? owner, params object[] parameters) where TViewModel : IViewModel;
    public Task<TResult?> ShowDialogAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] TViewModel, TResult>(object? owner, params object[] parameters) where TViewModel : IViewModel;

    public Task<TViewModel> ShowWindowAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] TViewModel>(object? owner, params object[] parameters) where TViewModel : IViewModel;
    public Task<(TViewModel ViewModel, Task<TResult?> ResultTask)> ShowWindowAsync<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicConstructors)] TViewModel, TResult>(object? owner, params object[] parameters) where TViewModel : IViewModel;

    public void CloseWindow(IViewModel viewModel);

}
