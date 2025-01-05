using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelDiffUI.Common;

internal sealed class DialogServiceWpf(IServiceProvider serviceProvider, ViewFactory viewFactory) : IDialogService
{
    public SynchronizationContext? SynchronizationContext { get; } = SynchronizationContext.Current;

    public DialogResult ShowMessageBox(object? owner, string title, string text, DialogButton buttons = DialogButton.OK)
    {
        owner = GetParentWindow(owner);
        MessageBoxResult result = MessageBox.Show(owner as Window ?? Application.Current.MainWindow, text, title, (MessageBoxButton)buttons);
        return (DialogResult)result;
    }

    public string ShowOpenFileDialog(object? owner, string? extensions = null, string? initialDirectory = null, string? fileName = null, string? title = null)
    {
        OpenFileDialog openFileDialog = new();
        if (!string.IsNullOrEmpty(fileName)) { openFileDialog.FileName = fileName; }
        if (!string.IsNullOrEmpty(extensions)) { openFileDialog.Filter = extensions; }
        if (!string.IsNullOrEmpty(title)) { openFileDialog.Title = title; }
        if (!string.IsNullOrEmpty(initialDirectory) && Directory.Exists(initialDirectory)) { openFileDialog.InitialDirectory = initialDirectory; }
        if (openFileDialog.ShowDialog(GetParentWindow(owner)) == true) { return openFileDialog.FileName ?? string.Empty; }
        return "";
    }

    public string[] ShowOpenFileMultiselectDialog(object? owner, string? extensions = null, string? initialDirectory = null, string? fileName = null, string? title = null)
    {
        OpenFileDialog openFileDialog = new() { Multiselect = true };
        if (!string.IsNullOrEmpty(fileName)) { openFileDialog.FileName = fileName; }
        if (!string.IsNullOrEmpty(extensions)) { openFileDialog.Filter = extensions; }
        if (!string.IsNullOrEmpty(title)) { openFileDialog.Title = title; }
        if (!string.IsNullOrEmpty(initialDirectory) && Directory.Exists(initialDirectory)) { openFileDialog.InitialDirectory = initialDirectory; }
        if (openFileDialog.ShowDialog(GetParentWindow(owner)) == true) { return openFileDialog.FileNames; }
        return [];
    }

    public string ShowSaveFileDialog(object? owner, string? extensions = null, string? initialDirectory = null, string? fileName = null, string? title = null)
    {
        SaveFileDialog saveFileDialog = new();
        if (!string.IsNullOrEmpty(fileName)) { saveFileDialog.FileName = fileName; }
        if (!string.IsNullOrEmpty(extensions)) { saveFileDialog.Filter = extensions; }
        if (!string.IsNullOrEmpty(title)) { saveFileDialog.Title = title; }
        if (!string.IsNullOrEmpty(initialDirectory) && Directory.Exists(initialDirectory)) { saveFileDialog.InitialDirectory = initialDirectory; }
        if (saveFileDialog.ShowDialog(GetParentWindow(owner)) == true) { return saveFileDialog.FileName ?? ""; }
        return "";
    }

    public string ShowOpenFolderDialog(object? owner, string? initialDirectory = null, string? title = null)
    {
#if NET8_0_OR_GREATER
        OpenFolderDialog openFolderDialog = new();
        if (!string.IsNullOrEmpty(title)) { openFolderDialog.Title = title; }
        if (!string.IsNullOrEmpty(initialDirectory) && Directory.Exists(initialDirectory)) { openFolderDialog.InitialDirectory = initialDirectory; }
        if (openFolderDialog.ShowDialog() == true) { return openFolderDialog.FolderName ?? ""; }
        return "";
#else
        throw new NotImplementedException();
#endif
    }

    public async Task ShowDialogAsync<TViewModel>(object? owner, params object[] parameters) where TViewModel : IViewModel
    {
        TViewModel viewModel = await GetOrCreateViewModel<TViewModel>(parameters);
        IView view = viewFactory.GetOrCreateView(viewModel);
        if (view is Window window)
        {
            window.Owner = GetParentWindow(owner);
            window.ShowDialog();
            Win_Closed(window, EventArgs.Empty);
        }
        else
        {
            await DisposeViewModel(viewModel);
            throw new InvalidOperationException($"The view model '{typeof(TViewModel).Name}' must be for a window!");
        }
    }

    public async Task<TResult?> ShowDialogAsync<TViewModel, TResult>(object? owner, params object[] parameters) where TViewModel : IViewModel
    {
        TResult? result = default;
        TViewModel viewModel = await GetOrCreateViewModel<TViewModel>(parameters);
        IView view = viewFactory.GetOrCreateView(viewModel);
        if (view is Window window)
        {
            window.Owner = GetParentWindow(owner);
            window.ShowDialog();
            if (viewModel is IResultViewModel<TResult> resultViewModel)
            {
                result = resultViewModel.Result;
            }
            Win_Closed(window, EventArgs.Empty);
            return result;
        }
        else
        {
            await DisposeViewModel(viewModel);
            throw new InvalidOperationException($"The view model '{typeof(TViewModel).Name}' must be for a window!");
        }
    }

    public async Task<TViewModel> ShowWindowAsync<TViewModel>(object? owner, params object[] parameters) where TViewModel : IViewModel
    {
        TViewModel viewModel = await GetOrCreateViewModel<TViewModel>(parameters);
        IView view = viewFactory.GetOrCreateView(viewModel);
        if (view is Window window)
        {
            window.Owner = GetParentWindow(owner);
            window.Closed += Win_Closed;
            window.Show();
        }
        else
        {
            await DisposeViewModel(viewModel);
            throw new InvalidOperationException($"The view model '{typeof(TViewModel).Name}' must be for a window!");
        }
        return viewModel;
    }

    public async Task<(TViewModel ViewModel, Task<TResult?> ResultTask)> ShowWindowAsync<TViewModel, TResult>(object? owner, params object[] parameters) where TViewModel : IViewModel
    {
        TaskCompletionSource<TResult?> tcs = new();
        TViewModel viewModel = await GetOrCreateViewModel<TViewModel>(parameters);
        IView view = viewFactory.GetOrCreateView(viewModel);
        if (view is Window window)
        {
            window.Owner = GetParentWindow(owner);
            window.Closed += OnClosed;
            window.Show();
        }
        else
        {
            await DisposeViewModel(viewModel);
            throw new InvalidOperationException($"The view model '{typeof(TViewModel).Name}' must be for a window!");
        }
        return (viewModel, tcs.Task);
        void OnClosed(object? sender, EventArgs e)
        {
            TResult? result = default;
            if (viewModel is IResultViewModel<TResult> resultViewModel)
            {
                result = resultViewModel.Result;
            }
            window.Closed -= OnClosed;
            Win_Closed(sender, e);
            tcs.SetResult(result);
        }
    }

    public void CloseWindow(IViewModel viewModel)
    {
        if (viewFactory.GetViewOrDefault(viewModel) is Window window)
        {
            window.Close();
        }
    }

    private async Task<TViewModel> GetOrCreateViewModel<TViewModel>(params object[] parameters) where TViewModel : IViewModel
    {
        TViewModel viewModel = ActivatorUtilities.CreateInstance<TViewModel>(serviceProvider, parameters);
        if (viewModel is IWithAsyncRunMethod viewModelWithAsyncRunMethod)
        {
            await viewModelWithAsyncRunMethod.RunAsync();
        }
        if (viewModel is IWithRunMethod viewModelWithRunMethod)
        {
            viewModelWithRunMethod.Run();
        }
        return viewModel;
    }

    private static async Task DisposeViewModel(IViewModel viewModel)
    {
        if (viewModel is IAsyncDisposable asyncDisposable)
        {
            await asyncDisposable.DisposeAsync();
        }
        else if (viewModel is IDisposable disposable)
        {
            disposable.Dispose();
        }
    }

    private static async void Win_Closed(object? sender, EventArgs e)
    {
        if (sender is Window window)
        {
            window.Closed -= Win_Closed;
            if (window.DataContext is IOnWindowClosedEvent onClosedEventImplementation)
            {
                onClosedEventImplementation.OnClosed();
            }
            BindingOperations.ClearAllBindings(window);
            if (window.DataContext is IAsyncDisposable asyncDisposable)
            {
                await asyncDisposable.DisposeAsync();
            }
            else if (window.DataContext is IDisposable disposable)
            {
                disposable.Dispose();
            }
            window.DataContext = null;
            window.Owner?.Activate();
        }
    }

    public Window GetParentWindow(object? instance)
    {
        if (instance is IViewModel vm)
        {
            instance = viewFactory.GetViewOrDefault(vm);
        }
        return FindParentWindow(instance as DependencyObject);
    }

    public static Window FindParentWindow(DependencyObject? view)
    {
        while (view is not null and not Window)
        {
            view = VisualTreeHelper.GetParent(view);
        }
        return (view as Window) ?? Application.Current.MainWindow;
    }

}
