using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.IO;
using XlsxDiffTool.Common;
using XlsxDiffTool.Models;

namespace XlsxDiffTool.ViewModels;

public partial class OutputSelectorViewModel(
    IDialogService dialogService,
    DiffConfigModel optionsModel
    ) : ObservableObject, IViewModel
{
    public OutputFileConfigModel FileConfig { get; } = optionsModel.OutputFileConfig;

    [RelayCommand]
    public void ChooseFile()
    {
        var initialDirectory = FileConfig.IsFolderConfig ? FileConfig.FilePath : Path.GetDirectoryName(FileConfig.FilePath);
        initialDirectory ??= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string path = FileConfig.IsFolderConfig
            ? dialogService.ShowOpenFolderDialog(this, initialDirectory)
            : dialogService.ShowSaveFileDialog(this, "Excel (*.xlsx)|*.xlsx", initialDirectory);
        if (!string.IsNullOrEmpty(path))
        {
            FileConfig.FilePath = path;
        }
    }

}
