using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Localization;
using System.IO;
using XlsxDiffTool.Common;
using XlsxDiffTool.Models;

namespace XlsxDiffTool.ViewModels;

public class OldInputSelectorViewModel : InputSelectorViewModel
{
    public OldInputSelectorViewModel(
        IDialogService dialogService,
        DiffConfigModel optionsModel,
        IStringLocalizer<Resources.Resources> localizer
        ) : base(dialogService, optionsModel.OldFileConfig)
    {
        Title = localizer["FileInputOldHeader"];
    }
}

public class NewInputSelectorViewModel : InputSelectorViewModel
{
    public NewInputSelectorViewModel(IDialogService dialogService,
        DiffConfigModel optionsModel,
        IStringLocalizer<Resources.Resources> localizer
        ) : base(dialogService, optionsModel.NewFileConfig)
    {
        Title = localizer["FileInputNewHeader"];
    }
}

public abstract partial class InputSelectorViewModel(
    IDialogService dialogService,
    FileConfigModel fileConfig) : ObservableObject, IViewModel
{
    [ObservableProperty]
    public partial string Title { get; set; } = "File:";

    [ObservableProperty]
    public partial bool EnableSaveFileMode { get; set; }

    public FileConfigModel FileConfig { get; } = fileConfig;


    [RelayCommand]
    public void ChooseFile()
    {
        var initialDirectory = Path.GetDirectoryName(FileConfig.FilePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string filePath;
        if (FileConfig.IsFolderConfig)
        {
            filePath = dialogService.ShowOpenFolderDialog(this, initialDirectory);
        }
        else
        {
            filePath = dialogService.ShowOpenFileDialog(this, "Excel (*.xlsx)|*.xlsx", initialDirectory);
        }
        if (!string.IsNullOrEmpty(filePath))
        {
            FileConfig.FilePath = filePath;
        }
    }

}
