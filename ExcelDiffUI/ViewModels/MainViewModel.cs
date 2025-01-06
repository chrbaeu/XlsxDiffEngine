using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelDiffUI.Common;
using ExcelDiffUI.Services;
using Microsoft.Extensions.Localization;
using System.IO;

namespace ExcelDiffUI.ViewModels;

public sealed partial class MainViewModel(
    OldInputSelectorViewModel oldFile,
    NewInputSelectorViewModel newFile,
    ColumnSelectorViewModel columnsConfig,
    OptionsViewModel options,
    OutputSelectorViewModel outputFile,
    ExcelDiffService excelDiffService,
    IDialogService dialogService,
    DiffConfigService diffConfigService,
    IStringLocalizer<Resources.Resources> localizer
    ) : ObservableObject, IViewModel
{
    private readonly string userSettingsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), nameof(ExcelDiffUI));

    public OldInputSelectorViewModel OldFile { get; } = oldFile;

    public NewInputSelectorViewModel NewFile { get; } = newFile;

    public ColumnSelectorViewModel ColumnsConfig { get; } = columnsConfig;

    public OptionsViewModel Options { get; } = options;

    public OutputSelectorViewModel OutputFile { get; } = outputFile;

    [ObservableProperty]
    public partial bool IsBusy { get; set; }


    [RelayCommand]
    private async Task SaveConfig()
    {
        IsBusy = true;
        try
        {
            if (Options.DiffOptions.SaveAndRestoreInputFilePaths is null)
            {
                var result = dialogService.ShowMessageBox(this, localizer["MsgBoxTitleQuestion"], localizer["SaveAndRestoreInputFilePathsMg"], DialogButton.YesNo);
                Options.DiffOptions.SaveAndRestoreInputFilePaths = result switch
                {
                    DialogResult.Yes => true,
                    DialogResult.No => false,
                    _ => null
                };
            }
            Directory.CreateDirectory(Path.Combine(userSettingsFolder, "Configs"));
            var path = dialogService.ShowSaveFileDialog(this, "Config (*.json)|*.json", Path.Combine(userSettingsFolder, "Configs"));
            if (!string.IsNullOrEmpty(path) && !await diffConfigService.Export(path))
            {
                dialogService.ShowMessageBox(this, localizer["MsgBoxTitleError"], localizer.GetString("ConfigExportFailedMsg", path), DialogButton.OK);
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task LoadConfig()
    {
        IsBusy = true;
        try
        {
            Directory.CreateDirectory(Path.Combine(userSettingsFolder, "Configs"));
            var path = dialogService.ShowOpenFileDialog(this, "Config (*.json)|*.json", Path.Combine(userSettingsFolder, "Configs"));
            if (!string.IsNullOrEmpty(path) && !await diffConfigService.Import(path))
            {
                dialogService.ShowMessageBox(this, localizer["MsgBoxTitleError"], localizer.GetString("ConfigImportFailedMsg", path), DialogButton.OK);
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void ResetConfig()
    {
        IsBusy = true;
        try
        {
            diffConfigService.Reset();
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task SaveDiff()
    {
        IsBusy = true;
        try
        {
            if (!await Task.Run(excelDiffService.SaveDiff))
            {
                dialogService.ShowMessageBox(this, localizer["MsgBoxTitleError"], localizer["DiffSaveFailedMsg"], DialogButton.OK);
            }
        }
        catch (Exception e)
        {
            dialogService.ShowMessageBox(this, localizer["MsgBoxTitleError"], localizer["DiffSaveFailedMsg"], DialogButton.OK);
        }
        finally
        {
            IsBusy = false;
        }
    }

}
