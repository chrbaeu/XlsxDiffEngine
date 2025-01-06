using CommunityToolkit.Mvvm.ComponentModel;
using ExcelDiffUI.Common;
using ExcelDiffUI.Models;
using ExcelDiffUI.Services;
using System.Collections.ObjectModel;

namespace ExcelDiffUI.ViewModels;

public partial class OptionsViewModel(
    DiffConfigModel diffOptions,
    PluginService pluginService
    ) : ObservableObject, IViewModel
{
    public DiffConfigModel DiffOptions { get; } = diffOptions;

    public bool ShowPlugins => pluginService.Plugins.Count > 0;

    public ObservableCollection<PluginModel> Plugins { get; } = [
        .. pluginService.Plugins.Select(x => new PluginModel() { DiffOptions = diffOptions, Name = x.Name })
        ];

    public class PluginModel : ObservableObject
    {
        public required DiffConfigModel DiffOptions { get; init; }
        public required string Name { get; init; }
        public bool IsChecked
        {
            get => DiffOptions.Plugins.Contains(Name);
            set
            {
                if (DiffOptions.Plugins.Contains(Name) == value) { return; }
                if (value)
                {
                    DiffOptions.Plugins.Add(Name);
                }
                else
                {
                    DiffOptions.Plugins.Remove(Name);
                }
            }
        }
    }
}


