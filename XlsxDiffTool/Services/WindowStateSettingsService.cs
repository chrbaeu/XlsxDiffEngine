using Serilog;
using System.IO;
using System.Text.Json;
using XlsxDiffTool.Common;

namespace XlsxDiffTool.Services;

public class WindowStateSettingsService
{
    private readonly string userSettingsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), nameof(XlsxDiffTool));

    public async Task TryToRestoreWindowStateAsync(WindowStateModel windowStateModel)
    {
        try
        {
            string path = Path.Combine(userSettingsFolder, "WindowState.json");
            string json = await File.ReadAllTextAsync(path);
            if (string.IsNullOrEmpty(json)) { return; }
            WindowStateModel? loadedWindowStateModel = JsonSerializer.Deserialize<WindowStateModel>(json);
            if (loadedWindowStateModel is null) { return; }
            MappingHelper.Map(loadedWindowStateModel, windowStateModel);
        }
        catch (Exception e)
        {
            Log.Error($"Restoring last window state failed!", e);
        }
    }

    public async Task SaveWindowStateAsync(WindowStateModel windowStateModel)
    {
        try
        {
            Directory.CreateDirectory(userSettingsFolder);
            string path = Path.Combine(userSettingsFolder, "WindowState.json");
            string json = JsonSerializer.Serialize<WindowStateModel>(windowStateModel);
            await File.WriteAllTextAsync(path, json);
        }
        catch (Exception e)
        {
            Log.Error($"Saving window state failed!", e);
        }
    }
}
