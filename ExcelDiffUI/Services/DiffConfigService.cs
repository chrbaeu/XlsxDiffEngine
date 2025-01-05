using ExcelDiffUI.Common;
using ExcelDiffUI.Models;
using Microsoft.Extensions.Localization;
using Serilog;
using System.IO;
using System.Text.Json;

namespace ExcelDiffUI.Services;

public class DiffConfigService(
    DiffConfigModel diffConfigModel,
    ColumnInfoService columnInfoService,
    IStringLocalizer<Resources.Resources> localizer
    )
{

    public void Reset()
    {
        DiffConfigModel defaultDiffConfigModel = new()
        {
            RowNumberColumnName = localizer["OptionsAddColumnRowName"],
            WorksheetNameColumnName = localizer["OptionsAddColumnWorksheetName"],
            DocumentNameColumnName = localizer["OptionsAddColumnDocumentName"],
        };
        UpdateDiffConfigModel(defaultDiffConfigModel);
    }

    public async Task<bool> Import(string filePath)
    {
        try
        {
            if (string.IsNullOrEmpty(filePath)) { return false; }
            string json = await File.ReadAllTextAsync(filePath);
            if (string.IsNullOrEmpty(json)) { return false; }
            DiffConfigModel? loadedDiffConfigModel = JsonSerializer.Deserialize<DiffConfigModel>(json);
            if (loadedDiffConfigModel is null) { return false; }
            UpdateDiffConfigModel(loadedDiffConfigModel);
            return true;
        }
        catch (Exception e)
        {
            Log.Error($"Importing options from file '{filePath}' failed!", e);
            return false;
        }
    }

    public async Task<bool> Export(string filePath)
    {
        try
        {
            string json = JsonSerializer.Serialize<DiffConfigModel>(diffConfigModel);
            await File.WriteAllTextAsync(filePath, json);
            return true;
        }
        catch (Exception e)
        {
            Log.Error($"Export options to file '{filePath}' failed!", e);
            return false;
        }
    }

    private void UpdateDiffConfigModel(DiffConfigModel newDiffConfigModel)
    {
        MappingHelper.Map(newDiffConfigModel, diffConfigModel);
        newDiffConfigModel.OldFileConfig.FilePath = diffConfigModel.OldFileConfig.FilePath;
        MappingHelper.Map(newDiffConfigModel.OldFileConfig, diffConfigModel.OldFileConfig);
        newDiffConfigModel.NewFileConfig.FilePath = diffConfigModel.NewFileConfig.FilePath;
        MappingHelper.Map(newDiffConfigModel.NewFileConfig, diffConfigModel.NewFileConfig);
        newDiffConfigModel.OutputFileConfig.FilePath = diffConfigModel.OutputFileConfig.FilePath;
        MappingHelper.Map(newDiffConfigModel.OutputFileConfig, diffConfigModel.OutputFileConfig);
        MappingHelper.Map(newDiffConfigModel.ValueChangedMarkers[0], diffConfigModel.ValueChangedMarkers[0]);
        MappingHelper.Map(newDiffConfigModel.ValueChangedMarkers[1], diffConfigModel.ValueChangedMarkers[1]);
        MappingHelper.Map(newDiffConfigModel.ValueChangedMarkers[2], diffConfigModel.ValueChangedMarkers[2]);
        MappingHelper.Map(newDiffConfigModel.ModificationRules[0], diffConfigModel.ModificationRules[0]);
        MappingHelper.Map(newDiffConfigModel.ModificationRules[1], diffConfigModel.ModificationRules[1]);
        MappingHelper.Map(newDiffConfigModel.ModificationRules[2], diffConfigModel.ModificationRules[2]);
        columnInfoService.LoadColumnsFromConfig(newDiffConfigModel.Columns);
    }

}
