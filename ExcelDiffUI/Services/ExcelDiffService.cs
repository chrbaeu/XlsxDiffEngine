using CommunityToolkit.Mvvm.ComponentModel;
using ExcelDiffEngine;
using ExcelDiffUI.Models;
using System.IO;

namespace ExcelDiffUI.Services;

public sealed partial class ExcelDiffService : ObservableObject
{
    private readonly DiffConfigModel optionsModel;
    private readonly OldFileConfigModel oldFileConfig;
    private readonly NewFileConfigModel newFileConfig;
    private readonly OutputFileConfigModel outputFileConfig;

    public ExcelDiffService(
        DiffConfigModel optionsModel
        )
    {
        this.optionsModel = optionsModel;
        this.oldFileConfig = optionsModel.OldFileConfig;
        this.newFileConfig = optionsModel.NewFileConfig;
        this.outputFileConfig = optionsModel.OutputFileConfig;
    }

    public bool SaveDiff()
    {
        if (oldFileConfig.IsFolderConfig || newFileConfig.IsFolderConfig)
        {
        }

        if (!oldFileConfig.IsExisitingFile() || !newFileConfig.IsExisitingFile() || !oldFileConfig.IsValidPath()) { return false; }
        var oldFileInfo = new XlsxFileInfo(oldFileConfig.FilePath) { FromRow = oldFileConfig.StartRow, FromColumn = oldFileConfig.StartColumn };
        var newFileInfo = new XlsxFileInfo(newFileConfig.FilePath) { FromRow = newFileConfig.StartRow, FromColumn = newFileConfig.StartColumn };
        ExcelDiffBuilder builder = new ExcelDiffBuilder()
            .AddFiles(oldFileInfo, newFileInfo);
        if (optionsModel.AddRowNumberColumn)
        {
            builder.AddRowNumberAsColumn(optionsModel.RowNumberColumnName);
        }
        if (optionsModel.AddWorksheetNameColumn)
        {
            builder.AddWorksheetNameAsColumn(optionsModel.WorksheetNameColumnName);
        }
        if (optionsModel.AddDocumentNameColumn)
        {
            builder.AddDocumentNameAsColumn(optionsModel.DocumentNameColumnName);
        }
        builder.SetAutoFitColumns(optionsModel.AutoFitColumns);
        builder.SetAutoFilter(optionsModel.AutoFilterColumns);
        builder.CopyCellFormat(optionsModel.CopyCellFormats);
        builder.CopyCellStyle(optionsModel.CopyCellStyles);
        builder.HideOldColumns(optionsModel.HideOldColumns);
        if (optionsModel.AddOldValueComment)
        {
            builder.AddOldValueAsComment(); // TODO prefix
        }
        builder.MergeWorksheets(optionsModel.MergeWorksheets);
        builder.MergeDocuments(optionsModel.MergeDocuments);
        foreach (var valueChangedMaker in optionsModel.ValueChangedMarkers)
        {
            CellStyle cellStyle = new() { BackgroundColor = valueChangedMaker.Color };
            builder.AddValueChangedMarker(valueChangedMaker.MinDeviationInPercent, valueChangedMaker.MinDeviationAbsolute, cellStyle);
        }
        foreach (var modificationRuleModel in optionsModel.ModificationRules)
        {
            if (modificationRuleModel.Target is null) { continue; }
            ModificationRule modificationRule = new(modificationRuleModel.RegexPattern, modificationRuleModel.ModificationKind, modificationRuleModel.Value,
                modificationRuleModel.Target.Value, modificationRuleModel.AdditionalValue);
            builder.AddModificationRules(modificationRule);
        }
        builder.Build(GetOutputFileName());
        return true;
    }

    private string GetOutputFileName()
    {
        string fileName = "Diff";
        string? path;
        if (outputFileConfig.IsFolderConfig)
        {
            path = outputFileConfig.FilePath;
            if (oldFileConfig.IsValidPath() && newFileConfig.IsValidPath())
            {
                var oldFileName = Path.GetFileNameWithoutExtension(oldFileConfig.FilePath);
                var newFileName = Path.GetFileNameWithoutExtension(newFileConfig.FilePath);
                fileName = $"{oldFileName}_vs_{newFileName}";
            }
        }
        else
        {
            path = Path.GetDirectoryName(oldFileConfig.FilePath);
            if (oldFileConfig.IsValidPath())
            {
                fileName = Path.GetFileNameWithoutExtension(oldFileConfig.FilePath);
            }
        }
        if (outputFileConfig.AddDateTime)
        {
            fileName += $"_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
        }
        return Path.Combine(path, fileName + ".xlsx");
    }

    public List<string> GetColumnNames()
    {
        List<string> columns = [];
        XlsxFileInfo oldFileInfo = oldFileConfig.IsExisitingFile() ? new(oldFileConfig.FilePath) { FromRow = oldFileConfig.StartRow, FromColumn = oldFileConfig.StartColumn } : null;
        XlsxFileInfo newFileInfo = newFileConfig.IsExisitingFile() ? new(newFileConfig.FilePath) { FromRow = newFileConfig.StartRow, FromColumn = newFileConfig.StartColumn } : null;
        foreach (XlsxFileInfo xlsxFileInfo in new XlsxFileInfo?[] { newFileInfo, oldFileInfo }.OfType<XlsxFileInfo>())
        {
            using XlsxDataProvider xlsxDataProvider = new(xlsxFileInfo);
            foreach (IExcelDataSource dataSource in xlsxDataProvider.GetDataSources())
            {
                foreach (string column in dataSource.GetColumnNames())
                {
                    columns.Add(column);
                }
            }
        }
        return columns;
    }
}
