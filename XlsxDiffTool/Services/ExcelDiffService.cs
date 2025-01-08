using CommunityToolkit.Mvvm.ComponentModel;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using XlsxDiffEngine;
using XlsxDiffEngine.Diff;
using XlsxDiffTool.Models;

namespace XlsxDiffTool.Services;

public sealed partial class ExcelDiffService(
    DiffConfigModel optionsModel,
    PluginService pluginService
    ) : ObservableObject
{
    public bool SaveDiff()
    {
        if (optionsModel.OldFileConfig.IsFolderConfig || optionsModel.NewFileConfig.IsFolderConfig)
        {
        }

        if (!optionsModel.OldFileConfig.IsExisitingFile() || !optionsModel.NewFileConfig.IsExisitingFile() || !optionsModel.OutputFileConfig.IsValidPath()) { return false; }
        Action<ExcelPackage>? prepareAction = x =>
        {
            foreach (var plugin in optionsModel.Plugins)
            {
                pluginService.Plugins.Where(p => p.Name == plugin).FirstOrDefault()?.OnExcelPackageLoading(optionsModel, x);
            }
        };
        var oldFileInfo = new XlsxFileInfo(optionsModel.OldFileConfig.FilePath)
        {
            FromRow = optionsModel.OldFileConfig.StartRow,
            FromColumn = optionsModel.OldFileConfig.StartColumn,
            MergedWorksheetName = string.IsNullOrWhiteSpace(optionsModel.MergedWorksheetName) ? null : optionsModel.MergedWorksheetName,
            PrepareExcelPackageCallback = prepareAction,
        };
        var newFileInfo = new XlsxFileInfo(optionsModel.NewFileConfig.FilePath)
        {
            FromRow = optionsModel.NewFileConfig.StartRow,
            FromColumn = optionsModel.NewFileConfig.StartColumn,
            MergedWorksheetName = string.IsNullOrWhiteSpace(optionsModel.MergedWorksheetName) ? null : optionsModel.MergedWorksheetName,
            PrepareExcelPackageCallback = prepareAction,
        };
        ExcelDiffBuilder builder = new ExcelDiffBuilder()
            .AddFiles(oldFileInfo, newFileInfo);
        if (optionsModel.SkipEmptyRows)
        {
            builder.SetSkipRowRule(SkipRules.SkipEmptyRows);
        }
        builder.SkipUnchangedRows(optionsModel.SkipUnchangedRows);
        builder.AlwaysSetPrimaryKeyColumnValues(optionsModel.AlwaysSetPrimaryKeyColumnValues);
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
            CellStyle cellStyle = new() { BackgroundColor = ColorTranslator.FromHtml(valueChangedMaker.Color) };
            builder.AddValueChangedMarker(valueChangedMaker.MinDeviationInPercent, valueChangedMaker.MinDeviationAbsolute, cellStyle);
        }
        foreach (var modificationRuleModel in optionsModel.ModificationRules)
        {
            if (modificationRuleModel.Target is null) { continue; }
            ModificationRule modificationRule = new(modificationRuleModel.RegexPattern, modificationRuleModel.ModificationKind, modificationRuleModel.Value,
                modificationRuleModel.Target.Value, modificationRuleModel.AdditionalValue);
            builder.AddModificationRules(modificationRule);
        }
        var keyColumns = optionsModel.Columns.Where(x => x.Mode == ColumnMode.Key).Select(x => x.Name).ToArray();
        if (keyColumns.Length > 0)
        {
            builder.SetKeyColumns(keyColumns);
        }
        var secondaryKeyColumns = optionsModel.Columns.Where(x => x.Mode == ColumnMode.SecondaryKey).Select(x => x.Name).ToArray();
        if (secondaryKeyColumns.Length > 0)
        {
            builder.SetSecondaryKeyColumns(secondaryKeyColumns);
        }
        var groupKeyColumns = optionsModel.Columns.Where(x => x.Mode == ColumnMode.GroupKey).Select(x => x.Name).ToArray();
        if (groupKeyColumns.Length > 0)
        {
            builder.SetGroupKeyColumns(groupKeyColumns);
        }
        var columnsToTextCompareOnly = optionsModel.Columns.Where(x => x.Mode == ColumnMode.TextCompare).Select(x => x.Name).ToArray();
        if (columnsToTextCompareOnly.Length > 0)
        {
            builder.SetColumnsToTextCompareOnly(columnsToTextCompareOnly);
        }
        var columnsToIngore = optionsModel.Columns.Where(x => x.Mode == ColumnMode.Ignore).Select(x => x.Name).ToArray();
        if (columnsToIngore.Length > 0)
        {
            builder.SetColumnsToIgnore(columnsToIngore);
        }
        var columnsToOmit = optionsModel.Columns.Where(x => x.Mode == ColumnMode.Omit).Select(x => x.Name).ToArray();
        if (columnsToOmit.Length > 0)
        {
            builder.SetColumnsToOmit(columnsToOmit);
        }
        foreach (var plugin in optionsModel.Plugins)
        {
            pluginService.Plugins.Where(p => p.Name == plugin).FirstOrDefault()?.OnDiffCreation(optionsModel, builder);
        }
        builder.Build(GetOutputFileName(), x =>
        {
            foreach (var plugin in optionsModel.Plugins)
            {
                pluginService.Plugins.Where(p => p.Name == plugin).FirstOrDefault()?.OnExcelPackageSaving(optionsModel, x);
            }
        });
        return true;
    }

    private string GetOutputFileName()
    {
        string fileName = "Diff";
        string? path;
        if (optionsModel.OutputFileConfig.IsFolderConfig)
        {
            path = optionsModel.OutputFileConfig.FilePath;
            if (optionsModel.OldFileConfig.IsValidPath() && optionsModel.NewFileConfig.IsValidPath())
            {
                var oldFileName = Path.GetFileNameWithoutExtension(optionsModel.OldFileConfig.FilePath);
                var newFileName = Path.GetFileNameWithoutExtension(optionsModel.NewFileConfig.FilePath);
                fileName = $"{oldFileName}_vs_{newFileName}";
            }
        }
        else
        {
            path = Path.GetDirectoryName(optionsModel.OldFileConfig.FilePath);
            if (optionsModel.OldFileConfig.IsValidPath())
            {
                fileName = Path.GetFileNameWithoutExtension(optionsModel.OldFileConfig.FilePath);
            }
        }
        if (optionsModel.OutputFileConfig.AddDateTime)
        {
            fileName += $"_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
        }
        return Path.Combine(path, fileName + ".xlsx");
    }

    public List<string> GetColumnNames()
    {
        Action<ExcelPackage>? prepareAction = x =>
        {
            foreach (var plugin in optionsModel.Plugins)
            {
                pluginService.Plugins.Where(p => p.Name == plugin).FirstOrDefault()?.OnExcelPackageLoading(optionsModel, x);
            }
        };
        var oldFileInfo = optionsModel.OldFileConfig.IsExisitingFile() ? new XlsxFileInfo(optionsModel.OldFileConfig.FilePath)
        {
            FromRow = optionsModel.OldFileConfig.StartRow,
            FromColumn = optionsModel.OldFileConfig.StartColumn,
            MergedWorksheetName = string.IsNullOrWhiteSpace(optionsModel.MergedWorksheetName) ? null : optionsModel.MergedWorksheetName,
            PrepareExcelPackageCallback = prepareAction,
        } : null;
        var newFileInfo = optionsModel.NewFileConfig.IsExisitingFile() ? new XlsxFileInfo(optionsModel.NewFileConfig.FilePath)
        {
            FromRow = optionsModel.NewFileConfig.StartRow,
            FromColumn = optionsModel.NewFileConfig.StartColumn,
            MergedWorksheetName = string.IsNullOrWhiteSpace(optionsModel.MergedWorksheetName) ? null : optionsModel.MergedWorksheetName,
            PrepareExcelPackageCallback = prepareAction,
        } : null;
        List<string> columns = [];
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
        return columns.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }
}
