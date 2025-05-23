using CommunityToolkit.Mvvm.ComponentModel;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
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
        if (optionsModel.OldFileConfig.IsFolderConfig != optionsModel.NewFileConfig.IsFolderConfig)
        {
            return false;
        }
        else if (!optionsModel.OldFileConfig.IsFolderConfig && !optionsModel.NewFileConfig.IsFolderConfig)
        {
            if (!optionsModel.OldFileConfig.IsExisitingFile() || !optionsModel.NewFileConfig.IsExisitingFile()) { return false; }
        }
        if (!optionsModel.OutputFileConfig.IsValidPath()) { return false; }

        var (oldXlsxFileInfos, newXlsxFileInfos) = GetXlsxFilesInfos();
        var oldDict = oldXlsxFileInfos.ToDictionary(x => x.DocumentName, StringComparer.OrdinalIgnoreCase);
        var newDict = newXlsxFileInfos.ToDictionary(x => x.DocumentName, StringComparer.OrdinalIgnoreCase);
        var keys = oldDict.Keys.Intersect(newDict.Keys, StringComparer.OrdinalIgnoreCase).Order().ToList();
        if (optionsModel.MergeDocuments || (oldXlsxFileInfos.Count == 1 && newXlsxFileInfos.Count == 1))
        {
            if (optionsModel.MergeDocuments)
            {
                if (keys.Count == 0) { return false; }
                SaveDiff(keys.Select(key => (oldDict[key], newDict[key])).ToList(), null);
            }
            else
            {
                SaveDiff([(oldXlsxFileInfos[0], newXlsxFileInfos[0])], null);
            }
        }
        else
        {
            if (keys.Count == 0) { return false; }
            int fileNumber = 1;
            foreach (var key in keys)
            {
                SaveDiff([(oldDict[key], newDict[key])], fileNumber++);
            }
        }
        return true;
    }

    private bool SaveDiff(List<(XlsxFileInfo OldFile, XlsxFileInfo NewFile)> xlsxFileInfos, int? fileNumber)
    {
        ExcelDiffBuilder builder = new();
        foreach (var item in xlsxFileInfos)
        {
            builder.AddFiles(item.OldFile, item.NewFile);
        }
        if (optionsModel.SkipEmptyRows)
        {
            builder.SetSkipRowRule(PredefinedSkipRules.SkipEmptyRows);
        }
        builder.SkipUnchangedRows(optionsModel.SkipUnchangedRows);
        builder.SkipRemovedRows(optionsModel.SkipRemovedRows);
        if (optionsModel.AlwaysSetPrimaryKeyColumnValues)
        {
            builder.SetColumnsToFillWithOldValueIfNoNewValueExists([.. optionsModel.Columns.Where(x => x.Mode == ColumnMode.Key).Select(x => x.Name)]);
        }
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
        builder.IgnoreColumnsNotInBoth(optionsModel.IgnoreColumnsNotInBoth);
        if (optionsModel.AddOldValueComment)
        {
            builder.AddOldValueAsComment(); // TODO prefix
        }
        builder.MergeWorksheets(optionsModel.MergeWorksheets);
        builder.MergeDocuments(optionsModel.MergeDocuments);
        if (!string.IsNullOrEmpty(optionsModel.MergedDocumentName))
        {
            builder.SetMergedDocumentName(optionsModel.MergedDocumentName);
        }
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
        var columnsToIgnore = optionsModel.Columns.Where(x => x.Mode == ColumnMode.Ignore).Select(x => x.Name).ToArray();
        if (columnsToIgnore.Length > 0)
        {
            builder.SetColumnsToIgnore(columnsToIgnore);
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
        string fileName = GetOutputFileName(xlsxFileInfos[0].OldFile, xlsxFileInfos[0].NewFile, fileNumber);
        builder.Build(fileName, x =>
        {
            foreach (var plugin in optionsModel.Plugins)
            {
                pluginService.Plugins.Where(p => p.Name == plugin).FirstOrDefault()?.OnExcelPackageSaving(optionsModel, x);
            }
        });
        return true;
    }

    private string GetOutputFileName(XlsxFileInfo oldXlsxFileInfo, XlsxFileInfo newXlsxFileInfo, int? fileNumber)
    {
        string numberPostfix = fileNumber is int n ? $" {n}" : "";
        string fileName = "Diff" + numberPostfix;
        string path = "";
        if (optionsModel.OutputFileConfig.IsFolderConfig)
        {
            path = optionsModel.OutputFileConfig.FilePath;
            fileName = $"{oldXlsxFileInfo.DocumentName}_vs_{newXlsxFileInfo.DocumentName}";
        }
        else if (optionsModel.OutputFileConfig.IsValidPath() && Path.GetFileNameWithoutExtension(optionsModel.OldFileConfig.FilePath).Length > 0)
        {
            path = Path.GetDirectoryName(optionsModel.OutputFileConfig.FilePath) ?? "";
            fileName = Path.GetFileNameWithoutExtension(optionsModel.OldFileConfig.FilePath) + numberPostfix;
        }
        if (optionsModel.OutputFileConfig.AddDateTime)
        {
            fileName += $"_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
        }
        return Path.Combine(path, fileName + ".xlsx");
    }

    public List<string> GetColumnNames()
    {
        var (oldXlsxFileInfos, newXlsxFileInfos) = GetXlsxFilesInfos();
        List<string> columns = [];
        foreach (XlsxFileInfo xlsxFileInfo in oldXlsxFileInfos.Concat(newXlsxFileInfos))
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

    private (List<XlsxFileInfo> oldFiles, List<XlsxFileInfo> newFiles) GetXlsxFilesInfos()
    {
        List<XlsxFileInfo> oldFiles = [];
        List<XlsxFileInfo> newFiles = [];
        if (optionsModel.OldFileConfig.IsFolderConfig && optionsModel.NewFileConfig.IsFolderConfig)
        {
            oldFiles = new DirectoryInfo(optionsModel.OldFileConfig.FilePath)
                .GetFiles("*.xlsx")
                .Select(x => CreateXlsxFileInfo(optionsModel.OldFileConfig, x.FullName)).ToList();
            newFiles = new DirectoryInfo(optionsModel.NewFileConfig.FilePath)
                .GetFiles("*.xlsx")
                .Select(x => CreateXlsxFileInfo(optionsModel.NewFileConfig, x.FullName)).ToList();
        }
        else
        {
            if (optionsModel.OldFileConfig.IsExisitingFile())
            {
                oldFiles.Add(CreateXlsxFileInfo(optionsModel.OldFileConfig));
            }
            if (optionsModel.NewFileConfig.IsExisitingFile())
            {
                newFiles.Add(CreateXlsxFileInfo(optionsModel.NewFileConfig));
            }
        }
        return (oldFiles, newFiles);
    }

    private XlsxFileInfo CreateXlsxFileInfo(FileConfigModel fileConfigModel, string? path = null)
    {
        path ??= fileConfigModel.FilePath;
        var documentName = string.IsNullOrWhiteSpace(fileConfigModel.FileNameSelectorRegex)
            ? Path.GetFileNameWithoutExtension(path)
            : new Regex(fileConfigModel.FileNameSelectorRegex).Match(Path.GetFileNameWithoutExtension(path)).Value;
        return new(path)
        {
            FromRow = fileConfigModel.StartRow,
            FromColumn = fileConfigModel.StartColumn,
            MergedWorksheetName = string.IsNullOrWhiteSpace(optionsModel.MergedWorksheetName) ? null : optionsModel.MergedWorksheetName,
            PrepareExcelPackageCallback = prepareAction,
            DocumentName = documentName,
        };
        void prepareAction(ExcelPackage x)
        {
            if (!string.IsNullOrEmpty(optionsModel.Script))
            {
                ExcelScriptExecutor.ExecuteScript(x, optionsModel.Script);
            }
            foreach (var plugin in optionsModel.Plugins)
            {
                pluginService.Plugins.Where(p => p.Name == plugin).FirstOrDefault()?.OnExcelPackageLoading(optionsModel, x);
            }
        }
    }

}
