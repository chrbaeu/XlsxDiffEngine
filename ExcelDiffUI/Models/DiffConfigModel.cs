﻿using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.Drawing;

namespace ExcelDiffUI.Models;

public sealed partial class DiffConfigModel : ObservableObject
{
    public OldFileConfigModel OldFileConfig { get; init; } = new();

    public NewFileConfigModel NewFileConfig { get; init; } = new();

    public OutputFileConfigModel OutputFileConfig { get; init; } = new();

    public ObservableCollection<ColumnInfoModel> Columns { get; init; } = [];

    [ObservableProperty]
    public partial bool AddRowNumberColumn { get; set; }

    [ObservableProperty]
    public partial string RowNumberColumnName { get; set; } = "Row";

    [ObservableProperty]
    public partial bool AddWorksheetNameColumn { get; set; }

    [ObservableProperty]
    public partial string WorksheetNameColumnName { get; set; } = "Worksheet name";

    [ObservableProperty]
    public partial bool AddDocumentNameColumn { get; set; }

    [ObservableProperty]
    public partial string DocumentNameColumnName { get; set; } = "Document name";


    [ObservableProperty]
    public partial bool AutoFitColumns { get; set; } = true;

    [ObservableProperty]
    public partial bool AutoFilterColumns { get; set; } = true;

    [ObservableProperty]
    public partial bool CopyCellFormats { get; set; } = true;

    [ObservableProperty]
    public partial bool CopyCellStyles { get; set; }

    [ObservableProperty]
    public partial bool HideOldColumns { get; set; }

    [ObservableProperty]
    public partial bool AddOldValueComment { get; set; }


    [ObservableProperty]
    public partial bool MergeWorksheets { get; set; }

    [ObservableProperty]
    public partial bool MergeDocuments { get; set; }


    public ObservableCollection<ValueChangedMarkerModel> ValueChangedMarkers { get; init; } = [
        new() { MinDeviationAbsolute = 0.00, MinDeviationInPercent = 0.00, Color = Color.Yellow },
        new() { MinDeviationAbsolute = 0.00, MinDeviationInPercent = 0.10, Color = Color.Orange },
        new() { MinDeviationAbsolute = 0.00, MinDeviationInPercent = 0.20, Color = Color.Red },
        ];

    public ObservableCollection<ModificationRuleModel> ModificationRules { get; init; } = [
        new ModificationRuleModel() { Name = "Rule 1", Value = "={#}" },
        new ModificationRuleModel() { Name = "Rule 2", Value = "={#}"  },
        new ModificationRuleModel() { Name = "Rule 3", Value = "={#}"  },
        ];
}
