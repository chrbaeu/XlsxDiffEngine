using CommunityToolkit.Mvvm.ComponentModel;
using ExcelDiffEngine;

namespace ExcelDiffUI.Models;

public sealed partial class ModificationRuleModel : ObservableObject
{
    [ObservableProperty]
    public partial string Name { get; set; } = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsAktive))]
    public partial DataKind? Target { get; set; }

    [ObservableProperty]
    public partial string RegexPattern { get; set; }

    [ObservableProperty]
    public partial ModificationKind ModificationKind { get; set; } = ModificationKind.Formula;

    [ObservableProperty]
    public partial string Value { get; set; }

    [ObservableProperty]
    public partial string? AdditionalValue { get; set; }

    public bool IsAktive => Target is not null;

}
