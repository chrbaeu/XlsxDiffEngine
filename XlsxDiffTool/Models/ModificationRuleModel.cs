using CommunityToolkit.Mvvm.ComponentModel;
using XlsxDiffEngine;

namespace XlsxDiffTool.Models;

public sealed partial class ModificationRuleModel : ObservableObject
{
    [ObservableProperty]
    public partial string Name { get; set; } = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsActive))]
    public partial DataKind? Target { get; set; }

    [ObservableProperty]
    public partial string RegexPattern { get; set; }

    [ObservableProperty]
    public partial ModificationKind ModificationKind { get; set; } = ModificationKind.Formula;

    [ObservableProperty]
    public partial string Value { get; set; }

    [ObservableProperty]
    public partial string? AdditionalValue { get; set; }

    public bool IsActive => Target is not null;

}
