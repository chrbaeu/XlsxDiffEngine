using CommunityToolkit.Mvvm.ComponentModel;

namespace ExcelDiffUI.Models;

public enum ColumnMode
{
    Default,
    Key,
    SecondaryKey,
    GroupKey,
    Ignore,
    Omit,
    TextCompare,
}

public enum ColumnKind
{
    Default,
    RowNumber,
    WorksheetName,
    DocumentName,
}

public sealed partial class ColumnInfoModel : ObservableObject
{
    [ObservableProperty]
    public partial string Name { get; set; }

    [ObservableProperty]
    public partial ColumnMode Mode { get; set; } = ColumnMode.Default;

    [ObservableProperty]
    public partial bool IsNotMapped { get; set; }

    [ObservableProperty]
    public partial ColumnKind ColumnKind { get; set; }
}
