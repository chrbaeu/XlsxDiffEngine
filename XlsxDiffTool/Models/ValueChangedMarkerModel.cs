using CommunityToolkit.Mvvm.ComponentModel;

namespace XlsxDiffTool.Models;

public sealed partial class ValueChangedMarkerModel : ObservableObject
{
    [ObservableProperty]
    public partial double MinDeviationInPercent { get; set; }

    [ObservableProperty]
    public partial double MinDeviationAbsolute { get; set; }

    [ObservableProperty]
    public partial string Color { get; set; }
}
