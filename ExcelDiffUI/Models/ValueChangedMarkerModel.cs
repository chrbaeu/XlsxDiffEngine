using CommunityToolkit.Mvvm.ComponentModel;
using System.Drawing;

namespace ExcelDiffUI.Models;

public sealed partial class ValueChangedMarkerModel : ObservableObject
{
    [ObservableProperty]
    public partial double MinDeviationInPercent { get; set; }

    [ObservableProperty]
    public partial double MinDeviationAbsolute { get; set; }

    [ObservableProperty]
    public partial Color Color { get; set; }
}
