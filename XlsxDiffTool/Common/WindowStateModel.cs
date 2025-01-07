using CommunityToolkit.Mvvm.ComponentModel;

namespace XlsxDiffTool.Common;

public enum WindowMode
{
    Normal = 0,
    Minimized = 1,
    Maximized = 2
}

public class WindowStateModel : ObservableObject
{
    public double? Width
    {
        get;
        set { if (WindowMode == WindowMode.Normal && value is double d && !double.IsNaN(d)) { field = d; } }
    }
    public double? Height
    {
        get;
        set { if (WindowMode == WindowMode.Normal && value is double d && !double.IsNaN(d)) { field = d; } }
    }

    public WindowMode WindowMode { get; set; } = WindowMode.Normal;

    public Dictionary<string, bool> CustomFlags { get; set; } = [];
}
