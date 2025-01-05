using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelDiffUI.Common;

public class WindowModeConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is WindowMode viewStateMode)
        {
            return viewStateMode switch
            {
                WindowMode.Maximized => WindowState.Maximized,
                WindowMode.Minimized => WindowState.Minimized,
                _ => WindowState.Normal,
            };
        }
        return WindowState.Normal;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is WindowState windowState)
        {
            return windowState switch
            {
                WindowState.Maximized => WindowMode.Maximized,
                WindowState.Minimized => WindowMode.Minimized,
                _ => WindowMode.Normal,
            };
        }
        return WindowMode.Normal;
    }

}
