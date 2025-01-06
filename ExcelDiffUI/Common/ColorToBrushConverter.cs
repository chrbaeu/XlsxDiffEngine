using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace ExcelDiffUI.Common;

public sealed class ColorToBrushConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is System.Drawing.Color color)
        {
            return new SolidColorBrush(Color.FromArgb(color.A, color.R, color.G, color.B));
        }
        if (value is string colorString)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorString));
        }
        return null;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is SolidColorBrush brush)
        {
            return System.Drawing.Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B); ;
        }
        return null;
    }
}
