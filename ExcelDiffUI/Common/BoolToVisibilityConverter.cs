using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelDiffUI.Common;

public sealed class BoolToVisibilityConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        bool valueAsBool = (value as bool?) ?? false;
        if (parameter is string parameterAsString)
        {
            if (parameterAsString.Contains('!'))
            {
                valueAsBool = !valueAsBool;
            }
            if (parameterAsString.Contains("hidden", StringComparison.CurrentCultureIgnoreCase))
            {
                return valueAsBool ? Visibility.Visible : Visibility.Hidden;
            }
        }
        return valueAsBool ? Visibility.Visible : Visibility.Collapsed;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (parameter is string parameterAsString && parameterAsString.Contains('!'))
        {
            return !(value as Visibility? == Visibility.Visible);
        }
        return value as Visibility? == Visibility.Visible;
    }
}
