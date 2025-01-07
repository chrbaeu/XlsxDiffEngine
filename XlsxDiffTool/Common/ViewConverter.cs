using System.Globalization;
using System.Windows.Data;

namespace XlsxDiffTool.Common;

public sealed class ViewConverter : IValueConverter
{
    public static ViewFactory? ViewFactory { get; set; }

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (ViewFactory != null && value is IViewModel viewModel)
        {
            return ViewFactory.GetOrCreateView(viewModel);
        }
        return value;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
