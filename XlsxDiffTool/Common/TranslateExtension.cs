using Microsoft.Extensions.Localization;
using System.Windows.Markup;

namespace XlsxDiffTool.Common;

public class TranslateExtension : MarkupExtension
{
    public static IStringLocalizer? Localizer { get; set; }

    public string Key { get; set; }

    public TranslateExtension(string key)
    {
        Key = key;
    }

    public override object ProvideValue(IServiceProvider serviceProvider)
    {
        return Localizer?[Key] ?? Key;
    }
}
