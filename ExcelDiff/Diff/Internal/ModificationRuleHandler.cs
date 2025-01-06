using OfficeOpenXml;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelDiffEngine;

internal sealed class ModificationRuleHandler
{
    private readonly List<(Regex regex, ModificationRule)> regexRules = [];

    internal ModificationRuleHandler(IReadOnlyList<ModificationRule> rules, bool ignoreCase)
    {
        RegexOptions options = ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None;
        foreach (ModificationRule rule in rules)
        {
            regexRules.Add((new Regex(rule.RegexPattern, options), rule));
        }
    }

    internal void ApplyRules(ExcelRange excelCell, string columnName, DataKind dataKind)
    {
        foreach ((Regex regex, ModificationRule rule) in regexRules)
        {
            if (regex.IsMatch(columnName)) { ApplyRule(excelCell, rule, dataKind); }
        }
    }

    private static void ApplyRule(ExcelRange excelCell, ModificationRule rule, DataKind dataKind)
    {
        if (!rule.Target.HasFlag(dataKind)) { return; }
        if (rule.Target.HasFlag(DataKind.NonEmpty) && excelCell.Value is null) { return; }
        switch (rule.ModificationKind)
        {
            case ModificationKind.NumberFormat:
                excelCell.Style.Numberformat.Format = rule.Value;
                break;
            case ModificationKind.Multiply:
                excelCell.Value = (double?)excelCell.Value * double.Parse(rule.Value, CultureInfo.InvariantCulture);
                break;
            case ModificationKind.Formula:
                if (rule.Value.Contains("{#}", StringComparison.Ordinal))
                {
                    if (excelCell.Value is string)
                    {
                        excelCell.Formula = rule.Value.Replace("{#}", excelCell.Text, StringComparison.Ordinal);
                    }
                    else
                    {
                        excelCell.Formula = rule.Value.Replace("{#}", (((double?)excelCell.Value) ?? 0).ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
                    }
                }
                else
                {
                    excelCell.Formula = rule.Value;
                }
                excelCell.Calculate();
                break;
            case ModificationKind.RegexReplace:
                excelCell.Value = Regex.Replace(excelCell.Text, rule.Value, rule.AdditionalValue ?? "");
                break;
            default:
                break;
        }
    }
}
