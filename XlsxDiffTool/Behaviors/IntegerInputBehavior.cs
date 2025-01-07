
using Microsoft.Xaml.Behaviors;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace XlsxDiffTool.Behaviors;

public sealed partial class IntegerInputBehavior : Behavior<TextBox>
{
    [GeneratedRegex(@"^[0-9]+$")]
    private static partial Regex IsNumberRegex { get; }

    protected override void OnAttached()
    {
        base.OnAttached();
        AssociatedObject.PreviewTextInput += OnPreviewTextInput;
        DataObject.AddPastingHandler(AssociatedObject, OnPasting);
    }

    protected override void OnDetaching()
    {
        base.OnDetaching();
        AssociatedObject.PreviewTextInput -= OnPreviewTextInput;
        DataObject.RemovePastingHandler(AssociatedObject, OnPasting);
    }

    private void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        e.Handled = !IsNumberRegex.IsMatch(e.Text);
    }

    private void OnPasting(object sender, DataObjectPastingEventArgs e)
    {
        if (e.DataObject.GetDataPresent(typeof(string)))
        {
            string text = (string)e.DataObject.GetData(typeof(string));
            if (!IsNumberRegex.IsMatch(text))
            {
                e.CancelCommand();
            }
        }
        else
        {
            e.CancelCommand();
        }
    }

}
