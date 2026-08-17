using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace OfficeMetadataEditor.Converters;

public sealed class InverseBooleanToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is bool b && b ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is Visibility v && v != Visibility.Visible;
}

public sealed class DirtyToCloseLabelConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is bool dirty && dirty ? "Close without saving" : "Close";

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

public sealed class StatusLevelToBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is ViewModels.StatusLevel level
            ? level switch
            {
                ViewModels.StatusLevel.Ok => System.Windows.Media.Brushes.MediumSeaGreen,
                ViewModels.StatusLevel.Warning => System.Windows.Media.Brushes.DarkOrange,
                ViewModels.StatusLevel.Error => System.Windows.Media.Brushes.IndianRed,
                _ => System.Windows.Media.Brushes.Gray
            }
            : System.Windows.Media.Brushes.Gray;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
