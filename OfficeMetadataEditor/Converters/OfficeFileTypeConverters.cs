using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using OfficeMetadataEditor.Models;

namespace OfficeMetadataEditor.Converters;

public sealed class FileTypeToBadgeTextConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is OfficeFileType type ? type.BadgeText() : "?";

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

public sealed class FileTypeToAccentBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var hex = value is OfficeFileType type ? type.AccentHex() : "#5B5A57";
        var color = (Color)ColorConverter.ConvertFromString(hex)!;
        return new SolidColorBrush(color);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
