namespace OfficeMetadataEditor.Models;

public enum OfficeFileType
{
    Unknown,
    Word,
    Excel,
    PowerPoint
}

public static class OfficeFileTypeExtensions
{
    public static OfficeFileType FromExtension(string extension) => extension.ToLowerInvariant() switch
    {
        ".docx" => OfficeFileType.Word,
        ".xlsx" => OfficeFileType.Excel,
        ".pptx" => OfficeFileType.PowerPoint,
        _ => OfficeFileType.Unknown
    };

    public static string BadgeText(this OfficeFileType type) => type switch
    {
        OfficeFileType.Word => "W",
        OfficeFileType.Excel => "X",
        OfficeFileType.PowerPoint => "P",
        _ => "?"
    };

    /// <summary>Office's own brand colors - mirrors the accent swap in the HTML mockup's applyTheme().</summary>
    public static string AccentHex(this OfficeFileType type) => type switch
    {
        OfficeFileType.Word => "#185ABD",
        OfficeFileType.Excel => "#107C41",
        OfficeFileType.PowerPoint => "#B7472A",
        _ => "#5B5A57"
    };
}
