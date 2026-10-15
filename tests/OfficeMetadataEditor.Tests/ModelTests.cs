using OfficeMetadataEditor.Models;

namespace OfficeMetadataEditor.Tests;

public static class ModelTests
{
    public static void Run()
    {
        FileTypeMapping();
        MetadataEquality();
    }

    private static void FileTypeMapping()
    {
        T.Section("File type classification");
        T.Equal(".docx is Word", OfficeFileType.Word, OfficeFileTypeExtensions.FromExtension(".docx"));
        T.Equal(".xlsx is Excel", OfficeFileType.Excel, OfficeFileTypeExtensions.FromExtension(".xlsx"));
        T.Equal(".pptx is PowerPoint", OfficeFileType.PowerPoint, OfficeFileTypeExtensions.FromExtension(".pptx"));
        T.Equal(".DOCX is Word too", OfficeFileType.Word, OfficeFileTypeExtensions.FromExtension(".DOCX"));
        T.Equal(".doc is not", OfficeFileType.Unknown, OfficeFileTypeExtensions.FromExtension(".doc"));
        T.Equal("nothing is not", OfficeFileType.Unknown, OfficeFileTypeExtensions.FromExtension(""));

        T.Equal("Word badge", "W", OfficeFileType.Word.BadgeText());
        T.Equal("Excel badge", "X", OfficeFileType.Excel.BadgeText());
        T.Equal("PowerPoint badge", "P", OfficeFileType.PowerPoint.BadgeText());
        T.Equal("unknown badge", "?", OfficeFileType.Unknown.BadgeText());

        foreach (var t in Enum.GetValues<OfficeFileType>())
            T.Ok($"{t} accent is a parseable hex colour",
                System.Text.RegularExpressions.Regex.IsMatch(t.AccentHex(), "^#[0-9A-Fa-f]{6}$"));
    }

    private static void MetadataEquality()
    {
        T.Section("DocumentMetadata equality and cloning");
        var a = new DocumentMetadata
        {
            Creator = "Alice",
            LastModifiedBy = "Bob",
            Revision = "3",
            Created = new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc),
            Modified = new DateTime(2024, 2, 2, 0, 0, 0, DateTimeKind.Utc)
        };

        var clone = a.Clone();
        T.Ok("a clone is a different object", !ReferenceEquals(a, clone));
        T.Ok("but equal", a.Equals(clone));
        T.Ok("equality is symmetric", clone.Equals(a));

        clone.Revision = "4";
        T.Ok("and notices a difference", !a.Equals(clone));

        T.Ok("nothing equals null", !a.Equals((DocumentMetadata?)null));

        // The dirty-tracking comparison has to agree with what Save actually writes,
        // and Save turns a blank string into an absent property. So blank and absent
        // must not read as a difference.
        var blank = new DocumentMetadata { Creator = "Alice", LastModifiedBy = "Bob", Revision = "" };
        var absent = new DocumentMetadata { Creator = "Alice", LastModifiedBy = "Bob", Revision = null };
        T.Ok("a blank field and an absent one are the same document",
            blank.Normalized().Equals(absent.Normalized()));

        var spaces = new DocumentMetadata { Creator = "Alice", LastModifiedBy = "Bob", Revision = "   " };
        T.Ok("whitespace counts as blank", spaces.Normalized().Equals(absent.Normalized()));

        var real = new DocumentMetadata { Creator = "Alice", LastModifiedBy = "Bob", Revision = "1" };
        T.Ok("a real value is still a difference", !real.Normalized().Equals(absent.Normalized()));

        // Equals is also reachable through object, which is where a typed-only
        // overload quietly falls back to reference equality.
        object boxed = a.Clone();
        T.Ok("equality survives being boxed", a.Equals(boxed));
        T.Equal("and hash codes agree", a.GetHashCode(), ((DocumentMetadata)boxed).GetHashCode());
    }
}
