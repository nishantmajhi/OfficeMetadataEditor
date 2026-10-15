using System.IO;
using OfficeMetadataEditor.Models;
using OfficeMetadataEditor.Services;

namespace OfficeMetadataEditor.Tests;

public static class MetadataServiceTests
{
    private static readonly PackageMetadataService Svc = new();

    public static void Run(string dir)
    {
        Supported();
        RoundTripEachFormat(dir);
        ScrubbedFields(dir);
        ScrubbedParts(dir);
        ExtendedProperties(dir);
        ContentIsPreserved(dir);
        TimestampsRoundTrip(dir);
        SaveIsIdempotent(dir);
        MissingOptionalParts(dir);
        FailedSaveLeavesTheFileAlone(dir);
        LoadWhileOpenElsewhere(dir);
        SaveWhileLockedExclusively(dir);
        LegacyBinaryFormats(dir);
        MissingFile(dir);
        FileIdentityIsPreserved(dir);
    }

    private static string Fresh(string dir, string name)
    {
        var d = Path.Combine(dir, name);
        if (Directory.Exists(d)) Directory.Delete(d, true);
        Directory.CreateDirectory(d);
        return d;
    }

    private static DocumentMetadata Edited() => new()
    {
        Creator = "New Author",
        LastModifiedBy = "New Editor",
        Revision = "12",
        Created = new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc),
        Modified = new DateTime(2021, 6, 7, 8, 9, 10, DateTimeKind.Utc)
    };

    // ---------------------------------------------------------------------

    private static void Supported()
    {
        T.Section("IsSupported");
        foreach (var good in new[] { "a.docx", "a.xlsx", "a.pptx", "A.DOCX", "a.XlSx", @"C:\x\y z\report.pptx" })
            T.Ok($"accepts {good}", Svc.IsSupported(good));
        foreach (var bad in new[] { "a.doc", "a.xls", "a.ppt", "a.txt", "noextension", "", "a.docx.txt", @"C:\dir.docx\file.txt" })
            T.Ok($"rejects {(bad == "" ? "(empty)" : bad)}", !Svc.IsSupported(bad));
    }

    private static void RoundTripEachFormat(string dir)
    {
        T.Section("Load and Save round-trip, all three formats");
        var d = Fresh(dir, "roundtrip");

        foreach (var kind in new[] { Kind.Docx, Kind.Xlsx, Kind.Pptx })
        {
            var f = Path.Combine(d, "doc" + Opc.Extension(kind));
            Opc.Build(f, kind);

            var loaded = Svc.Load(f);
            T.Equal($"{kind}: Creator read", "Alice Author", loaded.Creator);
            T.Equal($"{kind}: LastModifiedBy read", "Bob Editor", loaded.LastModifiedBy);
            T.Equal($"{kind}: Revision read", "7", loaded.Revision);

            Svc.Save(f, Edited());

            var again = Svc.Load(f);
            T.Equal($"{kind}: Creator written", "New Author", again.Creator);
            T.Equal($"{kind}: LastModifiedBy written", "New Editor", again.LastModifiedBy);
            T.Equal($"{kind}: Revision written", "12", again.Revision);
        }
    }

    private static void ScrubbedFields(string dir)
    {
        T.Section("Save clears the core properties the user cannot edit");
        var d = Fresh(dir, "scrubfields");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        var before = Opc.ReadText(f, "docProps/core.xml")!;
        T.Ok("fixture really does carry a title", before.Contains("Quarterly Report"));
        T.Ok("fixture really does carry keywords", before.Contains("confidential"));

        Svc.Save(f, Edited());

        var after = Opc.ReadText(f, "docProps/core.xml")!;
        foreach (var gone in new[] { "Quarterly Report", "Finance", "confidential", "Not for circulation",
                                     "Reports", "Draft", "2.1", "en-GB", "DOC-00042" })
            T.Ok($"cleared from core.xml: {gone}", !after.Contains(gone));

        T.Ok("the five editable fields survive", after.Contains("New Author") && after.Contains("New Editor"));
    }

    private static void ScrubbedParts(string dir)
    {
        T.Section("Save removes custom properties and the thumbnail");
        var d = Fresh(dir, "scrubparts");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        T.Ok("fixture has docProps/custom.xml", Opc.HasPart(f, "docProps/custom.xml"));
        T.Ok("fixture has a thumbnail", Opc.HasPart(f, "docProps/thumbnail.jpeg"));

        Svc.Save(f, Edited());

        T.Ok("custom.xml part removed", !Opc.HasPart(f, "docProps/custom.xml"));
        T.Ok("thumbnail part removed", !Opc.HasPart(f, "docProps/thumbnail.jpeg"));

        var rels = Opc.RelationshipTypes(f);
        T.Ok("custom-properties relationship removed", !rels.Contains(Opc.CustomPropertiesRel));
        T.Ok("thumbnail relationship removed", !rels.Contains(Opc.ThumbnailRel));
        T.Ok("the officeDocument relationship is still there",
            rels.Any(r => r.EndsWith("/officeDocument")));
    }

    private static void ExtendedProperties(string dir)
    {
        T.Section("Save scrubs docProps/app.xml without deleting it");
        var d = Fresh(dir, "appxml");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        Svc.Save(f, Edited());

        var app = Opc.ReadText(f, "docProps/app.xml");
        T.Ok("app.xml still present", app is not null);
        foreach (var gone in new[] { "Acme Holdings Ltd", "Carol Manager", @"C:\Users\alice\Documents", "Normal.dotm" })
            T.Ok($"cleared from app.xml: {gone}", !app!.Contains(gone));
        T.Ok("TotalTime zeroed, not removed", app!.Contains("<TotalTime>0</TotalTime>"));
        T.Ok("Pages kept", app.Contains("<Pages>12</Pages>"));
        T.Ok("Words kept", app.Contains("<Words>3000</Words>"));
        T.Ok("AppVersion kept", app.Contains("<AppVersion>16.0000</AppVersion>"));
    }

    private static void ContentIsPreserved(string dir)
    {
        T.Section("Save does not touch the document body");
        var d = Fresh(dir, "content");

        foreach (var kind in new[] { Kind.Docx, Kind.Xlsx, Kind.Pptx })
        {
            var f = Path.Combine(d, "doc" + Opc.Extension(kind));
            Opc.Build(f, kind);
            var mainPart = kind switch
            {
                Kind.Docx => "word/document.xml",
                Kind.Xlsx => "xl/workbook.xml",
                _ => "ppt/presentation.xml"
            };
            var bodyBefore = Opc.ReadPart(f, mainPart)!;

            Svc.Save(f, Edited());

            var bodyAfter = Opc.ReadPart(f, mainPart);
            T.Ok($"{kind}: main part byte-identical", bodyAfter is not null && bodyBefore.SequenceEqual(bodyAfter));
            var media = Opc.ReadPart(f, "media/image1.png");
            T.Ok($"{kind}: embedded binary byte-identical", media is not null && media.SequenceEqual(Opc.MediaBytes));

            // Removing parts means rewriting [Content_Types].xml. If the override for
            // the main part went with them, every Office application would refuse the
            // file - and nothing else in this suite would notice, because Package
            // itself reads it back quite happily.
            var contentTypes = Opc.ReadText(f, "[Content_Types].xml") ?? "";
            T.Ok($"{kind}: the main part is still declared in [Content_Types].xml",
                contentTypes.Contains(mainPart));
            T.Ok($"{kind}: core.xml is still declared in [Content_Types].xml",
                contentTypes.Contains("docProps/core.xml"));
            T.Ok($"{kind}: the removed parts are no longer declared",
                !contentTypes.Contains("docProps/custom.xml"));
        }
    }

    private static void TimestampsRoundTrip(string dir)
    {
        T.Section("Created/Modified survive the trip through the UI's date+time split");
        var d = Fresh(dir, "timestamps");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        // The view model shows a DatePicker and a time box, so it splits the value and
        // recombines it. If that were lossy, every save would shift the timestamps by
        // the machine's UTC offset - which is why this is asserted rather than assumed.
        var loaded = Svc.Load(f);
        static DateTime? Recombine(DateTime? v) => v is null ? null : v.Value.Date + v.Value.TimeOfDay;

        var resaved = new DocumentMetadata
        {
            Creator = loaded.Creator,
            LastModifiedBy = loaded.LastModifiedBy,
            Revision = loaded.Revision,
            Created = Recombine(loaded.Created),
            Modified = Recombine(loaded.Modified)
        };
        Svc.Save(f, resaved);

        var after = Svc.Load(f);
        T.Equal("Created unchanged by a no-op save", loaded.Created, after.Created);
        T.Equal("Modified unchanged by a no-op save", loaded.Modified, after.Modified);

        // And an explicitly-UTC value written by the app comes back as the same instant.
        var utc = new DateTime(2019, 11, 12, 13, 14, 15, DateTimeKind.Utc);
        Svc.Save(f, new DocumentMetadata { Creator = "a", LastModifiedBy = "b", Revision = "1", Created = utc, Modified = utc });
        var utcBack = Svc.Load(f);
        T.Equal("a UTC instant round-trips to the same instant",
            utc, utcBack.Created!.Value.ToUniversalTime());

        // A value the DatePicker produces has Kind=Unspecified. It must be read as wall
        // clock in the user's own zone, not silently treated as UTC.
        var wall = new DateTime(2019, 11, 12, 13, 14, 15, DateTimeKind.Unspecified);
        Svc.Save(f, new DocumentMetadata { Creator = "a", LastModifiedBy = "b", Revision = "1", Created = wall, Modified = wall });
        var wallBack = Svc.Load(f).Created!.Value;
        T.Equal("an unspecified-kind value comes back as the same wall clock",
            wall, DateTime.SpecifyKind(wallBack.ToLocalTime(), DateTimeKind.Unspecified));
    }

    private static void SaveIsIdempotent(string dir)
    {
        T.Section("Saving twice changes nothing the second time");
        var d = Fresh(dir, "idempotent");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        Svc.Save(f, Edited());
        var first = Opc.ReadPart(f, "docProps/core.xml")!;
        var firstApp = Opc.ReadPart(f, "docProps/app.xml")!;

        Svc.Save(f, Edited());
        var second = Opc.ReadPart(f, "docProps/core.xml")!;
        var secondApp = Opc.ReadPart(f, "docProps/app.xml")!;

        T.Ok("core.xml identical on a second save", first.SequenceEqual(second));
        T.Ok("app.xml identical on a second save", firstApp.SequenceEqual(secondApp));
    }

    private static void MissingOptionalParts(string dir)
    {
        T.Section("Documents without app.xml / custom.xml / thumbnail still save");
        var d = Fresh(dir, "minimal");
        var f = Path.Combine(d, "bare.docx");
        Opc.Build(f, app: null, custom: null, thumbnail: false);

        T.DoesNotThrow("save with no optional parts", () => Svc.Save(f, Edited()));
        T.Equal("the five fields still went in", "New Author", Svc.Load(f).Creator);
        T.Ok("no app.xml was invented", !Opc.HasPart(f, "docProps/app.xml"));

        // app.xml present but with none of the elements the scrub looks for.
        var g = Path.Combine(d, "sparse.docx");
        Opc.Build(g, app: "<?xml version=\"1.0\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Pages>1</Pages></Properties>");
        T.DoesNotThrow("save with a sparse app.xml", () => Svc.Save(g, Edited()));
        T.Ok("sparse app.xml keeps its one element", Opc.ReadText(g, "docProps/app.xml")!.Contains("<Pages>1</Pages>"));
    }

    // ---- the headline bug ---------------------------------------------------

    private static void FailedSaveLeavesTheFileAlone(string dir)
    {
        T.Section("A save that fails part-way must leave the document exactly as it was");
        var d = Fresh(dir, "atomic");
        var f = Path.Combine(d, "doc.docx");

        // A document whose app.xml is not well-formed XML. ScrubExtendedProperties
        // throws on it - but only after the core properties have already been
        // rewritten, and before custom.xml and the thumbnail have been removed.
        // Truncated or BOM-mangled app.xml parts do occur in the wild; the point of
        // the test is that ANY mid-save failure must not be half-applied.
        Opc.Build(f, app: "<Properties this is not well formed");

        var bytesBefore = File.ReadAllBytes(f);
        var filesBefore = Directory.GetFiles(d).OrderBy(x => x).ToArray();

        T.Throws<Exception>("the save reports a failure", () => Svc.Save(f, Edited()));

        var bytesAfter = File.ReadAllBytes(f);
        T.Ok("the document is byte-for-byte what it was", bytesBefore.SequenceEqual(bytesAfter));
        T.Equal("the author was not silently changed", "Alice Author", Svc.Load(f).Creator);
        T.Ok("the title was not silently deleted",
            Opc.ReadText(f, "docProps/core.xml")!.Contains("Quarterly Report"));
        T.Ok("the custom properties are still there", Opc.HasPart(f, "docProps/custom.xml"));
        T.Ok("no scratch file was left behind",
            Directory.GetFiles(d).OrderBy(x => x).SequenceEqual(filesBefore));
    }

    private static void LoadWhileOpenElsewhere(string dir)
    {
        T.Section("A document open in Word can still be viewed");
        var d = Fresh(dir, "sharing");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        // What a word processor's own handle looks like: opened for writing, shared
        // for reading. Viewing metadata is a read; it has no business failing here.
        using var held = new FileStream(f, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        T.DoesNotThrow("Load succeeds while another process holds the file open for writing",
            () => Svc.Load(f));
    }

    private static void SaveWhileLockedExclusively(string dir)
    {
        T.Section("A document locked exclusively cannot be saved, and is not damaged");
        var d = Fresh(dir, "locked");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);
        var before = File.ReadAllBytes(f);

        using (var held = new FileStream(f, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            T.Throws<IOException>("Save fails with an IOException", () => Svc.Save(f, Edited()));
        }

        T.Ok("the document is untouched", before.SequenceEqual(File.ReadAllBytes(f)));
        T.Ok("no scratch file was left behind", Directory.GetFiles(d).Length == 1);
    }

    private static void LegacyBinaryFormats(string dir)
    {
        T.Section("A pre-2007 binary document says so instead of talking about zips");
        var d = Fresh(dir, "legacy");
        var f = Path.Combine(d, "old.docx");   // a .doc that somebody renamed

        var bytes = new byte[512];
        Opc.OleMagic.CopyTo(bytes, 0);
        File.WriteAllBytes(f, bytes);

        var ex = T.Throws<Exception>("Load rejects it", () => Svc.Load(f));
        var msg = ex?.Message ?? "";
        T.Ok($"the message explains the format, not the zip parser (got: {msg})",
            msg.Contains("legacy", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("pre-2007", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("binary", StringComparison.OrdinalIgnoreCase));

        // Plain rubbish should also produce something a person can act on.
        var g = Path.Combine(d, "rubbish.docx");
        File.WriteAllText(g, "this is not an Office file at all");
        var ex2 = T.Throws<Exception>("Load rejects rubbish", () => Svc.Load(g));
        var msg2 = ex2?.Message ?? "";
        T.Ok($"the message mentions the file, not the central directory (got: {msg2})",
            !msg2.Contains("Central Directory", StringComparison.OrdinalIgnoreCase));
    }

    private static void MissingFile(string dir)
    {
        T.Section("A file that is not there");
        var d = Fresh(dir, "missing");
        T.Throws<FileNotFoundException>("Load throws FileNotFoundException",
            () => Svc.Load(Path.Combine(d, "nope.docx")));
        T.Throws<FileNotFoundException>("Save throws FileNotFoundException",
            () => Svc.Save(Path.Combine(d, "nope.docx"), Edited()));
    }

    private static void FileIdentityIsPreserved(string dir)
    {
        T.Section("Save keeps the file's identity, not just its name");
        var d = Fresh(dir, "identity");
        var f = Path.Combine(d, "doc.docx");
        Opc.Build(f);

        File.SetAttributes(f, File.GetAttributes(f) | FileAttributes.Hidden);
        var createdBefore = File.GetCreationTimeUtc(f);
        Thread.Sleep(30);

        Svc.Save(f, Edited());

        T.Equal("creation time preserved", createdBefore, File.GetCreationTimeUtc(f));
        T.Ok("the Hidden attribute survived", File.GetAttributes(f).HasFlag(FileAttributes.Hidden));
        T.Ok("last-write time moved", File.GetLastWriteTimeUtc(f) >= createdBefore);
        T.Ok("exactly one file in the directory", Directory.GetFiles(d).Length == 1);

        File.SetAttributes(f, FileAttributes.Normal);
    }
}
