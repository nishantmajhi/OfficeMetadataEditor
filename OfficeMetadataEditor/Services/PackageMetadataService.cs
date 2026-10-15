using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeMetadataEditor.Models;

namespace OfficeMetadataEditor.Services;

/// <summary>
/// .docx, .xlsx and .pptx are all OPC (Open Packaging Conventions) containers - plain
/// zip files with a docProps/core.xml part holding creator, last-modified-by, revision,
/// created and modified. System.IO.Packaging (which comes with the Windows Desktop SDK
/// this app already targets, not a separate NuGet package) reads and writes that part
/// directly through Package.PackageProperties, which is why one code path covers all
/// three formats.
///
/// Saving does two things: it writes the five fields the user edited, and it scrubs
/// everything else the package carries - the remaining core properties (title, subject,
/// keywords, etc.), docProps/app.xml's identifying fields (company, manager, hyperlink
/// base, template path), the docProps/custom.xml part entirely, and any embedded
/// thumbnail preview. That mirrors what Word/Excel/PowerPoint's own "Inspect Document ->
/// Remove Properties and Personal Information" does, without touching document content
/// (text, comments, tracked changes) - which needs format-specific handling this app
/// doesn't attempt.
///
/// Every edit is applied to a copy and swapped in at the end. The input to this app is
/// somebody's own documents, often the only copy, so a save has to be all-or-nothing:
/// see Save below for why doing it in place was not.
/// </summary>
public sealed class PackageMetadataService : IMetadataService
{
    private static readonly HashSet<string> SupportedExtensions =
        new(StringComparer.OrdinalIgnoreCase) { ".docx", ".xlsx", ".pptx" };

    private const string ExtendedPropertiesRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    private const string CustomPropertiesRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    private const string ThumbnailRelationshipType =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    private static readonly XNamespace ExtendedPropertiesNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    public bool IsSupported(string filePath) =>
        SupportedExtensions.Contains(Path.GetExtension(filePath));

    public DocumentMetadata Load(string filePath)
    {
        // FileShare.ReadWrite, not the FileShare.None that Package.Open's path-based
        // overload uses. Reading someone's metadata is a read; a word processor holding
        // the document open for writing is no reason to refuse to show it. Delete is in
        // there so a file being replaced by another program while we read it - which is
        // exactly what this app's own Save does - is not blocked by us either.
        using var stream = new FileStream(
            filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

        EnsureOpenPackagingFormat(stream, filePath);

        using var package = Package.Open(stream, FileMode.Open, FileAccess.Read);
        var props = package.PackageProperties;

        return new DocumentMetadata
        {
            Creator = props.Creator,
            LastModifiedBy = props.LastModifiedBy,
            Revision = props.Revision,
            Created = props.Created,
            Modified = props.Modified
        };
    }

    /// <summary>
    /// Copy, edit the copy, swap it in.
    ///
    /// Editing the original in place meant that any failure part-way left a
    /// half-applied save on disk: the core properties are written first, the extended
    /// ones second, and the custom-properties part and thumbnail last, so a document
    /// whose app.xml will not parse came back with its author changed and its title
    /// gone, still carrying the custom properties the user was told would be removed -
    /// while the window said the save had failed. Package.Dispose() flushes on the way
    /// out of a failed call just as readily as a successful one.
    ///
    /// File.Replace is atomic on the same volume, and the temp file is created in the
    /// document's own directory so it always is one. It also preserves the destination's
    /// attributes, ACLs and creation time, which a delete-and-rename would not.
    /// </summary>
    public void Save(string filePath, DocumentMetadata metadata)
    {
        var target = Path.GetFullPath(filePath);
        var directory = Path.GetDirectoryName(target)
                        ?? throw new IOException($"'{filePath}' has no containing directory.");
        var temporary = Path.Combine(directory, $"{Path.GetFileName(target)}.{Guid.NewGuid():N}.tmp");

        try
        {
            // FileShare.None here, deliberately: the fail-fast the old code relied on.
            // If the document is open in Word/Excel/PowerPoint this throws before any
            // work is done, rather than after a copy has been made and edited.
            using (var source = new FileStream(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                EnsureOpenPackagingFormat(source, filePath);
                using var copy = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                source.CopyTo(copy);
            }

            ApplyTo(temporary, metadata);

            File.Replace(temporary, target, destinationBackupFileName: null);
        }
        finally
        {
            // On success File.Replace has consumed it; on any failure it has not.
            if (File.Exists(temporary))
            {
                try { File.Delete(temporary); }
                catch (IOException) { /* nothing useful to do, and the original is intact */ }
                catch (UnauthorizedAccessException) { /* ditto */ }
            }
        }
    }

    /// <summary>Writes the five editable fields and scrubs the rest. Always runs against
    /// the temporary copy, never the user's file.</summary>
    private static void ApplyTo(string packagePath, DocumentMetadata metadata)
    {
        using var package = Package.Open(packagePath, FileMode.Open, FileAccess.ReadWrite);
        var props = package.PackageProperties;

        // The five fields the app actually lets the user edit.
        props.Creator = NullIfEmpty(metadata.Creator);
        props.LastModifiedBy = NullIfEmpty(metadata.LastModifiedBy);
        props.Revision = NullIfEmpty(metadata.Revision);
        props.Created = metadata.Created;
        props.Modified = metadata.Modified;

        // Everything else in docProps/core.xml gets cleared out.
        props.Title = null;
        props.Subject = null;
        props.Keywords = null;
        props.Description = null;
        props.Category = null;
        props.ContentStatus = null;
        props.Identifier = null;
        props.Language = null;
        props.Version = null;
        props.LastPrinted = null;

        ScrubExtendedProperties(package);
        RemovePartByRelationshipType(package, CustomPropertiesRelationshipType);
        RemovePartByRelationshipType(package, ThumbnailRelationshipType);

        // Package.Dispose() (via using) flushes every changed part back into the zip.
    }

    /// <summary>
    /// Says what is wrong with a file that is not an OPC package, before the zip reader
    /// gets to say "End of Central Directory record could not be found" - which is true
    /// and tells the user nothing. The common case by far is a pre-2007 .doc/.xls/.ppt
    /// that somebody renamed: those are OLE2 compound files and start with a fixed
    /// eight-byte signature.
    /// </summary>
    private static void EnsureOpenPackagingFormat(Stream stream, string filePath)
    {
        Span<byte> header = stackalloc byte[8];
        var start = stream.Position;
        var read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
        stream.Position = start;

        ReadOnlySpan<byte> ole = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        if (read >= ole.Length && header.SequenceEqual(ole))
            throw new InvalidDataException(
                $"'{Path.GetFileName(filePath)}' is a legacy Office document - the pre-2007 binary " +
                "format used by .doc, .xls and .ppt - saved under a modern extension. Only Office " +
                "Open XML packages (.docx, .xlsx, .pptx) store their properties in a form this app " +
                "can read. Open it in Word, Excel or PowerPoint and save it in the newer format first.");

        if (read < 2 || header[0] != (byte)'P' || header[1] != (byte)'K')
            throw new InvalidDataException(
                $"'{Path.GetFileName(filePath)}' is not an Office Open XML package. Files with a " +
                ".docx, .xlsx or .pptx extension are zip containers, and this one is not - it is " +
                "either damaged or not the kind of file its name says it is.");
    }

    /// <summary>
    /// docProps/app.xml isn't exposed through PackageProperties, so it needs direct XML
    /// editing. Only the fields that can carry identifying information are cleared -
    /// Company, Manager, HyperlinkBase (often a local file path), Template, and TotalTime
    /// (cumulative editing time). The part itself is kept, since removing it entirely can
    /// make some Office versions flag the file for repair.
    /// </summary>
    private static void ScrubExtendedProperties(Package package)
    {
        var relationship = package.GetRelationshipsByType(ExtendedPropertiesRelationshipType)
            .FirstOrDefault();
        if (relationship is null) return;

        var partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
        if (!package.PartExists(partUri)) return;

        var part = package.GetPart(partUri);

        XDocument doc;
        using (var readStream = part.GetStream(FileMode.Open, FileAccess.Read))
        {
            doc = XDocument.Load(readStream);
        }

        if (doc.Root is null) return;

        foreach (var name in new[] { "Company", "Manager", "HyperlinkBase", "Template" })
        {
            doc.Root.Element(ExtendedPropertiesNamespace + name)?.Remove();
        }

        var totalTime = doc.Root.Element(ExtendedPropertiesNamespace + "TotalTime");
        if (totalTime is not null) totalTime.Value = "0";

        using var writeStream = part.GetStream(FileMode.Create, FileAccess.Write);
        doc.Save(writeStream);
    }

    /// <summary>Removes a part (and its relationship) by relationship type - used for
    /// custom.xml (custom document properties) and the embedded thumbnail preview.</summary>
    private static void RemovePartByRelationshipType(Package package, string relationshipType)
    {
        foreach (var relationship in package.GetRelationshipsByType(relationshipType).ToList())
        {
            var partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
            if (package.PartExists(partUri))
                package.DeletePart(partUri);

            package.DeleteRelationship(relationship.Id);
        }
    }

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
