using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeMetadataEditor.Models;

namespace OfficeMetadataEditor.Services;

/// <summary>
/// .docx, .xlsx and .pptx are all OPC (Open Packaging Conventions) containers - plain
/// zip files with a docProps/core.xml part holding creator, last-modified-by, revision,
/// created and modified. System.IO.Packaging (part of the base class library, not a
/// separate NuGet package) reads and writes that part directly through
/// Package.PackageProperties, which is why one code path covers all three formats.
///
/// Saving does two things: it writes the five fields the user edited, and it scrubs
/// everything else the package carries - the remaining core properties (title, subject,
/// keywords, etc.), docProps/app.xml's identifying fields (company, manager, hyperlink
/// base, template path), the docProps/custom.xml part entirely, and any embedded
/// thumbnail preview. That mirrors what Word/Excel/PowerPoint's own "Inspect Document ->
/// Remove Properties and Personal Information" does, without touching document content
/// (text, comments, tracked changes) - which needs format-specific handling this app
/// doesn't attempt.
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
        using var package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
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

    public void Save(string filePath, DocumentMetadata metadata)
    {
        // ReadWrite + FileShare.None gives a clear failure (IOException) up front if the
        // document is currently open in Word/Excel/PowerPoint, instead of a silent no-op.
        using var package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
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

