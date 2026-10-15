namespace OfficeMetadataEditor.Models;

/// <summary>
/// Maps directly to the OPC "core properties" part (docProps/core.xml) that every
/// .docx, .xlsx and .pptx file shares - that's what makes one editor work for all three.
/// </summary>
public sealed class DocumentMetadata
{
    public string? Creator { get; set; }
    public string? LastModifiedBy { get; set; }
    public string? Revision { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }

    public DocumentMetadata Clone() => new()
    {
        Creator = Creator,
        LastModifiedBy = LastModifiedBy,
        Revision = Revision,
        Created = Created,
        Modified = Modified
    };

    public bool Equals(DocumentMetadata? other) =>
        other is not null &&
        Creator == other.Creator &&
        LastModifiedBy == other.LastModifiedBy &&
        Revision == other.Revision &&
        Created == other.Created &&
        Modified == other.Modified;

    /// <summary>
    /// The same document with blank text fields written as absent ones. Saving turns
    /// an empty or whitespace-only box into a missing property, so anything comparing
    /// two of these - dirty tracking, change detection - has to compare what would be
    /// written rather than what happens to be typed.
    /// </summary>
    public DocumentMetadata Normalized() => new()
    {
        Creator = Blank(Creator),
        LastModifiedBy = Blank(LastModifiedBy),
        Revision = Blank(Revision),
        Created = Created,
        Modified = Modified
    };

    private static string? Blank(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    public override bool Equals(object? obj) => Equals(obj as DocumentMetadata);

    public override int GetHashCode() => HashCode.Combine(Creator, LastModifiedBy, Revision, Created, Modified);
}
