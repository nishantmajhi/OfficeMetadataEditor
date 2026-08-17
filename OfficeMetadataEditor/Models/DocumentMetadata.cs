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
}
