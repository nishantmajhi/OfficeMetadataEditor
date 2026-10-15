using OfficeMetadataEditor.Models;

namespace OfficeMetadataEditor.Services;

public interface IMetadataService
{
    /// <summary>True for the extensions this service can open. It is a check on the
    /// name only; whether the bytes are really an Office Open XML package is decided
    /// by <see cref="Load"/>.</summary>
    bool IsSupported(string filePath);

    /// <summary>
    /// Reads the five core properties. Shares the file, so a document open in
    /// Word/Excel/PowerPoint can still be read.
    /// Throws <see cref="FileNotFoundException"/> if it is not there,
    /// <see cref="InvalidDataException"/> if it is not an Office Open XML package.
    /// </summary>
    DocumentMetadata Load(string filePath);

    /// <summary>
    /// Writes the five editable properties into docProps/core.xml and scrubs the rest
    /// of the package's metadata. All or nothing: the edits are applied to a copy and
    /// swapped in, so a failure leaves the document exactly as it was.
    /// Throws <see cref="IOException"/> if the file is locked (e.g. still open in
    /// Word/Excel/PowerPoint), <see cref="UnauthorizedAccessException"/> if it or its
    /// directory cannot be written.
    /// </summary>
    void Save(string filePath, DocumentMetadata metadata);
}
