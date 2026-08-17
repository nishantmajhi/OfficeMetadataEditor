using OfficeMetadataEditor.Models;

namespace OfficeMetadataEditor.Services;

public interface IMetadataService
{
    bool IsSupported(string filePath);

    DocumentMetadata Load(string filePath);

    /// <summary>
    /// Writes the given metadata back into the file's docProps/core.xml part.
    /// Throws <see cref="IOException"/> if the file is locked (e.g. still open in Word/Excel/PowerPoint).
    /// </summary>
    void Save(string filePath, DocumentMetadata metadata);
}
