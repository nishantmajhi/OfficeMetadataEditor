namespace OfficeMetadataEditor.Services;

public interface IRecentFilesService
{
    IReadOnlyList<string> GetRecent();
    void Add(string filePath);
    void Remove(string filePath);
    void Clear();
}
