using System.IO;
using System.Text.Json;

namespace OfficeMetadataEditor.Services;

public sealed class JsonRecentFilesService : IRecentFilesService
{
    private const int MaxEntries = 5;

    private readonly string _storePath;
    private List<string> _entries;

    public JsonRecentFilesService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OfficeMetadataEditor");
        Directory.CreateDirectory(dir);
        _storePath = Path.Combine(dir, "recent.json");
        _entries = Load();
    }

    public IReadOnlyList<string> GetRecent() => _entries.Where(File.Exists).ToList();

    public void Add(string filePath)
    {
        _entries.RemoveAll(p => string.Equals(p, filePath, StringComparison.OrdinalIgnoreCase));
        _entries.Insert(0, filePath);
        if (_entries.Count > MaxEntries)
            _entries = _entries.Take(MaxEntries).ToList();
        Save();
    }

    public void Remove(string filePath)
    {
        _entries.RemoveAll(p => string.Equals(p, filePath, StringComparison.OrdinalIgnoreCase));
        Save();
    }

    public void Clear()
    {
        _entries.Clear();
        Save();
    }

    private List<string> Load()
    {
        try
        {
            if (!File.Exists(_storePath)) return new List<string>();
            var json = File.ReadAllText(_storePath);
            return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
        }
        catch
        {
            // Corrupt or unreadable store - start fresh rather than crash the app.
            return new List<string>();
        }
    }

    private void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(_entries, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_storePath, json);
        }
        catch
        {
            // Non-critical - losing the recent-files list is not worth surfacing an error for.
        }
    }
}
