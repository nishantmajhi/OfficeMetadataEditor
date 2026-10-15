using System.IO;
using System.Text.Json;

namespace OfficeMetadataEditor.Services;

public sealed class JsonRecentFilesService : IRecentFilesService
{
    private const int MaxEntries = 5;

    private readonly string _storePath;
    private List<string> _entries;

    public JsonRecentFilesService()
        : this(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OfficeMetadataEditor",
            "recent.json"))
    {
    }

    /// <summary>Used by the tests, which need a store that is not the real one in
    /// %APPDATA%. The app itself uses the parameterless constructor.</summary>
    public JsonRecentFilesService(string storePath)
    {
        _storePath = storePath;
        var dir = Path.GetDirectoryName(storePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        _entries = Load();
    }

    /// <summary>
    /// Entries whose file has since been deleted or moved are dropped, and dropped from
    /// the store as well as from the answer. Hiding them and leaving them on disk meant
    /// the paths of documents the user has deleted stayed in %APPDATA% indefinitely -
    /// an odd thing for an app whose whole subject is metadata hygiene, and the reason
    /// its README now says what this file holds.
    /// </summary>
    public IReadOnlyList<string> GetRecent()
    {
        var live = _entries.Where(File.Exists).ToList();
        if (live.Count != _entries.Count)
        {
            _entries = live;
            Save();
        }
        return live.ToList();
    }

    public void Add(string filePath)
    {
        _entries.RemoveAll(p => string.Equals(p, filePath, StringComparison.OrdinalIgnoreCase));
        _entries.Insert(0, filePath);
        Trim();
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

    private void Trim()
    {
        if (_entries.Count > MaxEntries)
            _entries = _entries.Take(MaxEntries).ToList();
    }

    private List<string> Load()
    {
        try
        {
            if (!File.Exists(_storePath)) return new List<string>();
            var json = File.ReadAllText(_storePath);
            var entries = JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();

            // The cap was applied on the way out and not on the way in, so a store that
            // had grown - an older build, a hand edit, a sync conflict - stayed oversized
            // until the next Add. It is a plain JSON array in %APPDATA%; nothing else
            // guarantees its length.
            return entries.Count > MaxEntries ? entries.Take(MaxEntries).ToList() : entries;
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
