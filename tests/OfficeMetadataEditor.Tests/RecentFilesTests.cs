using System.IO;
using System.Text.Json;
using OfficeMetadataEditor.Services;

namespace OfficeMetadataEditor.Tests;

public static class RecentFilesTests
{
    public static void Run(string dir)
    {
        var d = Path.Combine(dir, "recent");
        if (Directory.Exists(d)) Directory.Delete(d, true);
        Directory.CreateDirectory(d);

        Ordering(d);
        Cap(d);
        CapOnLoad(d);
        Pruning(d);
        CorruptStore(d);
        RemoveAndClear(d);
        UnwritableStore(d);
    }

    private static string Doc(string dir, string name)
    {
        var p = Path.Combine(dir, name);
        File.WriteAllText(p, "x");
        return p;
    }

    private static string Store(string dir, string name) => Path.Combine(dir, name + ".json");

    private static void Ordering(string dir)
    {
        T.Section("Recent files: most recent first, no duplicates");
        var store = Store(dir, "ordering");
        var svc = new JsonRecentFilesService(store);

        var a = Doc(dir, "a.docx");
        var b = Doc(dir, "b.docx");

        svc.Add(a);
        svc.Add(b);
        T.Equal("newest first", b, svc.GetRecent()[0]);
        T.Equal("two entries", 2, svc.GetRecent().Count);

        svc.Add(a);
        T.Equal("re-adding moves to the front", a, svc.GetRecent()[0]);
        T.Equal("re-adding does not duplicate", 2, svc.GetRecent().Count);

        svc.Add(a.ToUpperInvariant());
        T.Equal("a different casing of the same path is the same entry", 2, svc.GetRecent().Count);

        var reopened = new JsonRecentFilesService(store);
        T.Equal("the list survives a restart", 2, reopened.GetRecent().Count);
    }

    private static void Cap(string dir)
    {
        T.Section("Recent files: the list is bounded");
        var store = Store(dir, "cap");
        var svc = new JsonRecentFilesService(store);

        for (var i = 0; i < 12; i++) svc.Add(Doc(dir, $"cap{i}.docx"));
        T.Equal("Add caps the list at five", 5, svc.GetRecent().Count);
        T.Equal("and keeps the newest", Path.Combine(dir, "cap11.docx"), svc.GetRecent()[0]);

        var onDisk = JsonSerializer.Deserialize<List<string>>(File.ReadAllText(store))!;
        T.Equal("the file on disk is capped too", 5, onDisk.Count);
    }

    private static void CapOnLoad(string dir)
    {
        T.Section("Recent files: an oversized store is trimmed when it is read");
        var store = Store(dir, "caponload");

        // The store is a plain JSON array in %APPDATA%; nothing stops it growing
        // (an older build, a hand edit, a sync conflict). Reading it should apply
        // the same bound that writing it does.
        var many = Enumerable.Range(0, 30).Select(i => Doc(dir, $"big{i}.docx")).ToList();
        File.WriteAllText(store, JsonSerializer.Serialize(many));

        var svc = new JsonRecentFilesService(store);
        T.Equal("an oversized store is trimmed on load", 5, svc.GetRecent().Count);
    }

    private static void Pruning(string dir)
    {
        T.Section("Recent files: entries for deleted files are forgotten");
        var store = Store(dir, "pruning");
        var svc = new JsonRecentFilesService(store);

        var keep = Doc(dir, "keep.docx");
        var gone = Doc(dir, "gone.docx");
        svc.Add(keep);
        svc.Add(gone);
        File.Delete(gone);

        T.Equal("a deleted file is not listed", 1, svc.GetRecent().Count);

        // ...and not just hidden: the entry should be gone from the store, otherwise
        // a deleted document's path sits in %APPDATA% forever. That matters here
        // more than in most apps - the whole point of this one is metadata hygiene.
        var onDisk = JsonSerializer.Deserialize<List<string>>(File.ReadAllText(store))!;
        T.Ok("the deleted file's path is not still in the store on disk",
            !onDisk.Any(p => string.Equals(p, gone, StringComparison.OrdinalIgnoreCase)));
    }

    private static void CorruptStore(string dir)
    {
        T.Section("Recent files: a corrupt store does not take the app down");
        var store = Store(dir, "corrupt");

        File.WriteAllText(store, "{ this is not json");
        T.DoesNotThrow("constructing over broken JSON", () => new JsonRecentFilesService(store));
        T.Equal("and the list is simply empty", 0, new JsonRecentFilesService(store).GetRecent().Count);

        File.WriteAllText(store, "\"a string, not an array\"");
        T.DoesNotThrow("constructing over the wrong JSON shape", () => new JsonRecentFilesService(store));

        File.WriteAllText(store, "null");
        T.Equal("a literal null is an empty list", 0, new JsonRecentFilesService(store).GetRecent().Count);

        File.WriteAllText(store, "[]");
        var svc = new JsonRecentFilesService(store);
        var f = Doc(dir, "after-corrupt.docx");
        svc.Add(f);
        T.Equal("and it can be written to again afterwards", 1, svc.GetRecent().Count);
    }

    private static void RemoveAndClear(string dir)
    {
        T.Section("Recent files: Remove and Clear");
        var store = Store(dir, "removeclear");
        var svc = new JsonRecentFilesService(store);

        var a = Doc(dir, "r1.docx");
        var b = Doc(dir, "r2.docx");
        svc.Add(a);
        svc.Add(b);

        svc.Remove(a);
        T.Equal("Remove drops one entry", 1, svc.GetRecent().Count);
        T.Equal("and leaves the other", b, svc.GetRecent()[0]);

        svc.Remove(b.ToUpperInvariant());
        T.Equal("Remove is case-insensitive too", 0, svc.GetRecent().Count);

        svc.Add(a);
        svc.Add(b);
        svc.Clear();
        T.Equal("Clear empties the list", 0, svc.GetRecent().Count);
        T.Equal("and the store on disk", 0,
            JsonSerializer.Deserialize<List<string>>(File.ReadAllText(store))!.Count);
    }

    private static void UnwritableStore(string dir)
    {
        T.Section("Recent files: an unwritable store is not fatal");
        var store = Store(dir, "unwritable");
        File.WriteAllText(store, "[]");
        File.SetAttributes(store, FileAttributes.ReadOnly);
        try
        {
            var svc = new JsonRecentFilesService(store);
            T.DoesNotThrow("Add over a read-only store", () => svc.Add(Doc(dir, "u.docx")));
            T.Equal("and the in-memory list still reflects it", 1, svc.GetRecent().Count);
        }
        finally
        {
            File.SetAttributes(store, FileAttributes.Normal);
        }
    }
}
