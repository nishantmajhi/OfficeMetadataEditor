using System.IO;
using OfficeMetadataEditor.Models;
using OfficeMetadataEditor.Services;
using OfficeMetadataEditor.ViewModels;

namespace OfficeMetadataEditor.Tests;

/// <summary>
/// The view model has no view in it, so it can be driven directly. These fakes stand
/// in for the two services so a test can say "now the save fails with this exception"
/// or "now the file has changed underneath us" without a filesystem.
/// </summary>
public sealed class FakeMetadataService : IMetadataService
{
    public readonly Dictionary<string, DocumentMetadata> Files = new(StringComparer.OrdinalIgnoreCase);
    public readonly List<string> Saves = new();
    public Exception? LoadThrows;
    public Exception? SaveThrows;
    public int LoadCount;

    public bool IsSupported(string filePath) =>
        new[] { ".docx", ".xlsx", ".pptx" }.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase);

    public DocumentMetadata Load(string filePath)
    {
        LoadCount++;
        if (LoadThrows is not null) throw LoadThrows;
        if (!Files.TryGetValue(filePath, out var m)) throw new FileNotFoundException("no such file", filePath);
        return m.Clone();
    }

    public void Save(string filePath, DocumentMetadata metadata)
    {
        if (SaveThrows is not null) throw SaveThrows;
        Files[filePath] = metadata.Clone();
        Saves.Add(filePath);
    }
}

public sealed class FakeRecentFilesService : IRecentFilesService
{
    public readonly List<string> Entries = new();
    public IReadOnlyList<string> GetRecent() => Entries.ToList();
    public void Add(string filePath)
    {
        Entries.RemoveAll(p => string.Equals(p, filePath, StringComparison.OrdinalIgnoreCase));
        Entries.Insert(0, filePath);
    }
    public void Remove(string filePath) =>
        Entries.RemoveAll(p => string.Equals(p, filePath, StringComparison.OrdinalIgnoreCase));
    public void Clear() => Entries.Clear();
}

public static class ViewModelTests
{
    private const string Path1 = @"C:\docs\report.docx";
    private const string Path2 = @"C:\docs\budget.xlsx";

    private static readonly DateTime Created = new(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);
    private static readonly DateTime Modified = new(2025, 1, 2, 3, 4, 5, DateTimeKind.Utc);

    private static (MainViewModel Vm, FakeMetadataService Meta, FakeRecentFilesService Recent) Build(
        DocumentMetadata? contents = null)
    {
        var meta = new FakeMetadataService();
        meta.Files[Path1] = contents ?? new DocumentMetadata
        {
            Creator = "Alice",
            LastModifiedBy = "Bob",
            Revision = "3",
            Created = Created,
            Modified = Modified
        };
        meta.Files[Path2] = new DocumentMetadata { Creator = "Carol", LastModifiedBy = "Dan", Revision = "1" };
        var recent = new FakeRecentFilesService();
        return (new MainViewModel(meta, recent), meta, recent);
    }

    public static void Run()
    {
        Loading();
        UnsupportedFile();
        DirtyTracking();
        RevertingClearsDirty();
        EmptyAndNullAreTheSame();
        SaveHappyPath();
        SaveErrors();
        MissingFileIsForgotten();
        Closing();
        DateAndTimeCombination();
        ChangedOnDisk();
    }

    private static void Loading()
    {
        T.Section("View model: opening a document");
        var (vm, _, recent) = Build();

        vm.LoadFile(Path1);

        T.Ok("the document is loaded", vm.IsDocumentLoaded);
        T.Equal("path", Path1, vm.FilePath);
        T.Equal("file name", "report.docx", vm.FileName);
        T.Equal("directory", @"C:\docs", vm.FileDirectory);
        T.Equal("file type", OfficeFileType.Word, vm.FileType);
        T.Equal("author", "Alice", vm.Author);
        T.Equal("last modified by", "Bob", vm.LastModifiedBy);
        T.Equal("revision", "3", vm.Revision);
        T.Equal("created date", Created.Date, vm.CreatedDate);
        T.Equal("created time", Created.TimeOfDay, vm.CreatedTime);
        T.Ok("not dirty straight after loading", !vm.IsDirty);
        T.Equal("status", "No unsaved changes", vm.StatusText);
        T.Ok("save is available", vm.IsSaveEnabled);
        T.Equal("added to the recent list", Path1, recent.Entries[0]);
        T.Equal("and shown in the recent list", Path1, vm.RecentFiles[0]);

        vm.LoadFile(Path2);
        T.Equal("opening a second file replaces the first", "Carol", vm.Author);
        T.Equal("file type follows the extension", OfficeFileType.Excel, vm.FileType);
        T.Ok("still not dirty", !vm.IsDirty);
        T.Equal("recent list is newest-first", Path2, vm.RecentFiles[0]);
    }

    private static void UnsupportedFile()
    {
        T.Section("View model: a file it cannot open");
        var (vm, meta, recent) = Build();

        vm.LoadFile(@"C:\docs\notes.txt");
        T.Ok("nothing is loaded", !vm.IsDocumentLoaded);
        T.Equal("the status says why", StatusLevel.Error, vm.StatusLevel);
        T.Ok("and names the formats it does take", vm.StatusText.Contains(".docx"));
        T.Equal("it is not added to the recent list", 0, recent.Entries.Count);
        T.Equal("and the file was never opened", 0, meta.LoadCount);
    }

    private static void DirtyTracking()
    {
        T.Section("View model: unsaved-changes tracking");
        var (vm, _, _) = Build();
        vm.LoadFile(Path1);

        vm.Author = "Alice Smith";
        T.Ok("editing marks it dirty", vm.IsDirty);
        T.Equal("and the status says so", "Unsaved changes", vm.StatusText);
        T.Equal("with a warning level", StatusLevel.Warning, vm.StatusLevel);

        vm.Author = "";
        T.Equal("clearing a required field blocks the save", StatusLevel.Error, vm.StatusLevel);
        T.Ok("and the button goes flat", !vm.IsSaveEnabled);
    }

    private static void RevertingClearsDirty()
    {
        T.Section("View model: putting a value back is not a change");
        var (vm, _, _) = Build();
        vm.LoadFile(Path1);

        vm.Author = "Someone Else";
        T.Ok("dirty after the edit", vm.IsDirty);

        vm.Author = "Alice";
        T.Ok("not dirty once the original value is back", !vm.IsDirty);
        T.Equal("and the status agrees", "No unsaved changes", vm.StatusText);

        vm.CreatedDate = new DateTime(1999, 1, 1);
        T.Ok("dirty after changing the date", vm.IsDirty);
        vm.CreatedDate = Created.Date;
        T.Ok("not dirty once the date is back", !vm.IsDirty);

        vm.Revision = "4";
        vm.LastModifiedBy = "Bobby";
        T.Ok("dirty with two fields changed", vm.IsDirty);
        vm.Revision = "3";
        T.Ok("still dirty with one of them changed", vm.IsDirty);
        vm.LastModifiedBy = "Bob";
        T.Ok("clean once both are back", !vm.IsDirty);
    }

    private static void EmptyAndNullAreTheSame()
    {
        T.Section("View model: a blank field equals an absent one");
        // A document with no revision recorded shows an empty box. Typing into it and
        // deleting again has to come back to "no unsaved changes" - the file stores
        // "absent" and "blank" identically, so the app must compare them identically.
        var (vm, _, _) = Build(new DocumentMetadata { Creator = "Alice", LastModifiedBy = "Bob", Revision = null });
        vm.LoadFile(Path1);

        T.Equal("an absent revision shows as blank", "", vm.Revision);
        vm.Revision = "9";
        T.Ok("dirty after typing", vm.IsDirty);
        vm.Revision = "";
        T.Ok("clean again after deleting it", !vm.IsDirty);

        vm.Revision = "   ";
        T.Ok("whitespace only is still blank", !vm.IsDirty);
    }

    private static void SaveHappyPath()
    {
        T.Section("View model: saving");
        var (vm, meta, _) = Build();
        vm.LoadFile(Path1);

        vm.Author = "Alice Smith";
        vm.Revision = "4";
        vm.SaveCommand.Execute(null);

        T.Equal("the service was asked to save once", 1, meta.Saves.Count);
        T.Equal("to the right path", Path1, meta.Saves[0]);
        T.Equal("with the edited author", "Alice Smith", meta.Files[Path1].Creator);
        T.Equal("and the edited revision", "4", meta.Files[Path1].Revision);
        T.Equal("created is carried through", Created, meta.Files[Path1].Created!.Value);
        T.Ok("no longer dirty", !vm.IsDirty);
        T.Equal("and the status says so", "Saved", vm.StatusText);
        T.Equal("at an ok level", StatusLevel.Ok, vm.StatusLevel);

        // Saving establishes a new baseline: editing back to what was on screen
        // before the save must count as a change again.
        vm.Author = "Alice";
        T.Ok("dirty again after a further edit", vm.IsDirty);
        vm.Author = "Alice Smith";
        T.Ok("clean against the newly saved values", !vm.IsDirty);
    }

    private static void SaveErrors()
    {
        T.Section("View model: what the user is told when a save fails");

        var (vm1, meta1, _) = Build();
        vm1.LoadFile(Path1);
        vm1.Author = "x";
        meta1.SaveThrows = new IOException("used by another process");
        vm1.SaveCommand.Execute(null);
        T.Equal("a locked file is reported as a lock", StatusLevel.Error, vm1.StatusLevel);
        T.Ok("and says to close it in Office", vm1.StatusText.Contains("Word/Excel/PowerPoint"));
        T.Ok("the changes are still marked unsaved", vm1.IsDirty);

        var (vm2, meta2, _) = Build();
        vm2.LoadFile(Path1);
        vm2.Author = "x";
        meta2.SaveThrows = new UnauthorizedAccessException(@"Access to the path 'C:\docs\report.docx' is denied.");
        vm2.SaveCommand.Execute(null);
        T.Equal("a permission problem is an error", StatusLevel.Error, vm2.StatusLevel);
        T.Ok($"and is not blamed on Office (got: {vm2.StatusText})",
            !vm2.StatusText.Contains("Word/Excel/PowerPoint"));
        T.Ok($"it mentions permission or read-only (got: {vm2.StatusText})",
            vm2.StatusText.Contains("permission", StringComparison.OrdinalIgnoreCase) ||
            vm2.StatusText.Contains("read-only", StringComparison.OrdinalIgnoreCase));
        T.Ok("the changes are still marked unsaved", vm2.IsDirty);

        var (vm3, meta3, _) = Build();
        vm3.LoadFile(Path1);
        vm3.Author = "x";
        meta3.SaveThrows = new InvalidDataException("that file is not an Office Open XML package");
        vm3.SaveCommand.Execute(null);
        T.Ok("an unexpected failure shows its own message",
            vm3.StatusText.Contains("not an Office Open XML package"));
        T.Ok("and is still marked unsaved", vm3.IsDirty);
    }

    private static void MissingFileIsForgotten()
    {
        T.Section("View model: a recent file that is no longer there");
        var (vm, meta, recent) = Build();
        vm.LoadFile(Path1);
        vm.LoadFile(Path2);
        T.Equal("both are in the recent list", 2, recent.Entries.Count);

        meta.LoadThrows = new FileNotFoundException("gone", Path1);
        vm.LoadFile(Path1);

        T.Equal("the failure is reported", StatusLevel.Error, vm.StatusLevel);
        T.Ok("a missing file is dropped from the recent list",
            !recent.Entries.Contains(Path1, StringComparer.OrdinalIgnoreCase));
        T.Ok("and from the list on screen",
            !vm.RecentFiles.Contains(Path1, StringComparer.OrdinalIgnoreCase));
        T.Equal("the document that was open stays open", Path2, vm.FilePath);
    }

    private static void Closing()
    {
        T.Section("View model: closing a document");
        var (vm, _, _) = Build();
        vm.LoadFile(Path1);
        vm.Author = "edited";

        vm.CloseDocumentCommand.Execute(null);

        T.Ok("nothing is loaded", !vm.IsDocumentLoaded);
        T.Equal("no path", null, vm.FilePath);
        T.Equal("no author", "", vm.Author);
        T.Equal("no created date", null, vm.CreatedDate);
        T.Ok("not dirty", !vm.IsDirty);
        T.Equal("back to Ready", "Ready", vm.StatusText);
        T.Ok("save is unavailable", !vm.IsSaveEnabled);
    }

    private static void DateAndTimeCombination()
    {
        T.Section("View model: the date box and the time box make one timestamp");
        var (vm, meta, _) = Build();
        vm.LoadFile(Path1);

        vm.CreatedDate = new DateTime(2000, 3, 4);
        vm.CreatedTime = new TimeSpan(5, 6, 7);
        vm.SaveCommand.Execute(null);
        T.Equal("date and time are combined",
            new DateTime(2000, 3, 4, 5, 6, 7), meta.Files[Path1].Created!.Value);

        vm.CreatedTime = null;
        vm.SaveCommand.Execute(null);
        T.Equal("a date with no time is midnight",
            new DateTime(2000, 3, 4, 0, 0, 0), meta.Files[Path1].Created!.Value);

        vm.CreatedDate = null;
        vm.SaveCommand.Execute(null);
        T.Equal("no date at all means no timestamp", null, meta.Files[Path1].Created);
    }

    private static void ChangedOnDisk()
    {
        T.Section("View model: the file changed underneath us");
        var (vm, meta, _) = Build();
        vm.LoadFile(Path1);
        vm.Author = "My Edit";

        // Somebody saved the document in Word while this window sat open. The five
        // fields on screen were read before that happened, so saving them now would
        // quietly undo their change.
        meta.Files[Path1] = new DocumentMetadata
        {
            Creator = "Someone Else",
            LastModifiedBy = "Someone Else",
            Revision = "99",
            Created = Created,
            Modified = DateTime.UtcNow
        };

        vm.SaveCommand.Execute(null);
        T.Equal("the first Save does not write", 0, meta.Saves.Count);
        T.Equal("it warns instead", StatusLevel.Warning, vm.StatusLevel);
        T.Ok($"and says the file changed (got: {vm.StatusText})",
            vm.StatusText.Contains("changed", StringComparison.OrdinalIgnoreCase));
        T.Ok("the edit is still pending", vm.IsDirty);

        vm.SaveCommand.Execute(null);
        T.Equal("pressing Save again goes through", 1, meta.Saves.Count);
        T.Equal("with the user's value", "My Edit", meta.Files[Path1].Creator);
        T.Ok("and it is saved", !vm.IsDirty);

        // Opening a different document must re-arm the check.
        var (vm2, meta2, _) = Build();
        vm2.LoadFile(Path1);
        vm2.Author = "x";
        vm2.SaveCommand.Execute(null);
        T.Equal("an unchanged file saves on the first press", 1, meta2.Saves.Count);
    }
}
