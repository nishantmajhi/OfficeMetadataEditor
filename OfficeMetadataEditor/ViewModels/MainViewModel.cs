using System.Collections.ObjectModel;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeMetadataEditor.Models;
using OfficeMetadataEditor.Services;

namespace OfficeMetadataEditor.ViewModels;

public enum StatusLevel { Ok, Warning, Error }

public partial class MainViewModel : ObservableObject
{
    private readonly IMetadataService _metadataService;
    private readonly IRecentFilesService _recentFilesService;
    private DocumentMetadata? _originalSnapshot;

    public MainViewModel(IMetadataService metadataService, IRecentFilesService recentFilesService)
    {
        _metadataService = metadataService;
        _recentFilesService = recentFilesService;
        RefreshRecentFiles();
        UpdateStatus();
    }

    // ---- Document state -------------------------------------------------

    [ObservableProperty] private bool isDocumentLoaded;
    [ObservableProperty] private string? filePath;
    [ObservableProperty] private string fileDirectory = string.Empty;
    [ObservableProperty] private string fileName = string.Empty;
    [ObservableProperty] private OfficeFileType fileType = OfficeFileType.Unknown;

    // ---- Editable fields --------------------------------------------------

    [ObservableProperty] private string author = string.Empty;
    [ObservableProperty] private string lastModifiedBy = string.Empty;
    [ObservableProperty] private string revision = string.Empty;
    [ObservableProperty] private DateTime? createdDate;
    [ObservableProperty] private TimeSpan? createdTime;
    [ObservableProperty] private DateTime? modifiedDate;
    [ObservableProperty] private TimeSpan? modifiedTime;

    // ---- Status ---------------------------------------------------------

    [ObservableProperty] private bool isDirty;
    [ObservableProperty] private string statusText = "Ready";
    [ObservableProperty] private StatusLevel statusLevel = StatusLevel.Ok;
    [ObservableProperty] private bool isSaveEnabled;

    public ObservableCollection<string> RecentFiles { get; } = new();

    partial void OnAuthorChanged(string value) => OnFieldChanged();
    partial void OnLastModifiedByChanged(string value) => OnFieldChanged();
    partial void OnRevisionChanged(string value) => OnFieldChanged();
    partial void OnCreatedDateChanged(DateTime? value) => OnFieldChanged();
    partial void OnCreatedTimeChanged(TimeSpan? value) => OnFieldChanged();
    partial void OnModifiedDateChanged(DateTime? value) => OnFieldChanged();
    partial void OnModifiedTimeChanged(TimeSpan? value) => OnFieldChanged();

    private bool _suppressDirtyTracking;

    /// <summary>Set once the user has been warned that the document changed on disk, so
    /// the second press of Save goes through. Cleared by every load, close and save.</summary>
    private bool _overwriteConfirmed;

    /// <summary>
    /// "Dirty" means the fields differ from what was last read or written, not that
    /// somebody touched a box. Setting it on any keystroke meant typing a character and
    /// deleting it again left the window claiming unsaved changes for the rest of the
    /// session, and made the close prompt appear for a document nobody had changed.
    /// The comparison normalises blank to absent, because that is what Save writes.
    /// </summary>
    private void OnFieldChanged()
    {
        if (_suppressDirtyTracking) return;
        IsDirty = _originalSnapshot is null ||
                  !BuildMetadataFromFields().Normalized().Equals(_originalSnapshot.Normalized());
        UpdateStatus();
    }

    // ---- Loading / saving --------------------------------------------------

    public void LoadFile(string path)
    {
        if (!_metadataService.IsSupported(path))
        {
            StatusLevel = StatusLevel.Error;
            StatusText = "Unsupported file type - choose a .docx, .xlsx, or .pptx file";
            return;
        }

        try
        {
            var metadata = _metadataService.Load(path);
            _originalSnapshot = metadata.Clone();

            _suppressDirtyTracking = true;
            FilePath = path;
            FileDirectory = Path.GetDirectoryName(path) ?? string.Empty;
            FileName = Path.GetFileName(path);
            FileType = OfficeFileTypeExtensions.FromExtension(Path.GetExtension(path));

            Author = metadata.Creator ?? string.Empty;
            LastModifiedBy = metadata.LastModifiedBy ?? string.Empty;
            Revision = metadata.Revision ?? string.Empty;
            CreatedDate = metadata.Created?.Date;
            CreatedTime = metadata.Created?.TimeOfDay;
            ModifiedDate = metadata.Modified?.Date;
            ModifiedTime = metadata.Modified?.TimeOfDay;
            _suppressDirtyTracking = false;

            IsDocumentLoaded = true;
            IsDirty = false;
            _overwriteConfirmed = false;
            _recentFilesService.Add(path);
            RefreshRecentFiles();
            UpdateStatus();
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
        {
            // Almost always a recent-list entry for something that has since been moved,
            // renamed or deleted. Leaving it in the list means offering it again next
            // time; IRecentFilesService.Remove exists for this and had no caller.
            _recentFilesService.Remove(path);
            RefreshRecentFiles();
            StatusLevel = StatusLevel.Error;
            StatusText = $"Couldn't open file - it is no longer at {path}";
        }
        catch (UnauthorizedAccessException)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = "Couldn't open file - you don't have permission to read it";
        }
        catch (Exception ex)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = $"Couldn't open file: {ex.Message}";
        }
    }

    [RelayCommand(CanExecute = nameof(CanSave))]
    private void Save()
    {
        if (FilePath is null) return;

        var metadata = BuildMetadataFromFields();
        try
        {
            if (!_overwriteConfirmed && HasChangedOnDisk())
            {
                // The window can sit open for hours. Whatever is in these five boxes was
                // read when the document was opened, so writing them now would silently
                // undo whatever has been written since. One press to be told, a second
                // to go ahead - the status bar is where this app says everything else.
                _overwriteConfirmed = true;
                StatusLevel = StatusLevel.Warning;
                StatusText = "This document changed on disk since you opened it - press Save again to overwrite";
                return;
            }

            _metadataService.Save(FilePath, metadata);
            _originalSnapshot = metadata.Clone();
            _overwriteConfirmed = false;
            IsDirty = false;
            StatusLevel = StatusLevel.Ok;
            StatusText = "Saved";
        }
        catch (IOException)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = "Couldn't save - close the file in Word/Excel/PowerPoint and try again";
        }
        catch (UnauthorizedAccessException)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = "Couldn't save - the file or its folder is read-only, or you don't have permission to change it";
        }
        catch (Exception ex)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = $"Couldn't save: {ex.Message}";
        }
    }

    /// <summary>
    /// True when the document's properties are no longer the ones that were read into
    /// these fields. Re-reading is cheap next to the save that follows, and it asks the
    /// question that matters - "has anything changed" - rather than the one a timestamp
    /// answers, which is "has anything been written".
    /// </summary>
    private bool HasChangedOnDisk()
    {
        if (FilePath is null || _originalSnapshot is null) return false;

        try
        {
            return !_metadataService.Load(FilePath).Normalized().Equals(_originalSnapshot.Normalized());
        }
        catch
        {
            // Unreadable for some other reason: let the save itself fail and report why,
            // rather than turning a real error into a confusing warning about overwriting.
            return false;
        }
    }

    private bool CanSave() =>
        IsDocumentLoaded &&
        !string.IsNullOrWhiteSpace(Author) &&
        !string.IsNullOrWhiteSpace(LastModifiedBy);

    [RelayCommand]
    private void CloseDocument()
    {
        _suppressDirtyTracking = true;

        IsDocumentLoaded = false;
        FilePath = null;
        FileDirectory = string.Empty;
        FileName = string.Empty;
        FileType = OfficeFileType.Unknown;
        Author = string.Empty;
        LastModifiedBy = string.Empty;
        Revision = string.Empty;
        CreatedDate = null;
        CreatedTime = null;
        ModifiedDate = null;
        ModifiedTime = null;
        _originalSnapshot = null;
        _overwriteConfirmed = false;

        _suppressDirtyTracking = false;
        IsDirty = false;
        UpdateStatus();
    }

    private DocumentMetadata BuildMetadataFromFields() => new()
    {
        Creator = Author,
        LastModifiedBy = LastModifiedBy,
        Revision = Revision,
        Created = Combine(CreatedDate, CreatedTime),
        Modified = Combine(ModifiedDate, ModifiedTime)
    };

    private static DateTime? Combine(DateTime? date, TimeSpan? time) =>
        date is null ? null : date.Value.Date + (time ?? TimeSpan.Zero);

    private void RefreshRecentFiles()
    {
        RecentFiles.Clear();
        foreach (var path in _recentFilesService.GetRecent())
            RecentFiles.Add(path);
    }

    public void ClearRecentFiles()
    {
        _recentFilesService.Clear();
        RecentFiles.Clear();
    }

    private void UpdateStatus()
    {
        SaveCommand.NotifyCanExecuteChanged();
        IsSaveEnabled = CanSave();

        if (!IsDocumentLoaded)
        {
            StatusLevel = StatusLevel.Ok;
            StatusText = "Ready";
            return;
        }

        if (string.IsNullOrWhiteSpace(Author) || string.IsNullOrWhiteSpace(LastModifiedBy))
        {
            StatusLevel = StatusLevel.Error;
            StatusText = "Cannot save - Author and Last modified by are required";
        }
        else if (IsDirty)
        {
            StatusLevel = StatusLevel.Warning;
            StatusText = "Unsaved changes";
        }
        else
        {
            StatusLevel = StatusLevel.Ok;
            StatusText = "No unsaved changes";
        }
    }
}
