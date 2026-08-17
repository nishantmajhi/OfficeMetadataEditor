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

    private void OnFieldChanged()
    {
        if (_suppressDirtyTracking) return;
        IsDirty = true;
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
            _recentFilesService.Add(path);
            RefreshRecentFiles();
            UpdateStatus();
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
            _metadataService.Save(FilePath, metadata);
            _originalSnapshot = metadata.Clone();
            IsDirty = false;
            StatusLevel = StatusLevel.Ok;
            StatusText = "Saved";
        }
        catch (IOException)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = "Couldn't save - close the file in Word/Excel/PowerPoint and try again";
        }
        catch (Exception ex)
        {
            StatusLevel = StatusLevel.Error;
            StatusText = $"Couldn't save: {ex.Message}";
        }
    }

    private bool CanSave() =>
        IsDocumentLoaded &&
        !string.IsNullOrWhiteSpace(Author) &&
        !string.IsNullOrWhiteSpace(LastModifiedBy);

    [RelayCommand]
    private void CloseDocument()
    {
        IsDocumentLoaded = false;
        FilePath = null;
        FileDirectory = string.Empty;
        FileName = string.Empty;
        FileType = OfficeFileType.Unknown;
        _originalSnapshot = null;
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
