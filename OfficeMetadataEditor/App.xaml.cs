using System.Windows;
using OfficeMetadataEditor.Services;
using OfficeMetadataEditor.ViewModels;

namespace OfficeMetadataEditor;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // The app is small enough that a DI container would be overkill -
        // services and the view model are just composed by hand here.
        IMetadataService metadataService = new PackageMetadataService();
        IRecentFilesService recentFilesService = new JsonRecentFilesService();
        var viewModel = new MainViewModel(metadataService, recentFilesService);

        var mainWindow = new MainWindow(viewModel);
        mainWindow.Show();
    }
}
