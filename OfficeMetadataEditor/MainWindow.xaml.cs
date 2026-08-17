using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using OfficeMetadataEditor.ViewModels;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace OfficeMetadataEditor;

public partial class MainWindow : Wpf.Ui.Controls.FluentWindow
{
    private readonly MainViewModel _viewModel;

    public MainWindow(MainViewModel viewModel)
    {
        _viewModel = viewModel;
        DataContext = _viewModel;

        InitializeComponent();

        // Applies the current Windows theme immediately and keeps listening for
        // WM_WININICHANGE so the window flips light/dark the moment the user
        // changes it in Windows Settings - no restart needed.
        SystemThemeWatcher.Watch(this);

        Closing += OnClosing;
    }

    private void OnOpenClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Open document",
            Filter = "Office documents (*.docx;*.xlsx;*.pptx)|*.docx;*.xlsx;*.pptx|All files (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) == true)
        {
            _viewModel.LoadFile(dialog.FileName);
        }
    }

    private void OnOpenRecentClick(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.MenuItem { DataContext: string path })
        {
            _viewModel.LoadFile(path);
        }
    }

    private void OnOpenRecentSubmenuOpened(object sender, RoutedEventArgs e)
    {
        var submenu = (System.Windows.Controls.MenuItem)sender;
        submenu.Items.Clear();

        var recent = _viewModel.RecentFiles;

        if (recent.Count == 0)
        {
            submenu.Items.Add(new System.Windows.Controls.MenuItem
            {
                Header = "No recent files",
                IsEnabled = false
            });
            return;
        }

        foreach (var path in recent)
        {
            var item = new System.Windows.Controls.MenuItem { Header = path, DataContext = path };
            item.Click += OnOpenRecentClick;
            submenu.Items.Add(item);
        }

        submenu.Items.Add(new System.Windows.Controls.Separator());

        var clearItem = new System.Windows.Controls.MenuItem { Header = "Clear recently opened" };
        clearItem.Click += OnClearRecentClick;
        submenu.Items.Add(clearItem);
    }

    private async void OnClearRecentClick(object sender, RoutedEventArgs e)
    {
        var box = new Wpf.Ui.Controls.MessageBox
        {
            Title = "Clear recently opened",
            Content = "Remove all files from the recent list? This can't be undone.",
            PrimaryButtonText = "Clear",
            CloseButtonText = "Cancel"
        };

        var result = await box.ShowDialogAsync();
        if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
        {
            _viewModel.ClearRecentFiles();
        }
    }

    private async void OnSaveClick(object sender, RoutedEventArgs e)
    {
        var box = new Wpf.Ui.Controls.MessageBox
        {
            Title = "Save changes",
            Content = "Saving also clears other document metadata (title, keywords, custom properties, thumbnail) to keep the file clean.",
            PrimaryButtonText = "Save",
            CloseButtonText = "Cancel"
        };

        var result = await box.ShowDialogAsync();
        if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
        {
            _viewModel.SaveCommand.Execute(null);
        }
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) =>
        _viewModel.CloseDocumentCommand.Execute(null);

    private void OnExitClick(object sender, RoutedEventArgs e) =>
        System.Windows.Application.Current.Shutdown();

    private async void OnAboutClick(object sender, RoutedEventArgs e)
    {
        var content = new System.Windows.Controls.StackPanel { Margin = new Thickness(4, 4, 4, 12) };
        content.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Version 1.0.0",
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 10)
        });
        content.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "View and edit document properties - author, last modified by, " +
                   "revision number, and created/modified timestamps - for Word, " +
                   "Excel, and PowerPoint files.",
            TextWrapping = TextWrapping.Wrap
        });

        var box = new Wpf.Ui.Controls.MessageBox
        {
            Title = "Office Metadata Editor",
            Content = content,
            CloseButtonText = "Close",
            MinWidth = 420
        };
        await box.ShowDialogAsync();
    }

    private bool _forceClose;

    private async void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!_viewModel.IsDirty || _forceClose) return;

        e.Cancel = true;

        var box = new Wpf.Ui.Controls.MessageBox
        {
            Title = "Unsaved changes",
            Content = "You have unsaved changes. Close anyway?",
            PrimaryButtonText = "Yes",
            CloseButtonText = "No"
        };

        var result = await box.ShowDialogAsync();
        if (result == Wpf.Ui.Controls.MessageBoxResult.Primary)
        {
            _forceClose = true;
            Close();
        }
    }
}
