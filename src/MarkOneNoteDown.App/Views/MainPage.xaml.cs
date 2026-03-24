using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using MarkOneNoteDown.Core;
using MarkOneNoteDown.Export;
using MarkOneNoteDown.OneNote;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace MarkOneNoteDown.App.Views;

public partial class MainPage : Page
{
    private readonly ObservableCollection<PageRef> pages = new();
    private readonly IOneNoteClient sourceClient = new HtmlExportClient();
    private readonly IPageParser parser = new HtmlPageParser();
    private readonly IExportWriter writer = new FileSystemExportWriter();

    public MainPage()
    {
        InitializeComponent();

        PagesList.ItemsSource = pages;

        SourcePathBox.Text = App.Settings.SourceFolder ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(SourcePathBox.Text))
        {
            _ = LoadPagesAsync(SourcePathBox.Text);
        }
    }

    private async void OnChooseSourceClicked(object sender, RoutedEventArgs e)
    {
        StorageFolder? folder = await PickFolderAsync();
        if (folder is null)
        {
            Log("Source picker canceled.");
            return;
        }

        SourcePathBox.Text = folder.Path;
        StatusText.Text = $"Source folder: {folder.Path}";
        await LoadPagesAsync(folder.Path);
    }

    private async void OnLoadPagesClicked(object sender, RoutedEventArgs e)
    {
        await LoadPagesAsync(SourcePathBox.Text);
    }

    private async Task LoadPagesAsync(string? sourceFolder)
    {
        pages.Clear();
        string folder = sourceFolder?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(folder))
        {
            StatusText.Text = "Source folder is empty.";
            Log("Source folder is empty.");
            return;
        }

        try
        {
            StatusText.Text = "Loading pages from HTML export...";
            IReadOnlyList<PageRef> result = await sourceClient.GetPagesAsync(folder, CancellationToken.None);
            foreach (PageRef page in result)
            {
                pages.Add(page);
            }

            StatusText.Text = $"Loaded {pages.Count} pages.";
            Log($"Loaded {pages.Count} pages from export folder.");
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to load pages.";
            Log($"Error loading pages: {ex.Message}");
        }
    }

    private async void OnChooseOutputClicked(object sender, RoutedEventArgs e)
    {
        StorageFolder? folder = await PickFolderAsync();
        if (folder is null)
        {
            Log("Output picker canceled.");
            return;
        }

        OutputPathBox.Text = folder.Path;
        StatusText.Text = $"Output folder: {folder.Path}";
    }

    private async void OnExportClicked(object sender, RoutedEventArgs e)
    {
        string outputDirectory = OutputPathBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            StorageFolder? folder = await PickFolderAsync();
            if (folder is null)
            {
                Log("Export canceled (no output directory).");
                return;
            }

            outputDirectory = folder.Path;
            OutputPathBox.Text = outputDirectory;
        }

        ExportProgress.IsIndeterminate = false;
        ExportProgress.Value = 0;

        IReadOnlyList<PageRef> exportPages = GetSelectedOrAllPages();
        if (exportPages.Count == 0)
        {
            StatusText.Text = "No pages loaded to export.";
            Log("Export aborted: no pages loaded.");
            return;
        }

        var pipeline = new ExportPipeline(sourceClient, parser, writer);
        var options = new ExportOptions(
            outputDirectory,
            IncludeAttachmentsCheckBox.IsChecked == true,
            FlattenHierarchyCheckBox.IsChecked == true);

        try
        {
            StatusText.Text = "Exporting...";
            var progress = new Progress<ExportProgress>(p =>
            {
                double value = p.Total > 0 ? (double)p.Completed / p.Total * 100 : 0;
                ExportProgress.Value = value;
                StatusText.Text = $"Exporting {p.CurrentItem} ({p.Completed}/{p.Total})";
            });

            ExportResult result = await pipeline.ExportPagesAsync(exportPages, options, progress, CancellationToken.None);
            StatusText.Text = $"Export complete. Pages: {result.ExportedPages}, Assets: {result.ExportedAssets}";
            Log($"Export completed. Pages: {result.ExportedPages}, Assets: {result.ExportedAssets}");
        }
        catch (Exception ex)
        {
            StatusText.Text = "Export failed.";
            Log($"Export failed: {ex.Message}");
        }
    }

    private IReadOnlyList<PageRef> GetSelectedOrAllPages()
    {
        var selected = PagesList.SelectedItems?.OfType<PageRef>().ToList() ?? new List<PageRef>();
        if (selected.Count > 0)
        {
            Log($"Exporting selected pages: {selected.Count}.");
            return selected;
        }

        Log($"No selection. Exporting all loaded pages: {pages.Count}.");
        return pages.ToList();
    }

    private static async Task<StorageFolder?> PickFolderAsync()
    {
        if (App.MainWindow is null)
        {
            return null;
        }

        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");

        IntPtr hwnd = WindowNative.GetWindowHandle(App.MainWindow);
        InitializeWithWindow.Initialize(picker, hwnd);

        return await picker.PickSingleFolderAsync();
    }

    private void Log(string message)
    {
        string line = $"{DateTime.Now:HH:mm:ss} {message}";
        LogList.Items.Add(line);
    }
}
