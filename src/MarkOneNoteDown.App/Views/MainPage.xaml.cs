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
    private readonly ObservableCollection<NotebookRef> notebooks = new();
    private readonly ObservableCollection<SectionRef> sections = new();
    private readonly ObservableCollection<PageRef> pages = new();
    private readonly IOneNoteClient oneNoteClient = new OneNoteComClient();
    private readonly IPageParser parser = new BasicPageParser();
    private readonly IExportWriter writer = new FileSystemExportWriter();

    public MainPage()
    {
        InitializeComponent();

        NotebooksList.ItemsSource = notebooks;
        SectionsList.ItemsSource = sections;
        PagesList.ItemsSource = pages;

        _ = LoadNotebooksAsync();
    }

    private async Task LoadNotebooksAsync()
    {
        try
        {
            StatusText.Text = "Loading notebooks...";
            notebooks.Clear();
            sections.Clear();
            pages.Clear();

            IReadOnlyList<NotebookRef> result = await oneNoteClient.GetNotebooksAsync(CancellationToken.None);
            foreach (NotebookRef notebook in result)
            {
                notebooks.Add(notebook);
            }

            StatusText.Text = $"Loaded {notebooks.Count} notebooks.";
            Log($"Loaded {notebooks.Count} notebooks.");
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to load notebooks.";
            Log($"Error loading notebooks: {ex.Message}");
        }
    }

    private async void OnNotebookSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        sections.Clear();
        pages.Clear();

        if (NotebooksList.SelectedItem is not NotebookRef notebook)
        {
            return;
        }

        try
        {
            StatusText.Text = $"Loading sections for {notebook.Name}...";
            IReadOnlyList<SectionRef> result = await oneNoteClient.GetSectionsAsync(notebook.Id, CancellationToken.None);
            foreach (SectionRef section in result)
            {
                sections.Add(section);
            }

            StatusText.Text = $"Loaded {sections.Count} sections.";
            Log($"Notebook selected: {notebook.Name}");
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to load sections.";
            Log($"Error loading sections: {ex.Message}");
        }
    }

    private async void OnSectionSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        pages.Clear();

        if (SectionsList.SelectedItem is not SectionRef section)
        {
            return;
        }

        try
        {
            StatusText.Text = $"Loading pages for {section.Name}...";
            IReadOnlyList<PageRef> result = await oneNoteClient.GetPagesAsync(section.Id, CancellationToken.None);
            foreach (PageRef page in result)
            {
                pages.Add(page);
            }

            StatusText.Text = $"Loaded {pages.Count} pages.";
            Log($"Section selected: {section.Name}");
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to load pages.";
            Log($"Error loading pages: {ex.Message}");
        }
    }

    private async void OnRefreshClicked(object sender, RoutedEventArgs e)
    {
        await LoadNotebooksAsync();
    }

    private async void OnChooseOutputClicked(object sender, RoutedEventArgs e)
    {
        StorageFolder? folder = await PickOutputFolderAsync();
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
            StorageFolder? folder = await PickOutputFolderAsync();
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

        IReadOnlyList<PageRef> exportPages = pages.ToList();
        if (exportPages.Count == 0)
        {
            StatusText.Text = "No pages loaded to export.";
            Log("Export aborted: no pages loaded.");
            return;
        }

        var pipeline = new ExportPipeline(oneNoteClient, parser, writer);
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

    private static async Task<StorageFolder?> PickOutputFolderAsync()
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
