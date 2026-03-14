using System;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using MarkOneNoteDown.Core;

namespace MarkOneNoteDown.App.Views;

public partial class MainPage : Page
{
    private readonly ObservableCollection<NotebookRef> notebooks = new();
    private readonly ObservableCollection<SectionRef> sections = new();
    private readonly ObservableCollection<PageRef> pages = new();

    public MainPage()
    {
        InitializeComponent();

        NotebooksList.ItemsSource = notebooks;
        SectionsList.ItemsSource = sections;
        PagesList.ItemsSource = pages;

        LoadDemoData();
    }

    private void LoadDemoData()
    {
        notebooks.Clear();
        sections.Clear();
        pages.Clear();

        notebooks.Add(new NotebookRef("nb-1", "Work"));
        notebooks.Add(new NotebookRef("nb-2", "Personal"));
        notebooks.Add(new NotebookRef("nb-3", "Archive"));

        Log("Loaded demo notebooks.");
    }

    private void OnNotebookSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        sections.Clear();
        pages.Clear();

        if (NotebooksList.SelectedItem is not NotebookRef notebook)
        {
            return;
        }

        sections.Add(new SectionRef("sec-1", $"{notebook.Name} - Plans", notebook.Id));
        sections.Add(new SectionRef("sec-2", $"{notebook.Name} - Notes", notebook.Id));
        sections.Add(new SectionRef("sec-3", $"{notebook.Name} - Ideas", notebook.Id));

        StatusText.Text = $"Selected notebook: {notebook.Name}";
        Log($"Notebook selected: {notebook.Name}");
    }

    private void OnSectionSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        pages.Clear();

        if (SectionsList.SelectedItem is not SectionRef section)
        {
            return;
        }

        pages.Add(new PageRef("page-1", $"{section.Name} / Overview", section.Id));
        pages.Add(new PageRef("page-2", $"{section.Name} / Tasks", section.Id));
        pages.Add(new PageRef("page-3", $"{section.Name} / Logs", section.Id));

        StatusText.Text = $"Selected section: {section.Name}";
        Log($"Section selected: {section.Name}");
    }

    private void OnRefreshClicked(object sender, RoutedEventArgs e)
    {
        LoadDemoData();
        StatusText.Text = "Refresh completed (demo data).";
    }

    private void OnChooseOutputClicked(object sender, RoutedEventArgs e)
    {
        StatusText.Text = "Output folder picker is not wired yet.";
        Log("Output picker not implemented.");
    }

    private void OnExportClicked(object sender, RoutedEventArgs e)
    {
        ExportProgress.IsIndeterminate = true;
        StatusText.Text = "Export started (pipeline not wired).";
        Log($"Export requested. Pages selected: {pages.Count}.");
    }

    private void Log(string message)
    {
        string line = $"{DateTime.Now:HH:mm:ss} {message}";
        LogList.Items.Add(line);
    }
}
