using Microsoft.Extensions.Configuration;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;
using MarkOneNoteDown.App.Views;

namespace MarkOneNoteDown.App;

public partial class App : Application
{
    public static Window? MainWindow { get; private set; }

    public static AppSettings Settings { get; private set; } = new();

    public App()
    {
        InitializeComponent();
        LoadSettings();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs e)
    {
        MainWindow ??= new Window();
        MainWindow.Title = "MarkOneNoteDown";

        if (MainWindow.Content is not Frame rootFrame)
        {
            rootFrame = new Frame();
            rootFrame.NavigationFailed += OnNavigationFailed;
            MainWindow.Content = rootFrame;
        }

        _ = rootFrame.Navigate(typeof(MainPage), e.Arguments);
        MainWindow.Activate();
    }

    private static void OnNavigationFailed(object sender, NavigationFailedEventArgs e)
        => throw new InvalidOperationException($"Failed to load Page {e.SourcePageType.FullName}");

    private static void LoadSettings()
    {
        IConfigurationRoot config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

        var settings = new AppSettings();
        config.Bind(settings);
        Settings = settings;
    }
}
