using System.Runtime.InteropServices;
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

        UnhandledException += OnUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
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

    private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        LogFatal("Unhandled UI exception", e.Exception);
        e.Handled = true;
    }

    private static void OnAppDomainUnhandledException(object? sender, System.UnhandledExceptionEventArgs e)
    {
        var ex = e.ExceptionObject as Exception ?? new Exception("Unknown exception");
        LogFatal("Unhandled exception", ex);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

    private static void LogFatal(string message, Exception ex)
    {
        var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MarkOneNoteDown");
        Directory.CreateDirectory(logDir);
        var logPath = Path.Combine(logDir, "startup.log");
        var details = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}{Environment.NewLine}{ex}{Environment.NewLine}";
        File.AppendAllText(logPath, details);
        MessageBoxW(IntPtr.Zero, $"{message}\n\n{ex.Message}\n\nLog: {logPath}", "MarkOneNoteDown", 0x00000010);
    }
}


