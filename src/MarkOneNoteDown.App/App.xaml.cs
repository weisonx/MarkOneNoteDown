using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Navigation;
using MarkOneNoteDown.App.Views;

namespace MarkOneNoteDown.App;

public partial class App : Application
{
    private Window? window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs e)
    {
        window ??= new Window();
        window.Title = "MarkOneNoteDown";

        if (window.Content is not Frame rootFrame)
        {
            rootFrame = new Frame();
            rootFrame.NavigationFailed += OnNavigationFailed;
            window.Content = rootFrame;
        }

        _ = rootFrame.Navigate(typeof(MainPage), e.Arguments);
        window.Activate();
    }

    private static void OnNavigationFailed(object sender, NavigationFailedEventArgs e)
        => throw new InvalidOperationException($"Failed to load Page {e.SourcePageType.FullName}");
}
