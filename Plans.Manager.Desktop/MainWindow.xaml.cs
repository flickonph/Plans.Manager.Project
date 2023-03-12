using System.Configuration;
using System.Windows;
using System.Windows.Media;
using Plans.Manager.Desktop.Views.Pages;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;
using Wpf.Ui.Controls.Window;

namespace Plans.Manager.Desktop;

public partial class MainWindow
{
    public MainWindow()
    {
        DataContext = this;

        Watcher.Watch(this);

        InitializeComponent();
        Loaded += (_, _) =>
        {
            Watcher.Watch(
                this,
                WindowBackdropType.Mica,
                true
            );
            RootNavigation.Navigate(typeof(MainPage));
            SetDefaultTheme_OnLoad();
        };
    }

    private static void SetDefaultTheme_OnLoad()
    {
        string? themeValue = ConfigurationManager.AppSettings.Get("Theme");
        switch (themeValue)
        {
            case "Dark":
                Theme.Apply(ThemeType.Dark);
                break;
            case "Light":
                Theme.Apply(ThemeType.Light);
                break;
            default:
                if (Theme.GetSystemTheme() == SystemThemeType.Dark)
                {
                    Theme.Apply(ThemeType.Dark);
                }
                else if (Theme.GetSystemTheme() == SystemThemeType.Light)
                {
                    Theme.Apply(ThemeType.Light);
                }
                break;
        }
    }

    private void ThemeButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (Theme.GetAppTheme() == ThemeType.Dark)
        {
            Theme.Apply(ThemeType.Light);
            ThemeSwitchButton.BorderBrush = new SolidColorBrush(Colors.Black);
        }
        else
        {
            Theme.Apply(ThemeType.Dark);
            ThemeSwitchButton.BorderBrush = new SolidColorBrush(Colors.White);
        }
    }

    private void NotifyIcon_OnLeftClick(NotifyIcon sender, RoutedEventArgs e)
    {
        NotifyIcon.Dispose();
    }
}
