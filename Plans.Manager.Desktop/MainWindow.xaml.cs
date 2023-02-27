using System.Windows;
using System.Windows.Media;
using System.Reflection;
using Plans.Manager.Desktop.Views.Pages;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace Plans.Manager.Desktop;

public partial class MainWindow
{
    public MainWindow()
    {
        DataContext = this;

        Watcher.Watch(this);

        InitializeComponent();

        Loaded += (_, _) => RootNavigation.Navigate(typeof(MainPage));
    }
    
    private static string GetAssemblyVersion()
    {
        return Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? string.Empty;
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
