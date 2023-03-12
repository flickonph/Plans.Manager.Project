using Wpf.Ui.Appearance;
using System.Configuration;
using System.Windows;

namespace Plans.Manager.Desktop.Views.Pages;

public partial class SettingsPage
{
    public SettingsPage()
    {
        InitializeComponent();
        Accent.ApplySystemAccent();
        Theme.GetAppTheme();
        DefaultThemeRadioButtons_Load();

        AppInfoTextBlock.Text = $"Plans Manager Desktop {GetAssemblyVersion()}";
    }

    private static string GetAssemblyVersion()
    {
        return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? string.Empty;
    }

    private void DefaultThemeRadioButtons_Load()
    {
        string? themeValue = ConfigurationManager.AppSettings.Get("Theme");
        switch (themeValue)
        {
            case "Dark":
                DarkThemeRadioButton.IsChecked = true;
                break;
            case "Light":
                LightThemeRadioButton.IsChecked = true;
                break;
            case "System":
                SystemThemeRadioButton.IsChecked = true;
                break;
            default:
                return;
        }
    }

    private void ChangeDefaultTheme(object sender, RoutedEventArgs e)
    {
        if (Equals(sender, DarkThemeRadioButton))
        {
            
        }
        else if (Equals(sender, LightThemeRadioButton))
        {
            
        }
    }
}
