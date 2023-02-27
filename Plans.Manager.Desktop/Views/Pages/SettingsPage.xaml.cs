using Wpf.Ui.Appearance;

namespace Plans.Manager.Desktop.Views.Pages;

public partial class SettingsPage
{
    public SettingsPage()
    {
        InitializeComponent();
        Accent.ApplySystemAccent();
        Theme.GetAppTheme();

        AboutAppTextBlock.Text = $"Plans Manager Desktop {GetAssemblyVersion()}";
    }

    private static string GetAssemblyVersion()
    {
        return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? string.Empty;
    }
}
