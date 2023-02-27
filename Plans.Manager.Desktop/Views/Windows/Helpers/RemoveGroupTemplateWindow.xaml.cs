using System.Windows;

namespace Plans_Manager_WPF.Windows.Helpers;

public partial class RemoveGroupTemplateWindow
{
    public RemoveGroupTemplateWindow()
    {
        InitializeComponent();
    }

    public string? GroupName { get; private set; }

    private void Ok_Button(object sender, RoutedEventArgs e)
    {
        GroupName = GroupNameTextBox.Text;
        Close();
    }

    private void Cancel_Button(object sender, RoutedEventArgs e)
    {
        Close();
    }
}