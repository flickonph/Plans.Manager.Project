using System.Windows;
using Plans.Manager.BLL.Readers;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace Plans.Manager.Desktop.Views.Pages;

public partial class GroupsPage
{
    public GroupsPage()
    {
        InitializeComponent();
        Accent.ApplySystemAccent();
        Theme.GetAppTheme();
        InitGroups();
    }

    private void InitGroups()
    {
        var allGroups = ConfigReader.GetGroups();
        foreach (var group in allGroups)
        {
            var groupCard = new Card
            {
                Content = group.Name,
                Footer = group.AllData(),
                FontSize = 12,
            };
            GroupsListBox.Items.Add(groupCard);
        }
    }

    private void AddGroupButton_OnClick(object sender, RoutedEventArgs e)
    {
        // TODO: Append method
    }

    private void RemoveGroupButton_OnClick(object sender, RoutedEventArgs e)
    {
        // TODO: Append method
    }
    
    private void ReloadGroups()
    {
        GroupsListBox.Items.Clear();
        InitGroups();
    }
}
