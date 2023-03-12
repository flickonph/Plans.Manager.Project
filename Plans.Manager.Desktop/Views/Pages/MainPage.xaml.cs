using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Plans.Manager.BLL.Services;
using Plans.Manager.Desktop.Models;
using Plans.Manager.Desktop.Views.Windows;
using Wpf.Ui.Appearance;
using Wpf.Ui.Common;
using Wpf.Ui.Controls;
using MessageBox = Wpf.Ui.Controls.MessageBox;

namespace Plans.Manager.Desktop.Views.Pages;

public partial class MainPage
{
    private readonly int _studyYear;
    
    public MainPage()
    {
        InitializeComponent();
        Accent.ApplySystemAccent();
        Theme.GetAppTheme();
        _studyYear = StudyYear.CurrentStudyYear();
    }

    private async void CreateStudyPlans(object sender, RoutedEventArgs e)
    {
        CreatePlansButton.IsEnabled = false;
        
        string dirToSave = WindowsApi.SelectFolder("Выберите папку для сохранения рабочих планов");
        if (string.IsNullOrEmpty(dirToSave))
        {
            MessageBox messageBox = MessageBox(string.Empty, "Папка не указана", string.Empty);
            messageBox.ButtonLeftClick += (_, _) => messageBox.Close();
            messageBox.ShowDialog();
            CreatePlansButton.IsEnabled = true;
            return;
        }

        PlansOptionsWindow plansOptionsWindow = new PlansOptionsWindow();
        bool? dialog = plansOptionsWindow.ShowDialog();
        if (dialog is null or false)
        {
            MessageBox messageBox = MessageBox(string.Empty, "Опции не выбраны", string.Empty);
            messageBox.ButtonLeftClick += (_, _) => messageBox.Close();
            messageBox.ShowDialog();
            CreatePlansButton.IsEnabled = true;
            return;
        }

        bool result = false;
        string parentDir = WindowsApi.SelectFolder("Выберите корневую папку с исходными учебными планами");
        List<string> plxFilesPaths = new List<string>();
        
        CreatePlansStatus.Visibility = Visibility.Visible;
        
        switch (plansOptionsWindow.Option)
        {
            case PlansOptionsWindow.PlansType.Default:
                plxFilesPaths = DefaultOptionPlxFilesPaths(parentDir);
                break;
            case PlansOptionsWindow.PlansType.Optional:
                plxFilesPaths = AdditionalOptionPlxFilesPaths(parentDir);
                break;
            case PlansOptionsWindow.PlansType.All:
                // TODO: Add option for all types
                break;
            case PlansOptionsWindow.PlansType.Undefined:
                // TODO: Add notification for undefined option
                return;
            default:
                return;
        }
        
        for (int course = 1; course <= 6; course++)
        {
            result = await Task.Factory.StartNew(()
                => DocumentService.CreateAndSavePlan(
                    plxFilesPaths,
                    course,
                    (int)plansOptionsWindow.Semester,
                    plansOptionsWindow.Weeks,
                    plansOptionsWindow.Year,
                    dirToSave)).ConfigureAwait(true);
        }
        
        CreatePlansStatus.Visibility = Visibility.Collapsed;

        switch (result)
        {
            case true:
                Notification("Успешное сохранение", $"Рабочие планы сохранены в {dirToSave}", SymbolRegular.Checkmark24, ControlAppearance.Success);
                break;
            default:
                Notification("Что-то пошло не так", "Рабочие планы не сохранены", SymbolRegular.ErrorCircle24, ControlAppearance.Caution);
                break;
        }

        CreatePlansButton.IsEnabled = true;
    }

    private async void CreateStudyLoad(object sender, RoutedEventArgs e)
    {
        CreateLoadButton.IsEnabled = false;
        
        string dirToSave = WindowsApi.SelectFolder("Выберите папку для сохранения кафедральной нагрузки");
        if (string.IsNullOrEmpty(dirToSave))
        {
            MessageBox messageBox = MessageBox(string.Empty, "Папка не указана", string.Empty);
            messageBox.ButtonLeftClick += (_, _) => messageBox.Close();
            messageBox.ShowDialog();
            CreateLoadButton.IsEnabled = true;
            return;
        }

        PlansOptionsWindow loadOptionsWindow = new PlansOptionsWindow();
        loadOptionsWindow.WeeksNumberBox.IsEnabled = false;
        loadOptionsWindow.WeeksTextBlock.IsEnabled = false;
        loadOptionsWindow.TitleBar.Title = "Опции для составления кафедральной нагрузки";
        loadOptionsWindow.DefaultCardActionMainTextBlock.Text = "Составить кафедральную нагрузку";
        loadOptionsWindow.DefaultCardActionAdditionalTextBlock.Text = "Нагрузка составляется с учётом заданных опций";
        loadOptionsWindow.TypeTextBlock.Text = "Тип кафедральной нагрузки";

        bool? dialog = loadOptionsWindow.ShowDialog();
        if (dialog is null or false)
        {
            MessageBox messageBox = MessageBox(string.Empty, "Опции не выбраны", string.Empty);
            messageBox.ButtonLeftClick += (_, _) => messageBox.Close();
            messageBox.ShowDialog();
            CreateLoadButton.IsEnabled = true;
            return;
        }

        string parentDir = WindowsApi.SelectFolder("Выберите корневую папку с исходными учебными планами");
        List<string> plxFilesPaths = new List<string>();
        
        CreateLoadStatus.Visibility = Visibility.Visible;
        
        switch (loadOptionsWindow.Option)
        {
            case PlansOptionsWindow.PlansType.Default:
                plxFilesPaths = DefaultOptionPlxFilesPaths(parentDir);
                break;
            case PlansOptionsWindow.PlansType.Optional:
                plxFilesPaths = AdditionalOptionPlxFilesPaths(parentDir);
                break;
            case PlansOptionsWindow.PlansType.All:
                // TODO: Add option for all types
                break;
            case PlansOptionsWindow.PlansType.Undefined:
                // TODO: Add notification for undefined option
                return;
            default:
                return;
        }
        
        bool result = await Task.Factory.StartNew(()
            => DocumentService.CreateAndSaveLoad(
                plxFilesPaths,
                (int)loadOptionsWindow.Semester,
                loadOptionsWindow.Year,
                dirToSave)).ConfigureAwait(true);
        
        CreateLoadStatus.Visibility = Visibility.Collapsed;

        switch (result)
        {
            case true:
                Notification("Успешное сохранение", $"Кафедральная нагрузка сохранена в {dirToSave}", SymbolRegular.Checkmark24, ControlAppearance.Success);
                break;
            default:
                Notification("Что-то пошло не так", "Кафедральная нагрузка не сохранена", SymbolRegular.ErrorCircle24, ControlAppearance.Caution);
                break;
        }

        CreateLoadButton.IsEnabled = true;
    }

    private async void CreateDepartments(object sender, RoutedEventArgs e)
    {
        const string dirToSave = @"Config";
        string pathToFile = WindowsApi.SelectFile("dat");
        if (pathToFile == string.Empty)
        {
            MessageBox messageBox = MessageBox(string.Empty, "Папка не указана.", string.Empty);
            messageBox.ButtonLeftClick += (_, _) => messageBox.Close();
            messageBox.ShowDialog();
            return;
        }

        bool result = await Task.Factory.StartNew(() => DocumentService.CreateAndSaveDepartments(pathToFile))
            .ConfigureAwait(true);
        if(result)
            Notification("Успешное сохранение",
                $"Справочник кафедр сохранён в {dirToSave}",
                SymbolRegular.Checkmark24,
                ControlAppearance.Success);
    }
}

public partial class MainPage
{
    private List<string> DefaultOptionPlxFilesPaths(string parentDir)
    {
        List<string> result = new List<string>();
        DirectoryInfo parentDirInfo = new DirectoryInfo(parentDir);
        if (!parentDirInfo.Exists) return result;
        if (parentDirInfo.GetDirectories().Length == 0) return result;
        foreach (DirectoryInfo yearDirInfo in parentDirInfo.GetDirectories())
        {
            if (yearDirInfo.Name is ".git" or ".idea") continue; // TODO: Refactor this method
            
            int yearDirName;
            try
            {
                yearDirName = Convert.ToInt32(yearDirInfo.Name);
            }
            catch
            {
                throw new InvalidCastException("invalid folder name");
            }
            
            if (yearDirName <= _studyYear && yearDirName >= (_studyYear - 3))
            {
                foreach (DirectoryInfo typeDirInfo in yearDirInfo.GetDirectories())
                {
                    switch (typeDirInfo.Name)
                    {
                        case "Бакалавриат":
                            result.AddRange(typeDirInfo.GetFiles().Select(file => file.FullName));
                            break;
                    }
                }
            }
            
            if (yearDirName <= _studyYear && yearDirName >= (_studyYear - 5))
            {
                foreach (DirectoryInfo typeDirInfo in yearDirInfo.GetDirectories())
                {
                    switch (typeDirInfo.Name)
                    {
                        case "Специалитет":
                            result.AddRange(typeDirInfo.GetFiles().Select(file => file.FullName));
                            break;
                    }
                }
            }
            
            
        }

        return result;
    }
    
    private List<string> AdditionalOptionPlxFilesPaths(string parentDir)
    {
        List<string> result = new List<string>();
        DirectoryInfo parentDirInfo = new DirectoryInfo(parentDir);
        if (!parentDirInfo.Exists) return result;
        if (parentDirInfo.GetDirectories().Length == 0) return result;
        foreach (DirectoryInfo yearDirInfo in parentDirInfo.GetDirectories())
        {
            int yearDirName;
            try
            {
                yearDirName = Convert.ToInt32(yearDirInfo.Name);
            }
            catch
            {
                throw new InvalidCastException("invalid folder name");
            }
            
            if (yearDirName <= (_studyYear) && yearDirName >= (_studyYear - 1))
            {
                foreach (DirectoryInfo typeDirInfo in yearDirInfo.GetDirectories())
                {
                    switch (typeDirInfo.Name)
                    {
                        case "Магистратура":
                            result.AddRange(typeDirInfo.GetFiles().Select(file => file.FullName));
                            break;
                    }
                }
            }
        }

        return result;
    }

    private void Notification(string title, string message, SymbolRegular symbol, ControlAppearance appearance)
    {
        NotificationSnackbar.Timeout = 5000;
        NotificationSnackbar.Show(title, message, symbol, appearance);
    }

    private static Wpf.Ui.Controls.MessageBox MessageBox(string title, string content, string actionButtonContent)
    {
        if (string.IsNullOrEmpty(title))
        {
            title = "Plans Manager Desktop";
        }

        if (string.IsNullOrEmpty(actionButtonContent))
        {
            actionButtonContent = "Ок";
        }
        Wpf.Ui.Controls.MessageBox messageBox = new()
        {
            Title = title,
            ButtonLeftName = actionButtonContent,
            ButtonRightName = "Отмена",
            Content = content
        };
        messageBox.ButtonRightClick += (_, _) => messageBox.Close();
        return messageBox;
    }
}
