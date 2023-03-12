using System;
using System.Windows;
using Plans.Manager.Desktop.Models;

namespace Plans.Manager.Desktop.Views.Windows;

public partial class PlansOptionsWindow
{
    public PlansOptionsWindow()
    {
        InitializeComponent();
        YearNumberBox.Value = StudyYear.CurrentStudyYear();
        YearSecondaryTextBlock.Text = $"{(int)YearNumberBox.Value}-{(int)YearNumberBox.Value + 1} гг.";
    }

    public PlansType Option { get; private set; }
    public PlansSemester Semester { get; private set; }
    public int Weeks { get; private set; }
    public int Year { get; private set; }

    private void DefaultCardAction_OnClick(object sender, RoutedEventArgs e)
    {
        Weeks = Convert.ToInt32(WeeksNumberBox.Value);
        Year = Convert.ToInt32(YearNumberBox.Value);
        
        if (DefaultTypeCheckBox.IsChecked == true & OptionalTypeCheckBox.IsChecked == false)
        {
            Option = PlansType.Default;
        }
        else if (OptionalTypeCheckBox.IsChecked == true & DefaultTypeCheckBox.IsChecked == false)
        {
            Option = PlansType.Optional;
        }
        else if (DefaultTypeCheckBox.IsChecked == true & OptionalTypeCheckBox.IsChecked == false)
        {
            Option = PlansType.All;
        }
        else
        {
            Option = PlansType.Undefined;
            DialogResult = false;
            Close();
            return;
        }

        if (SpringSemesterRadioButton.IsChecked == true)
        {
            Semester = PlansSemester.Spring;
        }
        else if (FallSemesterRadioButton.IsChecked == true)
        {
            Semester = PlansSemester.Fall;
        }
        else
        {
            Semester = PlansSemester.Undefined;
            DialogResult = false;
            Close();
            return;
        }

        DialogResult = true;
        Close();
    }
    
    public enum PlansType
    {
        Default = 0,
        Optional = 1,
        All = 2,
        Undefined = -1
    }
    
    public enum PlansSemester
    {
        Fall = 1,
        Spring = 2,
        Undefined = -1
    }

    private void YearNumberBox_OnValueChanged(object sender, RoutedEventArgs e)
    {
        int year = Convert.ToInt32(YearNumberBox.Value);
        YearSecondaryTextBlock.Text = $"{year}-{year + 1} гг.";
    }
}