using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Plans.Manager.BLL.Objects;
using Plans.Manager.BLL.Providers;
using Plans.Manager.BLL.Readers;

namespace Plans.Manager.BLL.Builders;

public class PlanBuilder : ABuilder
{
    private readonly List<PlanGroupPair> _allPlanGroupPairs;
    private readonly int _selectedYear;
    private static readonly List<string> AutoComplete = XmlReader.GetAutoCompleteDisciplines();

    public PlanBuilder(List<string> plxFilesPaths, int selectedYear) : base(plxFilesPaths, selectedYear)
    {
        _selectedYear = selectedYear;
        _allPlanGroupPairs = GetAllPlanGroupPairs();
    }

    public ExcelPackage PlansExcelPackage(int selectedCourse, int selectedSemester, int selectedWeeks)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelPackage plansPackage = new ExcelPackage();
        plansPackage.Workbook.Properties.Author = MetadataProvider.Author;
        plansPackage.Workbook.Properties.Title = MetadataProvider.PlanTitle;
        plansPackage.Workbook.Properties.Subject = MetadataProvider.Subject;
        plansPackage.Workbook.Properties.Company = MetadataProvider.Company;
        plansPackage.Workbook.Properties.Created = MetadataProvider.Created;

        string semesterDescription = selectedSemester switch
        {
            1 => $"на осенний семестр {_selectedYear}/{_selectedYear + 1} уч. год ({selectedWeeks} недель)",
            2 => $"на весенний семестр {_selectedYear}/{_selectedYear + 1} уч. год ({selectedWeeks} недель)",
            _ => "ошибка"
        };
        
        List<PlanGroupPair> planGroupPairs = _allPlanGroupPairs.Where(pair => pair.Groups.First().Years == selectedCourse).ToList(); // TODO: Fix for selecting proper groups

        string localGroupName, localNumberOfStudents;
        foreach (PlanGroupPair pair in planGroupPairs)
        {
            pair.Plan.Data.SortByCourseAndSemester(
                (_selectedYear - pair.Plan.Data.YearOfAdmission + 1), selectedSemester, true);
            if (pair.Plan.Data.Disciplines.Count == 0) continue;

            localGroupName = string.Empty;
            localNumberOfStudents = string.Empty;
            foreach (Group educationalGroup in pair.Groups)
            {
                localGroupName += educationalGroup.Name + ", ";
                localNumberOfStudents += educationalGroup.NumberOfStudents + ", ";
            }

            localGroupName = localGroupName[..^2];
            localNumberOfStudents = localNumberOfStudents[..^2];
            
            ExcelWorksheet groupWorksheet = plansPackage.Workbook.Worksheets.Add($"{localGroupName}");
            groupWorksheet.View.ZoomScale = 70;

            FillGroupWorksheet(
                groupWorksheet,
                pair.Plan,
                selectedWeeks + 1, localGroupName, localNumberOfStudents, semesterDescription); // TODO: Remove + 1
        }

        return plansPackage;
    }
    
    private static void FillGroupWorksheet(ExcelWorksheet groupWorksheet, Plan plan, int selectedWeeks, string localGroupName, string localNumberOfStudents, string semesterDescription)
    {
        List<Discipline> disciplines = plan.Data.Disciplines;
        List<Discipline> skippedDisciplines = new List<Discipline>();
        
        int worksheetIndex, worksheetLine = 0;
        int disciplineIndex = 0;
        for (worksheetIndex = 0; worksheetIndex < disciplines.Count; worksheetIndex++)
        {
            // Skipping
            if (disciplines[worksheetIndex].Name is "Элективные дисциплины по физической культуре и спорту" or "Физическая культура" or "Физическая культура и спорт")
            {
                skippedDisciplines.Add(disciplines[worksheetIndex]);
                continue;
            }
            
            worksheetLine = 9 + disciplineIndex * 3;

            FillPrimaryData(groupWorksheet, disciplines[worksheetIndex], worksheetLine, disciplineIndex);
            
            FillControlType(groupWorksheet, disciplines[worksheetIndex], selectedWeeks, worksheetLine);
            
            FillHours(groupWorksheet, disciplines[worksheetIndex], selectedWeeks, worksheetLine);

            disciplineIndex++;
        }
        PlanDecorator.DecorateHeader(groupWorksheet, plan, semesterDescription, localGroupName, selectedWeeks, disciplineIndex, localNumberOfStudents);

        worksheetLine += 5;
        worksheetIndex += 5;

        foreach (Discipline discipline in skippedDisciplines)
        {
            FillPrimaryData(groupWorksheet, discipline, worksheetLine, disciplineIndex);
            
            FillControlType(groupWorksheet, discipline, selectedWeeks, worksheetLine);
            
            FillHours(groupWorksheet, discipline, selectedWeeks, worksheetLine);
            
            worksheetLine += 3;
            worksheetIndex += 3;
            disciplineIndex++;
        }
        PlanDecorator.DecorateFooter(groupWorksheet, worksheetLine - 2); // TODO: Remove -2
    }

    private static void FillPrimaryData(ExcelWorksheet groupWorksheet, Discipline discipline, int line, int index)
    {
        // Name of discipline & it's number
        groupWorksheet.Cells[line, 2].Value = discipline.Name;
        groupWorksheet.Cells[line, 1].Value = index + 1;
        groupWorksheet.Cells[line, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        groupWorksheet.Cells[line, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        groupWorksheet.Cells[line, 2].Style.Font.Bold = false;
        groupWorksheet.Cells[line, 2].Style.WrapText = true;

        // Hours of discipline's types of study
        groupWorksheet.Cells[line, 4].Value =
            discipline.Data.Find(disciplineData => 
                disciplineData.Type == "Лекционные занятия")?.Hours ?? 0;
        groupWorksheet.Cells[line, 5].Value =
            discipline.Data.Find(disciplineData => 
                disciplineData.Type == "Практические занятия")?.Hours ?? 0;
        groupWorksheet.Cells[line, 6].Value =
            discipline.Data.Find(disciplineData => 
                disciplineData.Type == "Лабораторные занятия")?.Hours ?? 0;
        
        groupWorksheet.Cells[line, 3].Formula = "=SUM(" + groupWorksheet.Cells[line, 4].Address + ":" +
                                           groupWorksheet.Cells[line, 6].Address + ")";

        // Styling of primary discipline's data
        groupWorksheet.Cells[line, 1, line, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        groupWorksheet.Cells[line, 1, line, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        for (int j = 1; j < 7; ++j)
        {
            groupWorksheet.Cells[line, j, line + 2, j].Merge = true;
            groupWorksheet.Cells[line, j, line + 2, j].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }
    }

    private static void FillControlType(ExcelWorksheet groupWorksheet, Discipline discipline, int weeks, int line)
    {
        for (int index = 0; index < 4; index++)
        {
            groupWorksheet.Cells[line, 7 + weeks * 3 + index, line + 2, 7 + weeks * 3 + index].Merge = true;
            groupWorksheet.Cells[line, 7 + weeks * 3 + index, line + 2, 7 + weeks * 3 + index].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
            groupWorksheet.Cells[line, 7 + weeks * 3 + index, line + 2, 7 + weeks * 3 + index].Style.HorizontalAlignment =
                ExcelHorizontalAlignment.Center;
            groupWorksheet.Cells[line, 7 + weeks * 3 + index, line + 2, 7 + weeks * 3 + index].Style.VerticalAlignment =
                ExcelVerticalAlignment.Center;
        }

        List<DisciplineData> controlTypeMatches;
        try
        {
            controlTypeMatches = discipline.Data.Where(disciplineData => ControlType.Contains(disciplineData.Type)).ToList();
        }
        catch
        {
            controlTypeMatches = new List<DisciplineData> { new("Ошибка", 0, 0, 0) };
        }

        foreach (var controlTypeMatch in controlTypeMatches)
        {
            switch (controlTypeMatch.Type)
            {
                case "Зачет":
                    groupWorksheet.Cells[line, 7 + weeks * 3].Value = controlTypeMatch.Hours;
                    break;
                case "Зачет с оценкой":
                    groupWorksheet.Cells[line, 7 + weeks * 3].Value = controlTypeMatch.Hours;
                    groupWorksheet.Cells[line, 7 + weeks * 3].Style.Fill.SetBackground(Color.Gold);
                    break;
                case "Экзамен":
                    groupWorksheet.Cells[line, 7 + weeks * 3 + 1].Value = controlTypeMatch.Hours;
                    break;
                case "Курсовая работа":
                    groupWorksheet.Cells[line, 7 + weeks * 3 + 2].Value = controlTypeMatch.Hours;
                    break;
                case "Курсовой проект":
                    groupWorksheet.Cells[line, 7 + weeks * 3 + 2].Value = controlTypeMatch.Hours;
                    break;
                case "Ошибка":
                    groupWorksheet.Cells[line, 7 + weeks * 3 + 3].Value = "Ошибка";
                    groupWorksheet.Cells[line, 7 + weeks * 3 + 3].Style.Fill.SetBackground(Color.IndianRed);
                    break;
            }
        }
    }

    private static void FillHours(ExcelWorksheet groupWorksheet, Discipline discipline, int weeks, int line)
    {
        for (int j = 7; j <= weeks * 3 + 6; j += 3)
            groupWorksheet.Cells[line, j, line + 2, j + 2].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        
        if (!AutoComplete.Contains(discipline.Name)) return;
        
        double totalHours = discipline.Data.Find(disciplineData => disciplineData.Type == "Лекционные занятия")?.Hours ?? 0;
        double weekHours;
        if (totalHours != 0)
        {
            weekHours = totalHours / (weeks - 1);
            for (int j = 7; j <= (weeks - 1) * 3 + 6; j += 3)
            {
                groupWorksheet.Cells[line, j].Value = Math.Floor(weekHours);
            }
        }
                
        totalHours = discipline.Data.Find(disciplineData => disciplineData.Type == "Практические занятия")?.Hours ?? 0;
        if (totalHours != 0)
        {
            weekHours = totalHours / (weeks - 1);
            for (int j = 7; j <= (weeks - 1) * 3 + 6; j += 3)
            {
                groupWorksheet.Cells[line + 1, j + 1].Value = Math.Floor(weekHours);
            }
        }
                
        totalHours = discipline.Data.Find(disciplineData => disciplineData.Type == "Лабораторные занятия")?.Hours ?? 0;
        if (totalHours != 0)
        {
            weekHours = totalHours / (weeks - 1);
            for (int j = 7; j <= (weeks - 1) * 3 + 6; j += 3)
            {
                groupWorksheet.Cells[line + 2, j + 2].Value = Math.Floor(weekHours);
            }
        }
    }
}

file static class PlanDecorator
{
    internal static void DecorateHeader(ExcelWorksheet planWorksheet, Plan plan, string semesterDescription, string groupName, int selectedWeeks, int disciplinesCount, string numberOfStudents)
    {
        int yearOfAdmission = plan.Data.YearOfAdmission;
        string edDirection = plan.Data.Code;


        // Font formatting
        planWorksheet.Cells[1, 1, 100, 100].Style.Font.Bold = true;
        planWorksheet.Cells[1, 1, 100, 100].Style.Font.Name = "Times New Roman";
        planWorksheet.Cells[1, 1, 100, 100].Style.Font.Size = 12;

        // Rows formatting
        planWorksheet.Rows[1].Height = 6.00;
        planWorksheet.Rows[2].Height = 24.00;
        planWorksheet.Rows[3].Height = 6.00;
        planWorksheet.Rows[4].Height = 14.40;
        planWorksheet.Rows[5].Height = 19.20;
        planWorksheet.Rows[6].Height = 24.00;
        planWorksheet.Rows[7].Height = 57.60;
        for (int i = 8; i <= 60; i++) planWorksheet.Rows[i].Height = 15.00;

        // Columns formatting
        planWorksheet.Columns[1].Width = 4.78;
        planWorksheet.Columns[2].Width = 43.67;
        planWorksheet.Columns[3].Width = 5.89;
        planWorksheet.Columns[4].Width = 5.89;
        planWorksheet.Columns[5].Width = 5.89;
        planWorksheet.Columns[6].Width = 5.89;
        int col;
        for (col = 7; col <= selectedWeeks * 3 + 6; col++) planWorksheet.Columns[col].Width = 3.00;

        for (int i = 1; i <= 3; i++)
        {
            planWorksheet.Columns[col].Width = 4.67;
            col++;
        }

        planWorksheet.Columns[col].Width = 23.00;

        // Filling of the worksheet header
        planWorksheet.Cells[2, 2].Value = $"Год приёма: {yearOfAdmission}";
        planWorksheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        planWorksheet.Cells[2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

        planWorksheet.Cells[2, 3].Value = "Рабочий учебный план";
        planWorksheet.Cells[2, 3].Style.Font.Size = 22;
        planWorksheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[2, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        planWorksheet.Cells[2, 3, 2, 43].Merge = true;

        planWorksheet.Cells[2, 48].Value = $"Направление: {edDirection}";
        planWorksheet.Cells[2, 48].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        planWorksheet.Cells[2, 48].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        planWorksheet.Cells[2, 48, 2, 61].Merge = true;

        planWorksheet.Cells[4, 3].Value = semesterDescription;
        planWorksheet.Cells[4, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[4, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        planWorksheet.Cells[4, 3, 4, 43].Merge = true;

        planWorksheet.Cells[5, 2].Value = $"Студентов: {numberOfStudents}";
        planWorksheet.Cells[5, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        planWorksheet.Cells[5, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        planWorksheet.Cells[5, 2, 5, 8].Merge = true;

        planWorksheet.Cells[5, 48].Value = $"Группы: {groupName}";
        planWorksheet.Cells[5, 48].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        planWorksheet.Cells[5, 48].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        planWorksheet.Cells[5, 48, 5, 61].Merge = true;

        // Filling of the table header
        planWorksheet.Cells[6, 1].Value = "№ п/п";
        planWorksheet.Cells[6, 1].Style.WrapText = true;
        planWorksheet.Cells[6, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, 2].Value = "Дисциплина";
        planWorksheet.Cells[6, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, 3].Value = "Ауд.";
        planWorksheet.Cells[6, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, 4].Value = "Лек.";
        planWorksheet.Cells[6, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, 5].Value = "Сем.";
        planWorksheet.Cells[6, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, 6].Value = "Лаб.";
        planWorksheet.Cells[6, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        for (int i = 1; i < 7; ++i)
        {
            planWorksheet.Cells[6, i, 7, i].Merge = true;
            planWorksheet.Cells[6, i, 7, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            planWorksheet.Cells[8, i].Value = i;
            planWorksheet.Cells[8, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            planWorksheet.Cells[8, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[8, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        planWorksheet.Cells[6, 7].Value = "Распределение по неделям семестра (лекц./сем./лаб.)";
        planWorksheet.Cells[6, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        planWorksheet.Cells[6, 7, 6, selectedWeeks * 3 + 6].Merge = true;
        planWorksheet.Cells[6, 7, 6, selectedWeeks * 3 + 6].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);

        for (int i = 0; i < selectedWeeks; ++i)
        {
            col = 7 + i * 3;
            planWorksheet.Cells[7, col].Value = i + 1;
            planWorksheet.Cells[7, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[7, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            planWorksheet.Cells[7, col, 7, col + 2].Merge = true;
            planWorksheet.Cells[7, col, 7, col + 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            planWorksheet.Cells[8, col, 8, col + 2].Merge = true;
            planWorksheet.Cells[8, col].Value = i + 7;
            planWorksheet.Cells[8, col, 8, col + 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            planWorksheet.Cells[8, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[8, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        int nextColumn = 7 + selectedWeeks * 3;

        planWorksheet.Cells[6, nextColumn].Value = "Контроль";
        planWorksheet.Cells[6, nextColumn].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        planWorksheet.Cells[6, nextColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, nextColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[7, nextColumn].Value = "Зачет";
        planWorksheet.Cells[7, nextColumn].Style.TextRotation = 90;
        planWorksheet.Cells[7, nextColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[7, nextColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[7, nextColumn + 1].Value = "Экзамен";
        planWorksheet.Cells[7, nextColumn + 1].Style.TextRotation = 90;
        planWorksheet.Cells[7, nextColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[7, nextColumn + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[7, nextColumn + 2].Value = "К/р (пр-т)";
        planWorksheet.Cells[7, nextColumn + 2].Style.TextRotation = 90;
        planWorksheet.Cells[7, nextColumn + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[7, nextColumn + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, nextColumn + 3].Value = "Примечание";
        planWorksheet.Cells[6, nextColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[6, nextColumn + 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[6, nextColumn, 6, nextColumn + 2].Merge = true;
        planWorksheet.Cells[6, nextColumn, 6, nextColumn + 2].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        planWorksheet.Cells[6, nextColumn + 3, 7, nextColumn + 3].Merge = true;
        planWorksheet.Cells[6, nextColumn + 3, 7, nextColumn + 3].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        for (int i = nextColumn; i <= nextColumn + 3; i++)
        {
            planWorksheet.Cells[8, i].Value = selectedWeeks + 7 + i - nextColumn;
            planWorksheet.Cells[8, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            planWorksheet.Cells[8, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[8, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        // Filling of the table bottom part
        int line = 9 + disciplinesCount * 3;
        planWorksheet.Rows[line].Height = 27.60;
        planWorksheet.Rows[line + 1].Height = 27.60;
        planWorksheet.Cells[line, 2].Value = "Всего часов в неделю:";
        planWorksheet.Cells[line, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        planWorksheet.Cells[line, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        planWorksheet.Cells[line, 3].Formula = "=SUM(" + planWorksheet.Cells[9, 3].Address + ":" +
                                                  planWorksheet.Cells[line - 1, 3].Address + ")";
        planWorksheet.Cells[line, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[line, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        planWorksheet.Cells[line + 1, 2].Value = "В том числе лекций:";
        planWorksheet.Cells[line + 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        planWorksheet.Cells[line + 1, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        planWorksheet.Cells[line + 1, 3].Formula = "=SUM(" + planWorksheet.Cells[9, 4].Address + ":" +
                                                      planWorksheet.Cells[line - 1, 4].Address + ")";
        planWorksheet.Cells[line + 1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        planWorksheet.Cells[line + 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        for (int i = 1; i <= 6; i++)
        {
            planWorksheet.Cells[line, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            planWorksheet.Cells[line + 1, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }

        int fromCol;
        for (fromCol = 7; fromCol <= selectedWeeks * 3 + 6; fromCol += 3)
        {
            planWorksheet.Cells[line, fromCol].Formula = "=SUM(" + planWorksheet.Cells[9, fromCol].Address + ":" +
                                                            planWorksheet.Cells[line - 1, fromCol + 2].Address + ")";
            planWorksheet.Cells[line, fromCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[line, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            planWorksheet.Cells[line, fromCol, line, fromCol + 2].Merge = true;
            planWorksheet.Cells[line, fromCol, line, fromCol + 2].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);

            planWorksheet.Cells[line + 1, fromCol].Formula = "=SUM(" + planWorksheet.Cells[9, fromCol].Address +
                                                                ":" + planWorksheet.Cells[line - 1, fromCol]
                                                                    .Address + ")";
            planWorksheet.Cells[line + 1, fromCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[line + 1, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            planWorksheet.Cells[line + 1, fromCol, line + 1, fromCol + 2].Merge = true;
            planWorksheet.Cells[line + 1, fromCol, line + 1, fromCol + 2].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }

        for (int i = 0; i < 3; i++)
        {
            planWorksheet.Cells[line, fromCol + i].Formula = "=SUM(" +
                                                                planWorksheet.Cells[9, fromCol + i].Address + ":" +
                                                                planWorksheet.Cells[line - 1, fromCol + i].Address +
                                                                ")";
            planWorksheet.Cells[line, fromCol + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            planWorksheet.Cells[line, fromCol + i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            planWorksheet.Cells[line, fromCol + i].Merge = true;
            planWorksheet.Cells[line, fromCol + i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            planWorksheet.Cells[line + 1, fromCol + i].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }

        planWorksheet.Cells[line, fromCol + 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        planWorksheet.Cells[line, fromCol + 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        planWorksheet.Cells[6, 1, line + 1, selectedWeeks * 3 + 10].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);
    }

    internal static void DecorateFooter(ExcelWorksheet planWorksheet, int lineIndex)
    {
        planWorksheet.Cells[lineIndex + 2, 2].Value = "Теоретическое обучение: ";
        planWorksheet.Cells[lineIndex + 2, 2, lineIndex + 2, 19].Merge = true;

        planWorksheet.Cells[lineIndex + 3, 2].Value = "Зачетная неделя: ";
        planWorksheet.Cells[lineIndex + 3, 2, lineIndex + 3, 19].Merge = true;

        planWorksheet.Cells[lineIndex + 3, 20].Value = "Начальник Учебного управления";
        planWorksheet.Cells[lineIndex + 3, 20, lineIndex + 3, 39].Merge = true;

        planWorksheet.Cells[lineIndex + 3, 50].Value = "Декан";
        planWorksheet.Cells[lineIndex + 3, 50, lineIndex + 3, 59].Merge = true;

        planWorksheet.Cells[lineIndex + 4, 2].Value = "Экзаменационная сессия: ";
        planWorksheet.Cells[lineIndex + 4, 2, lineIndex + 4, 19].Merge = true;

        planWorksheet.Cells[lineIndex + 5, 2].Value = "Каникулы: ";
        planWorksheet.Cells[lineIndex + 5, 2, lineIndex + 5, 19].Merge = true;
    }
}