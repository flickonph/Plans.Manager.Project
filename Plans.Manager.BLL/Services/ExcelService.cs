using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Plans.Manager.BLL.Readers;

namespace Plans.Manager.BLL.Services;

public partial class ExcelHandler
{
    private readonly List<string> _plxFilesPaths;
    private readonly List<PlanGroupsPair> _plansGroupsPairs;
    private readonly int _studyYear;
    private static readonly List<string> AutoCompleteDisciplines = ConfigReader.GetAutoCompleteDisciplines();
    private static readonly string[] DisciplineControlType = { "Зачет", "Зачет с оценкой", "Экзамен", "Курсовая работа", "Курсовой проект" };

    public ExcelHandler(List<string> plxFilesPaths, int studyYear)
    {
        _plxFilesPaths = plxFilesPaths;
        _studyYear = studyYear;
        _plansGroupsPairs = GetPlansGroupsPairs();
    }

    public ExcelPackage PlansPackage(int course, int semester, int weeks, int studyYear)
    {
        // Package region
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var plansPackage = new ExcelPackage();
        plansPackage.Workbook.Properties.Author = "Отдел Планирования Учебного Процесса";
        plansPackage.Workbook.Properties.Title = "Рабочий учебный план";
        plansPackage.Workbook.Properties.Subject = "Экспорт Plans Manager Desktop";
        plansPackage.Workbook.Properties.Company = "РХТУ им. Д.И. Менделеева";
        plansPackage.Workbook.Properties.Created = DateTime.Now;

        var semesterDescription = semester switch
        {
            1 => $"на осенний семестр {studyYear}/{studyYear + 1} уч. год ({weeks} недель)",
            2 => $"на весенний семестр {studyYear}/{studyYear + 1} уч. год ({weeks} недель)",
            _ => string.Empty
        };
        
        List<PlanGroupsPair> plansGroupsPairs = _plansGroupsPairs.Where(x => x.EducationalGroups.First().Years == course).ToList();

        string localGroupName, localNumberOfStudents;
        foreach (var pair in plansGroupsPairs)
        {
            pair.Plan.PlanData.SortByCourseAndSemester(
                (studyYear - pair.Plan.PlanData.YearOfAdmission + 1), semester, true);
            if (pair.Plan.PlanData.Disciplines.Count == 0) continue;

            localGroupName = string.Empty;
            localNumberOfStudents = string.Empty;
            foreach (EducationalGroup educationalGroup in pair.EducationalGroups)
            {
                localGroupName += educationalGroup.Name + ", ";
                localNumberOfStudents += educationalGroup.NumberOfStudents + ", ";
            }

            localGroupName = localGroupName[..^2];
            localNumberOfStudents = localNumberOfStudents[..^2];
            
            var groupWorksheet = plansPackage.Workbook.Worksheets.Add($"{localGroupName}");
            groupWorksheet.View.ZoomScale = 75;

            FillGroupWorksheet(
                groupWorksheet,
                pair.Plan,
                weeks + 1, localGroupName, localNumberOfStudents, semesterDescription);
        }

        return plansPackage;
    }

    private static void FillGroupWorksheet(ExcelWorksheet groupWorksheet, Plan plan, int weeks, string localGroupName, string localNumberOfStudents, string semesterDescription)
    {
        List<Discipline> disciplines = plan.PlanData.Disciplines;
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
            
            FillControlType(groupWorksheet, disciplines[worksheetIndex], weeks, worksheetLine);
            
            FillHours(groupWorksheet, disciplines[worksheetIndex], weeks, worksheetLine);

            disciplineIndex++;
        }
        ExcelSheetStyle.ApplyPlansHeaderStyle(groupWorksheet, plan, semesterDescription, localGroupName, weeks, disciplineIndex, localNumberOfStudents);

        worksheetLine += 5;
        worksheetIndex += 5;

        foreach (var discipline in skippedDisciplines)
        {
            FillPrimaryData(groupWorksheet, discipline, worksheetLine, disciplineIndex);
            
            FillControlType(groupWorksheet, discipline, weeks, worksheetLine);
            
            FillHours(groupWorksheet, discipline, weeks, worksheetLine);
            
            worksheetLine += 3;
            worksheetIndex += 3;
            disciplineIndex++;
        }
        ExcelSheetStyle.ApplyPlansFooterStyle(groupWorksheet, worksheetLine - 2); // TODO: Remove -2
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
            discipline.DisciplineData.Find(disciplineData => 
                disciplineData.Type == "Лекционные занятия")?.Hours ?? 0;
        groupWorksheet.Cells[line, 5].Value =
            discipline.DisciplineData.Find(disciplineData => 
                disciplineData.Type == "Практические занятия")?.Hours ?? 0;
        groupWorksheet.Cells[line, 6].Value =
            discipline.DisciplineData.Find(disciplineData => 
                disciplineData.Type == "Лабораторные занятия")?.Hours ?? 0;
        
        groupWorksheet.Cells[line, 3].Formula = "=SUM(" + groupWorksheet.Cells[line, 4].Address + ":" +
                                           groupWorksheet.Cells[line, 6].Address + ")";

        // Styling of primary discipline's data
        groupWorksheet.Cells[line, 1, line, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        groupWorksheet.Cells[line, 1, line, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        for (var j = 1; j < 7; ++j)
        {
            groupWorksheet.Cells[line, j, line + 2, j].Merge = true;
            groupWorksheet.Cells[line, j, line + 2, j].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }
    }

    private static void FillControlType(ExcelWorksheet groupWorksheet, Discipline discipline, int weeks, int line)
    {
        for (var j = 0; j < 4; j++)
        {
            groupWorksheet.Cells[line, 7 + weeks * 3 + j, line + 2, 7 + weeks * 3 + j].Merge = true;
            groupWorksheet.Cells[line, 7 + weeks * 3 + j, line + 2, 7 + weeks * 3 + j].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
            groupWorksheet.Cells[line, 7 + weeks * 3 + j, line + 2, 7 + weeks * 3 + j].Style.HorizontalAlignment =
                ExcelHorizontalAlignment.Center;
            groupWorksheet.Cells[line, 7 + weeks * 3 + j, line + 2, 7 + weeks * 3 + j].Style.VerticalAlignment =
                ExcelVerticalAlignment.Center;
        }

        DisciplineData disciplineControlTypeMatch;
        try
        {
            disciplineControlTypeMatch = discipline.DisciplineData.First(disciplineData => DisciplineControlType.Contains(disciplineData.Type));
        }
        catch
        {
            disciplineControlTypeMatch = new DisciplineData("Ошибка", 0, 0, 0);
        }

        switch (disciplineControlTypeMatch.Type)
        {
            case "Зачет":
                groupWorksheet.Cells[line, 7 + weeks * 3].Value = disciplineControlTypeMatch.Hours;
                break;
            case "Зачет с оценкой":
                groupWorksheet.Cells[line, 7 + weeks * 3].Value = disciplineControlTypeMatch.Hours;
                groupWorksheet.Cells[line, 7 + weeks * 3].Style.Fill.SetBackground(Color.Gold);
                break;
            case "Экзамен":
                groupWorksheet.Cells[line, 7 + weeks * 3 + 1].Value = disciplineControlTypeMatch.Hours;
                break;
            case "Ошибка":
                groupWorksheet.Cells[line, 7 + weeks * 3].Value = 0;
                groupWorksheet.Cells[line, 7 + weeks * 3 + 1].Value = 0;
                groupWorksheet.Cells[line, 7 + weeks * 3 + 2].Value = 0;
                groupWorksheet.Cells[line, 7 + weeks * 3 + 3].Value = "Нет данных";
                groupWorksheet.Cells[line, 7 + weeks * 3, line, 7 + weeks * 3 + 3].Style.Fill.SetBackground(Color.Red);
                break;
        }

        var courseWork =
            discipline.DisciplineData.Find(disciplineData => disciplineData.Type is "Курсовая работа" or "Курсовой проект")?.Hours ?? 0;
        if (courseWork != 0) groupWorksheet.Cells[line, 7 + weeks * 3 + 2].Value = courseWork;
    }

    private static void FillHours(ExcelWorksheet groupWorksheet, Discipline discipline, int weeks, int line)
    {
        for (var j = 7; j <= weeks * 3 + 6; j += 3)
            groupWorksheet.Cells[line, j, line + 2, j + 2].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        
        if (!AutoCompleteDisciplines.Contains(discipline.Name)) return;
        
        double totalHours = discipline.DisciplineData.Find(disciplineData => disciplineData.Type == "Лекционные занятия")?.Hours ?? 0;
        double weekHours;
        if (totalHours != 0)
        {
            weekHours = totalHours / (weeks - 1);
            for (var j = 7; j <= (weeks - 1) * 3 + 6; j += 3)
            {
                groupWorksheet.Cells[line, j].Value = Math.Floor(weekHours);
            }
        }
                
        totalHours = discipline.DisciplineData.Find(disciplineData => disciplineData.Type == "Практические занятия")?.Hours ?? 0;
        if (totalHours != 0)
        {
            weekHours = totalHours / (weeks - 1);
            for (var j = 7; j <= (weeks - 1) * 3 + 6; j += 3)
            {
                groupWorksheet.Cells[line + 1, j + 1].Value = Math.Floor(weekHours);
            }
        }
                
        totalHours = discipline.DisciplineData.Find(disciplineData => disciplineData.Type == "Лабораторные занятия")?.Hours ?? 0;
        if (totalHours != 0)
        {
            weekHours = totalHours / (weeks - 1);
            for (var j = 7; j <= (weeks - 1) * 3 + 6; j += 3)
            {
                groupWorksheet.Cells[line + 2, j + 2].Value = Math.Floor(weekHours);
            }
        }
    }

    private List<PlanGroupsPair> GetPlansGroupsPairs()
    {
        var educationalGroups = ConfigReader.GetGroups();
        var plans = new List<Plan>();
        foreach (var plxFilePath in _plxFilesPaths)
        {
            var plxReader = new PlxReader(plxFilePath);
            plans.Add(new Plan(plxFilePath, plxReader.GetPlanData()));
        }

        var plansGroupsPairs = new List<PlanGroupsPair>();
        foreach (var plan in plans)
        {
            var matchingGroups = educationalGroups.Where(edGroup => 
                    edGroup.Direction == plan.PlanData.Direction && 
                    _studyYear - plan.PlanData.YearOfAdmission + 1 == edGroup.Years)
                .ToList();

            if (matchingGroups == null || matchingGroups.Count == 0)
            {
                continue;
                throw new InvalidDataException($"Для [{plan.Path}] не нашлось ни одной учебной группы");
            }
            plansGroupsPairs.Add(new PlanGroupsPair(plan, matchingGroups));
        }

        return plansGroupsPairs;
    }
}

public partial class ExcelHandler
{
     public ExcelPackage StudyLoad(int semesterOption)
    {
        // Package region
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var studyLoadPackage = new ExcelPackage();
        studyLoadPackage.Workbook.Properties.Author = @"Отдел Планирования Учебного Процесса";
        studyLoadPackage.Workbook.Properties.Title = @"Кафедральная нагрузка";
        studyLoadPackage.Workbook.Properties.Subject = @"Экспорт Plans Manager Desktop";
        studyLoadPackage.Workbook.Properties.Company = "РХТУ им. Д.И. Менделеева";
        studyLoadPackage.Workbook.Properties.Created = DateTime.Now;

        // String for header
        string edDescription = semesterOption switch
        {
            1 => $"на осенний семестр {_studyYear}/{_studyYear + 1} уч. год",
            2 => $"на весенний семестр {_studyYear}/{_studyYear + 1} уч. год",
            _ => string.Empty
        };
        
        // All departments
        var departments = ConfigReader.GetDepartments();

        // Loop for each department
        foreach (var department in departments)
        {
            // Worksheet creation
            ExcelWorksheet worksheet = studyLoadPackage.Workbook.Worksheets.Add(department.Key);
            ExcelSheetStyle.ApplyLoadHeaderStyle(worksheet, edDescription, department);

            int globalCounter = 3;
            foreach (var planGroupsPair in _plansGroupsPairs)
            {
                foreach (var educationalGroup in planGroupsPair.EducationalGroups)
                {
                    int from = globalCounter;
                    worksheet.Cells[globalCounter, 1].Value = educationalGroup.Name;
                    worksheet.Cells[globalCounter, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    worksheet.Cells[globalCounter, 1].Style.Fill.SetBackground(Color.LightGray);
                    
                    worksheet.Cells[globalCounter, 7].Value = educationalGroup.NumberOfStudents;
                    worksheet.Cells[globalCounter, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    worksheet.Cells[globalCounter, 7].Style.Fill.SetBackground(Color.LightGray);
                    planGroupsPair.Plan.PlanData.SortByCourseAndSemester(educationalGroup.Years, semesterOption, false);
                    List<Discipline> requiredDisciplines =
                        planGroupsPair.Plan.PlanData.Disciplines.Where(discipline => discipline.DepartmentCode == department.Key).ToList();
                    foreach (var discipline in requiredDisciplines)
                    {
                        worksheet.Cells[globalCounter, 2].Value = discipline.Name;
                        worksheet.Cells[globalCounter, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                        
                        worksheet.Cells[globalCounter, 3].Value = discipline.DisciplineData.Find(dictionary => dictionary.Type == "Лекционные занятия")?.Hours ?? 0;
                        worksheet.Cells[globalCounter, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                        
                        worksheet.Cells[globalCounter, 4].Value = discipline.DisciplineData.Find(dictionary => dictionary.Type == "Практические занятия")?.Hours ?? 0;
                        worksheet.Cells[globalCounter, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                        
                        worksheet.Cells[globalCounter, 5].Value = discipline.DisciplineData.Find(dictionary => dictionary.Type == "Лабораторные занятия")?.Hours ?? 0;
                        worksheet.Cells[globalCounter, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                        
                        worksheet.Cells[globalCounter, 6].Value =
                            discipline.DisciplineData.Find(disciplineData => disciplineData.Type is "Курсовой проект" or "Экзамен" or "Зачет" or "Зачет с оценкой")?.Type ?? "Ошибка";
                        worksheet.Cells[globalCounter, 6].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                        
                        globalCounter++;
                    }

                    if (globalCounter > from)
                    {
                        int to = globalCounter - 1;
                        worksheet.Cells[from, 1, to, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    }
                }
            }

            worksheet.Cells[globalCounter, 1, globalCounter, 7].Value = ""; // TODO: Remove this
            worksheet.Cells[globalCounter, 1, globalCounter, 7].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black); // TODO: Remove this
            worksheet.Cells[globalCounter, 1, globalCounter, 7].Style.Fill.SetBackground(Color.White); // TODO: Remove this
        }

        return studyLoadPackage;
    }
}

public class PlanGroupsPair
{
    public PlanGroupsPair(Plan plan, List<EducationalGroup> educationalGroups)
    {
        Plan = plan;
        EducationalGroups = educationalGroups;
    }
    public Plan Plan { get; }
    public List<EducationalGroup> EducationalGroups { get; }
}

public static class ExcelSheetStyle
{
    public static void ApplyPlansHeaderStyle(ExcelWorksheet currentWorksheet, Plan plan, string edDescription, string edGroup, int edWeeks, int numberOfDisciplines, string numberOfStudents)
    {
        int yearOfAdmission = plan.PlanData.YearOfAdmission;
        string edDirection = plan.PlanData.Code;


        // Font formatting
        currentWorksheet.Cells[1, 1, 100, 100].Style.Font.Bold = true;
        currentWorksheet.Cells[1, 1, 100, 100].Style.Font.Name = "Times New Roman";
        currentWorksheet.Cells[1, 1, 100, 100].Style.Font.Size = 12;

        // Rows formatting
        currentWorksheet.Rows[1].Height = 6.00;
        currentWorksheet.Rows[2].Height = 24.00;
        currentWorksheet.Rows[3].Height = 6.00;
        currentWorksheet.Rows[4].Height = 14.40;
        currentWorksheet.Rows[5].Height = 19.20;
        currentWorksheet.Rows[6].Height = 24.00;
        currentWorksheet.Rows[7].Height = 57.60;
        for (var i = 8; i <= 60; i++) currentWorksheet.Rows[i].Height = 15.00;

        // Columns formatting
        currentWorksheet.Columns[1].Width = 4.78;
        currentWorksheet.Columns[2].Width = 43.67;
        currentWorksheet.Columns[3].Width = 5.89;
        currentWorksheet.Columns[4].Width = 5.89;
        currentWorksheet.Columns[5].Width = 5.89;
        currentWorksheet.Columns[6].Width = 5.89;
        int col;
        for (col = 7; col <= edWeeks * 3 + 6; col++) currentWorksheet.Columns[col].Width = 3.00;

        for (var i = 1; i <= 3; i++)
        {
            currentWorksheet.Columns[col].Width = 4.67;
            col++;
        }

        currentWorksheet.Columns[col].Width = 23.00;

        // Filling of the worksheet header
        currentWorksheet.Cells[2, 2].Value = $"Год приёма: {yearOfAdmission}";
        currentWorksheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        currentWorksheet.Cells[2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

        currentWorksheet.Cells[2, 3].Value = "Рабочий учебный план";
        currentWorksheet.Cells[2, 3].Style.Font.Size = 22;
        currentWorksheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[2, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        currentWorksheet.Cells[2, 3, 2, 43].Merge = true;

        currentWorksheet.Cells[2, 48].Value = $"Направление: {edDirection}";
        currentWorksheet.Cells[2, 48].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        currentWorksheet.Cells[2, 48].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        currentWorksheet.Cells[2, 48, 2, 61].Merge = true;

        currentWorksheet.Cells[4, 3].Value = edDescription;
        currentWorksheet.Cells[4, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[4, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        currentWorksheet.Cells[4, 3, 4, 43].Merge = true;

        currentWorksheet.Cells[5, 2].Value = $"Студентов: {numberOfStudents}";
        currentWorksheet.Cells[5, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        currentWorksheet.Cells[5, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        currentWorksheet.Cells[5, 2, 5, 8].Merge = true;

        currentWorksheet.Cells[5, 48].Value = $"Группы: {edGroup}";
        currentWorksheet.Cells[5, 48].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        currentWorksheet.Cells[5, 48].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
        currentWorksheet.Cells[5, 48, 5, 61].Merge = true;

        // Filling of the table header
        currentWorksheet.Cells[6, 1].Value = "№ п/п";
        currentWorksheet.Cells[6, 1].Style.WrapText = true;
        currentWorksheet.Cells[6, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, 2].Value = "Дисциплина";
        currentWorksheet.Cells[6, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, 3].Value = "Ауд.";
        currentWorksheet.Cells[6, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, 4].Value = "Лек.";
        currentWorksheet.Cells[6, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, 5].Value = "Сем.";
        currentWorksheet.Cells[6, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, 6].Value = "Лаб.";
        currentWorksheet.Cells[6, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        for (var i = 1; i < 7; ++i)
        {
            currentWorksheet.Cells[6, i, 7, i].Merge = true;
            currentWorksheet.Cells[6, i, 7, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            currentWorksheet.Cells[8, i].Value = i;
            currentWorksheet.Cells[8, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            currentWorksheet.Cells[8, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[8, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        currentWorksheet.Cells[6, 7].Value = "Распределение по неделям семестра (лекц./сем./лаб.)";
        currentWorksheet.Cells[6, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        currentWorksheet.Cells[6, 7, 6, edWeeks * 3 + 6].Merge = true;
        currentWorksheet.Cells[6, 7, 6, edWeeks * 3 + 6].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);

        for (var i = 0; i < edWeeks; ++i)
        {
            col = 7 + i * 3;
            currentWorksheet.Cells[7, col].Value = i + 1;
            currentWorksheet.Cells[7, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[7, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            currentWorksheet.Cells[7, col, 7, col + 2].Merge = true;
            currentWorksheet.Cells[7, col, 7, col + 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);

            currentWorksheet.Cells[8, col, 8, col + 2].Merge = true;
            currentWorksheet.Cells[8, col].Value = i + 7;
            currentWorksheet.Cells[8, col, 8, col + 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            currentWorksheet.Cells[8, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[8, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        var nextColumn = 7 + edWeeks * 3;

        currentWorksheet.Cells[6, nextColumn].Value = "Контроль";
        currentWorksheet.Cells[6, nextColumn].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        currentWorksheet.Cells[6, nextColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, nextColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[7, nextColumn].Value = "Зачет";
        currentWorksheet.Cells[7, nextColumn].Style.TextRotation = 90;
        currentWorksheet.Cells[7, nextColumn].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[7, nextColumn].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[7, nextColumn + 1].Value = "Экзамен";
        currentWorksheet.Cells[7, nextColumn + 1].Style.TextRotation = 90;
        currentWorksheet.Cells[7, nextColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[7, nextColumn + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[7, nextColumn + 2].Value = "К/р (пр-т)";
        currentWorksheet.Cells[7, nextColumn + 2].Style.TextRotation = 90;
        currentWorksheet.Cells[7, nextColumn + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[7, nextColumn + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, nextColumn + 3].Value = "Примечание";
        currentWorksheet.Cells[6, nextColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[6, nextColumn + 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[6, nextColumn, 6, nextColumn + 2].Merge = true;
        currentWorksheet.Cells[6, nextColumn, 6, nextColumn + 2].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        currentWorksheet.Cells[6, nextColumn + 3, 7, nextColumn + 3].Merge = true;
        currentWorksheet.Cells[6, nextColumn + 3, 7, nextColumn + 3].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        for (var i = nextColumn; i <= nextColumn + 3; i++)
        {
            currentWorksheet.Cells[8, i].Value = edWeeks + 7 + i - nextColumn;
            currentWorksheet.Cells[8, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            currentWorksheet.Cells[8, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[8, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        // Filling of the table bottom part
        var line = 9 + numberOfDisciplines * 3;
        currentWorksheet.Rows[line].Height = 27.60;
        currentWorksheet.Rows[line + 1].Height = 27.60;
        currentWorksheet.Cells[line, 2].Value = "Всего часов в неделю:";
        currentWorksheet.Cells[line, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        currentWorksheet.Cells[line, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        currentWorksheet.Cells[line, 3].Formula = "=SUM(" + currentWorksheet.Cells[9, 3].Address + ":" +
                                                  currentWorksheet.Cells[line - 1, 3].Address + ")";
        currentWorksheet.Cells[line, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[line, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        currentWorksheet.Cells[line + 1, 2].Value = "В том числе лекций:";
        currentWorksheet.Cells[line + 1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        currentWorksheet.Cells[line + 1, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        currentWorksheet.Cells[line + 1, 3].Formula = "=SUM(" + currentWorksheet.Cells[9, 4].Address + ":" +
                                                      currentWorksheet.Cells[line - 1, 4].Address + ")";
        currentWorksheet.Cells[line + 1, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        currentWorksheet.Cells[line + 1, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        for (var i = 1; i <= 6; i++)
        {
            currentWorksheet.Cells[line, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            currentWorksheet.Cells[line + 1, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }

        int fromCol;
        for (fromCol = 7; fromCol <= edWeeks * 3 + 6; fromCol += 3)
        {
            currentWorksheet.Cells[line, fromCol].Formula = "=SUM(" + currentWorksheet.Cells[9, fromCol].Address + ":" +
                                                            currentWorksheet.Cells[line - 1, fromCol + 2].Address + ")";
            currentWorksheet.Cells[line, fromCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[line, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            currentWorksheet.Cells[line, fromCol, line, fromCol + 2].Merge = true;
            currentWorksheet.Cells[line, fromCol, line, fromCol + 2].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);

            currentWorksheet.Cells[line + 1, fromCol].Formula = "=SUM(" + currentWorksheet.Cells[9, fromCol].Address +
                                                                ":" + currentWorksheet.Cells[line - 1, fromCol]
                                                                    .Address + ")";
            currentWorksheet.Cells[line + 1, fromCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[line + 1, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            currentWorksheet.Cells[line + 1, fromCol, line + 1, fromCol + 2].Merge = true;
            currentWorksheet.Cells[line + 1, fromCol, line + 1, fromCol + 2].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }

        for (var i = 0; i < 3; i++)
        {
            currentWorksheet.Cells[line, fromCol + i].Formula = "=SUM(" +
                                                                currentWorksheet.Cells[9, fromCol + i].Address + ":" +
                                                                currentWorksheet.Cells[line - 1, fromCol + i].Address +
                                                                ")";
            currentWorksheet.Cells[line, fromCol + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            currentWorksheet.Cells[line, fromCol + i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            currentWorksheet.Cells[line, fromCol + i].Merge = true;
            currentWorksheet.Cells[line, fromCol + i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            currentWorksheet.Cells[line + 1, fromCol + i].Style.Border
                .BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }

        currentWorksheet.Cells[line, fromCol + 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        currentWorksheet.Cells[line, fromCol + 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        currentWorksheet.Cells[6, 1, line + 1, edWeeks * 3 + 10].Style.Border
            .BorderAround(ExcelBorderStyle.Medium, Color.Black);
    }

    public static void ApplyPlansFooterStyle(ExcelWorksheet currentWorksheet, int line)
    {
        currentWorksheet.Cells[line + 2, 2].Value = "Теоретическое обучение: ";
        currentWorksheet.Cells[line + 2, 2, line + 2, 19].Merge = true;

        currentWorksheet.Cells[line + 3, 2].Value = "Зачетная неделя: ";
        currentWorksheet.Cells[line + 3, 2, line + 3, 19].Merge = true;

        currentWorksheet.Cells[line + 3, 20].Value = "Начальник Учебного управления";
        currentWorksheet.Cells[line + 3, 20, line + 3, 39].Merge = true;

        currentWorksheet.Cells[line + 3, 50].Value = "Декан";
        currentWorksheet.Cells[line + 3, 50, line + 3, 59].Merge = true;

        currentWorksheet.Cells[line + 4, 2].Value = "Экзаменационная сессия: ";
        currentWorksheet.Cells[line + 4, 2, line + 4, 19].Merge = true;

        currentWorksheet.Cells[line + 5, 2].Value = "Каникулы: ";
        currentWorksheet.Cells[line + 5, 2, line + 5, 19].Merge = true;
    }

    public static void ApplyLoadHeaderStyle(ExcelWorksheet worksheet, string edDescription, KeyValuePair<string, string> department)
    {
        // Font formatting
            worksheet.Cells[1, 1, 100, 100].Style.Font.Bold = true;
            worksheet.Cells[1, 1, 100, 100].Style.Font.Name = "Times New Roman";
            worksheet.Cells[1, 1, 100, 100].Style.Font.Size = 10;
            worksheet.Cells[1, 1, 100, 100].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            
            worksheet.Columns[1].Width = 34.00;
            worksheet.Columns[2].Width = 66.00;
            worksheet.Columns[3].Width = 8.00;
            worksheet.Columns[4].Width = 8.00;
            worksheet.Columns[5].Width = 8.00;
            worksheet.Columns[6].Width = 17.00;
            worksheet.Columns[7].Width = 14.00;
            
            // Main header region
            worksheet.Cells[1, 1].Value = "Группа";
            worksheet.Cells[1, 1, 2, 1].Merge = true;
            worksheet.Cells[1, 1, 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1, 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 1, 2, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            worksheet.Cells[1, 2].Value = "Дисциплина";
            worksheet.Cells[1, 2, 2, 2].Merge = true;
            worksheet.Cells[1, 2, 2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 2, 2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 2, 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            worksheet.Cells[1, 3].Value = "Часы";
            worksheet.Cells[1, 3, 1, 5].Merge = true;
            worksheet.Cells[1, 3, 1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 3, 1, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 3, 1, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            worksheet.Cells[2, 3].Value = "Лек.";
            worksheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[2, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            worksheet.Cells[2, 4].Value = "Сем.";
            worksheet.Cells[2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[2, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            worksheet.Cells[2, 5].Value = "Лаб.";
            worksheet.Cells[2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[2, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            worksheet.Cells[1, 6].Value = "Форма контроля";
            worksheet.Cells[1, 6, 2, 6].Merge = true;
            worksheet.Cells[1, 6, 2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 6, 2, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 6, 2, 6].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            worksheet.Cells[1, 7].Value = "Численность";
            worksheet.Cells[1, 7, 2, 7].Merge = true;
            worksheet.Cells[1, 7, 2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 7, 2, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 7, 2, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            worksheet.Cells[1, 8].Value = $"Кафедра: {department.Value}";
            worksheet.Cells[1, 8, 1, 12].Merge = true;
            worksheet.Cells[2, 8].Value = $"Код кафедры: {department.Key}";
            worksheet.Cells[2, 8, 2, 12].Merge = true;
            worksheet.Cells[3, 8].Value = edDescription;
            worksheet.Cells[3, 8, 3, 12].Merge = true;
            worksheet.Cells[1, 8, 3, 12].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
    }
}