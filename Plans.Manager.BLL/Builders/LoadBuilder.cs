using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Plans.Manager.BLL.Objects;
using Plans.Manager.BLL.Providers;
using Plans.Manager.BLL.Readers;

namespace Plans.Manager.BLL.Builders;

public class LoadBuilder : ABuilder
{
    private readonly List<PlanGroupPair> _allPlanGroupPairs;
    private readonly int _selectedYear;
    private readonly List<Department> _departments;
    private readonly int _selectedSemester;

    public LoadBuilder(List<string> plxFilesPaths, int selectedYear, int selectedSemester) : base(plxFilesPaths, selectedYear)
    {
        _selectedYear = selectedYear;
        _selectedSemester = selectedSemester;
        _allPlanGroupPairs = GetAllPlanGroupPairs();
        _departments = XmlReader.GetDepartments();
    }

    public List<ExcelPackage> LoadExcelPackages()
    {
        return _departments.Select(department => BuildLoadExcelPackage(department)).ToList();
    } 
    public ExcelPackage BuildLoadExcelPackage(Department department)
    {
        // Package region
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelPackage loadPackage = new ExcelPackage();
        loadPackage.Workbook.Properties.Author = MetadataProvider.Author;
        loadPackage.Workbook.Properties.Title = MetadataProvider.LoadTitle;
        loadPackage.Workbook.Properties.Subject = MetadataProvider.Subject;
        loadPackage.Workbook.Properties.Company = MetadataProvider.Company;
        loadPackage.Workbook.Properties.Created = MetadataProvider.Created;
        loadPackage.Workbook.Properties.Comments = $"Нагрузка (№{department.Code}) {department.Name.Replace('\"', '\'')}";

        // String for header
        string semesterDescription = _selectedSemester switch
        {
            1 => $"на осенний семестр {_selectedYear}/{_selectedYear + 1} уч. год",
            2 => $"на весенний семестр {_selectedYear}/{_selectedYear + 1} уч. год",
            _ => "ошибка"
        };

        // Worksheet creation
        ExcelWorksheet worksheet = loadPackage.Workbook.Worksheets.Add(department.Code);
        LoadDecorator.DecorateHeader(worksheet, semesterDescription, department);

        int globalCounter = 3;
        foreach (PlanGroupPair planGroupPair in _allPlanGroupPairs)
        {
            foreach (Group educationalGroup in planGroupPair.Groups)
            {
                int from = globalCounter;
                
                planGroupPair.Plan.Data.SortByCourseAndSemester(educationalGroup.Years, _selectedSemester, false);
                List<Discipline> requiredDisciplines =
                    planGroupPair.Plan.Data.Disciplines.Where(discipline => discipline.Department.Code == department.Code).ToList();
                foreach (Discipline discipline in requiredDisciplines)
                {
                    worksheet.Cells[globalCounter, 1].Value = educationalGroup.Name;
                    worksheet.Cells[globalCounter, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    worksheet.Cells[globalCounter, 1].Style.Fill.SetBackground(Color.LightGray);
                    
                    worksheet.Cells[globalCounter, 2].Value = discipline.Name;
                    worksheet.Cells[globalCounter, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    
                    worksheet.Cells[globalCounter, 3].Value = discipline.Data.Find(dictionary => dictionary.Type == "Лекционные занятия")?.Hours ?? 0;
                    worksheet.Cells[globalCounter, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    
                    worksheet.Cells[globalCounter, 4].Value = discipline.Data.Find(dictionary => dictionary.Type == "Практические занятия")?.Hours ?? 0;
                    worksheet.Cells[globalCounter, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    
                    worksheet.Cells[globalCounter, 5].Value = discipline.Data.Find(dictionary => dictionary.Type == "Лабораторные занятия")?.Hours ?? 0;
                    worksheet.Cells[globalCounter, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    
                    worksheet.Cells[globalCounter, 6].Value =
                        discipline.Data.Find(disciplineData => ControlType.Contains(disciplineData.Type))?.Type ?? "Ошибка";
                    worksheet.Cells[globalCounter, 6].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    
                    worksheet.Cells[globalCounter, 7].Value = educationalGroup.NumberOfStudents;
                    worksheet.Cells[globalCounter, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                    worksheet.Cells[globalCounter, 7].Style.Fill.SetBackground(Color.LightGray);
                    
                    globalCounter++;
                }

                if (globalCounter > from)
                {
                    int to = globalCounter - 1;
                    worksheet.Cells[from, 1, to, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                }
            }
        }

        worksheet.Cells[globalCounter-1, 1, globalCounter-1, 7].Value = ""; // TODO: Remove this
        worksheet.Cells[globalCounter-1, 1, globalCounter-1, 7].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black); // TODO: Remove this
        worksheet.Cells[globalCounter-1, 1, globalCounter-1, 7].Style.Fill.SetBackground(Color.White); // TODO: Remove this
        
        return loadPackage;
    }
}

file static class LoadDecorator
{
    internal static void DecorateHeader(ExcelWorksheet loadWorksheet, string edDescription, Department department)
    {
        // Font formatting
            loadWorksheet.Cells[1, 1, 100, 100].Style.Font.Bold = true;
            loadWorksheet.Cells[1, 1, 100, 100].Style.Font.Name = "Times New Roman";
            loadWorksheet.Cells[1, 1, 100, 100].Style.Font.Size = 10;
            loadWorksheet.Cells[1, 1, 100, 100].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            
            loadWorksheet.Columns[1].Width = 34.00;
            loadWorksheet.Columns[2].Width = 66.00;
            loadWorksheet.Columns[3].Width = 8.00;
            loadWorksheet.Columns[4].Width = 8.00;
            loadWorksheet.Columns[5].Width = 8.00;
            loadWorksheet.Columns[6].Width = 17.00;
            loadWorksheet.Columns[7].Width = 14.00;
            
            // Main header region
            loadWorksheet.Cells[1, 1].Value = "Группа";
            loadWorksheet.Cells[1, 1, 2, 1].Merge = true;
            loadWorksheet.Cells[1, 1, 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 1, 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 1, 2, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 2].Value = "Дисциплина";
            loadWorksheet.Cells[1, 2, 2, 2].Merge = true;
            loadWorksheet.Cells[1, 2, 2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 2, 2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 2, 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 3].Value = "Часы";
            loadWorksheet.Cells[1, 3, 1, 5].Merge = true;
            loadWorksheet.Cells[1, 3, 1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 3, 1, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 3, 1, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[2, 3].Value = "Лек.";
            loadWorksheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[2, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[2, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            loadWorksheet.Cells[2, 4].Value = "Сем.";
            loadWorksheet.Cells[2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[2, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[2, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            loadWorksheet.Cells[2, 5].Value = "Лаб.";
            loadWorksheet.Cells[2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[2, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[2, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 6].Value = "Форма контроля";
            loadWorksheet.Cells[1, 6, 2, 6].Merge = true;
            loadWorksheet.Cells[1, 6, 2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 6, 2, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 6, 2, 6].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 7].Value = "Численность";
            loadWorksheet.Cells[1, 7, 2, 7].Merge = true;
            loadWorksheet.Cells[1, 7, 2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 7, 2, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 7, 2, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 8].Value = $"Кафедра: {department.Name}";
            loadWorksheet.Cells[1, 8].Style.Font.Size = 10; // TODO: Change
            loadWorksheet.Cells[1, 8, 1, 12].Merge = true;
            loadWorksheet.Cells[2, 8].Value = $"Код кафедры: {department.Code}";
            loadWorksheet.Cells[2, 8, 2, 12].Merge = true;
            loadWorksheet.Cells[3, 8].Value = edDescription;
            loadWorksheet.Cells[3, 8, 3, 12].Merge = true;
            loadWorksheet.Cells[1, 8, 3, 12].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
    }
}