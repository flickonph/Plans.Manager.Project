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
        return _departments.Select(BuildLoadExcelPackage).ToList();
    }

    private ExcelPackage BuildLoadExcelPackage(Department department)
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
        ExcelWorksheet loadWorksheet = loadPackage.Workbook.Worksheets.Add(department.Code);
        loadWorksheet.View.ZoomScale = 70;
        LoadDecorator.DecorateHeader(loadWorksheet, semesterDescription, department);

        int currentRow = 3;
        foreach (PlanGroupPair planGroupPair in _allPlanGroupPairs)
        {
            foreach (Group educationalGroup in planGroupPair.Groups)
            {
                int from = currentRow;
                
                planGroupPair.Plan.Data.SortByCourseAndSemester(educationalGroup.Years, _selectedSemester, false);
                List<Discipline> requiredDisciplines =
                    planGroupPair.Plan.Data.Disciplines.Where(discipline => discipline.Department.Code == department.Code).ToList();
                foreach (Discipline discipline in requiredDisciplines)
                {
                    loadWorksheet.Cells[currentRow, 1].Value = educationalGroup.Name;

                    loadWorksheet.Cells[currentRow, 2].Value = discipline.Name;
                    
                    loadWorksheet.Cells[currentRow, 3].Value = discipline.Data.Find(dictionary => dictionary.Type == "Лекционные занятия")?.Hours ?? 0;
                    
                    loadWorksheet.Cells[currentRow, 4].Value = discipline.Data.Find(dictionary => dictionary.Type == "Практические занятия")?.Hours ?? 0;

                    loadWorksheet.Cells[currentRow, 5].Value = discipline.Data.Find(dictionary => dictionary.Type == "Лабораторные занятия")?.Hours ?? 0;

                    List<DisciplineData> disciplineData =
                        discipline.Data.Where(disciplineData => ControlType.Contains(disciplineData.Type)).ToList();
                    if (disciplineData.Count > 0)
                    {
                        foreach (DisciplineData data in disciplineData)
                        {
                            if (PrimaryType.Contains(data.Type))
                            {
                                loadWorksheet.Cells[currentRow, 6].Value = data.Type;
                            }
                            else if (AdditionalType.Contains(data.Type))
                            {
                                loadWorksheet.Cells[currentRow, 7].Value = data.Type;
                            }
                        }
                    }

                    loadWorksheet.Cells[currentRow, 8].Value = educationalGroup.NumberOfStudents;
                    
                    LoadDecorator.DecorateRow(loadWorksheet, currentRow);

                    currentRow++;
                }

                if (currentRow > from)
                {
                    int to = currentRow - 1;
                    loadWorksheet.Cells[from, 1, to, 8].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                }
            }
        }

        loadWorksheet.Cells[currentRow-1, 1, currentRow-1, 8].Value = ""; // TODO: Remove this
        loadWorksheet.Cells[currentRow-1, 1, currentRow-1, 8].Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black); // TODO: Remove this
        loadWorksheet.Cells[currentRow-1, 1, currentRow-1, 8].Style.Fill.SetBackground(Color.LightGray); // TODO: Remove this
        
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
            loadWorksheet.Cells[1, 1, 200, 10].Style.Font.Size = 10;
            loadWorksheet.Cells[1, 1, 200, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            
            loadWorksheet.Rows[1].Height = 25.00;
            loadWorksheet.Columns[1].Width = 34.00;
            loadWorksheet.Columns[2].Width = 66.00;
            loadWorksheet.Columns[2].Style.WrapText = true;
            loadWorksheet.Columns[3].Width = 8.00;
            loadWorksheet.Columns[4].Width = 8.00;
            loadWorksheet.Columns[5].Width = 8.00;
            loadWorksheet.Columns[6].Width = 20.00;
            loadWorksheet.Columns[7].Width = 20.00;
            loadWorksheet.Columns[8].Width = 14.00;
            
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
            
            loadWorksheet.Cells[1, 7].Value = "К/р (пр-т)";
            loadWorksheet.Cells[1, 7, 2, 7].Merge = true;
            loadWorksheet.Cells[1, 7, 2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 7, 2, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 7, 2, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 8].Value = "Численность";
            loadWorksheet.Cells[1, 8, 2, 8].Merge = true;
            loadWorksheet.Cells[1, 8, 2, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            loadWorksheet.Cells[1, 8, 2, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[1, 8, 2, 8].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
            
            loadWorksheet.Cells[1, 9].Value = $"Кафедра: {department.Name}";
            loadWorksheet.Cells[1, 9].Style.WrapText = true;
            loadWorksheet.Cells[1, 9, 1, 13].Merge = true;
            loadWorksheet.Cells[1, 9, 1, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            loadWorksheet.Cells[2, 9].Value = $"Код кафедры: {department.Code}";
            loadWorksheet.Cells[2, 9, 2, 13].Merge = true;
            loadWorksheet.Cells[3, 9].Value = edDescription;
            loadWorksheet.Cells[3, 9, 3, 13].Merge = true;
            loadWorksheet.Cells[1, 9, 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
    }

    internal static void DecorateRow(ExcelWorksheet loadWorksheet, int currentRow)
    {
        for (int i = 1; i <= 8; i++)
        {
            loadWorksheet.Cells[currentRow, i].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }
        
        loadWorksheet.Cells[currentRow, 1].Style.Fill.SetBackground(Color.LightGray);
        loadWorksheet.Cells[currentRow, 8].Style.Fill.SetBackground(Color.LightGray);
        loadWorksheet.Rows[currentRow].Height = 26;
        loadWorksheet.Rows[currentRow].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
    }
}