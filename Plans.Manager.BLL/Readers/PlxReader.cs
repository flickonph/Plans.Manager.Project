using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml;

namespace Plans.Manager.BLL.Readers;

public class PlxReader : AReader
{
    public PlxReader(string path) : base(path) { }

    public PlanData GetPlanData()
    {
        // Public strings region
        /*string debugPath = _path;*/
        var planDirectionCode = GetXmlElements("ООП").First().GetAttribute("Шифр");
        var planDirectionName = GetPlanDirection();
        var tempYearOfAdmission = GetXmlElements("Планы").First().GetAttribute("УчебныйГод");
        var planYearOfAdmission = Convert.ToInt32(tempYearOfAdmission.Length > 4 ? tempYearOfAdmission.Substring(0, 4) : tempYearOfAdmission);

        // Disciplines region
        var planDisciplines = new List<Discipline>();
        IEnumerable<XmlElement> defaultDisciplinesXmlElements = GetXmlElements("ПланыСтроки");
        foreach (var defaultDisciplineXmlElement in defaultDisciplinesXmlElements)
        {
            var disciplineName = defaultDisciplineXmlElement.GetAttribute("Дисциплина");
            var disciplineCode = defaultDisciplineXmlElement.GetAttribute("Код");
            var parentCode = defaultDisciplineXmlElement.GetAttribute("КодРодителя");
            var departmentCode = defaultDisciplineXmlElement.GetAttribute("КодКафедры");
            var departmentName = string.Empty;
            if (!string.IsNullOrEmpty(departmentCode))
                departmentName = ConfigReader.GetDepartments()[departmentCode];

            var disciplineData = new List<DisciplineData>();
            var additionalDisciplineXmlElements = GetXmlElements("ПланыНовыеЧасы").Where(element => element.GetAttribute("КодОбъекта") == disciplineCode);
            foreach (var additionalDisciplineXmlElement in additionalDisciplineXmlElements)
            {
                var type = additionalDisciplineXmlElement.GetAttribute("КодВидаРаботы");
                var hours = Convert.ToDouble(additionalDisciplineXmlElement.GetAttribute("Количество"));
                var course = Convert.ToInt32(additionalDisciplineXmlElement.GetAttribute("Курс"));
                var semester = Convert.ToInt32(additionalDisciplineXmlElement.GetAttribute("Семестр"));
                /*var session = additionalDisciplineXmlElement.GetAttribute("Сессия");
                var weeks = additionalDisciplineXmlElement.GetAttribute("Недель");
                var days = additionalDisciplineXmlElement.GetAttribute("Дней");
                var hoursTypeCode = additionalDisciplineXmlElement.GetAttribute("КодТипаЧасов");*/

                var disciplineDataItem = new DisciplineData(type, hours, semester, course);
                IEnumerable<XmlElement> typesXmlElements = GetXmlElements("СправочникВидыРабот");
                foreach (var element in typesXmlElements)
                    if (element.GetAttribute("Код") == type)
                        disciplineDataItem.Type = element.GetAttribute("Название");

                disciplineData.Add(disciplineDataItem);
            }

            planDisciplines.Add(new Discipline(disciplineName, disciplineCode, parentCode, departmentCode, departmentName, disciplineData));
        }

        var planData = new PlanData(planDirectionCode, planDirectionName, planYearOfAdmission, planDisciplines);

        return planData;
    }

    private string GetPlanDirection()
    {
        string direction;
        var fullString = GetXmlElements("Планы").First().GetAttribute("Титул");
        var regex = new Regex("(\")([^\"]*)(\"[^\"]*)");
        try
        {
            direction = regex.Match(fullString).Groups[2].Value;
            direction = direction.Replace('ё', 'е');
            direction = direction.Trim();
        }
        catch (Exception e)
        {
            direction = string.Empty;
            Debug.Write(e);
        }

        return direction;
    }
}

public class Plan
{
    public Plan(string path, PlanData planData)
    {
        Path = path;
        PlanData = planData;
    }

    public string Path { get; }

    public PlanData PlanData { get; }
}

public class PlanData
{
    public PlanData(string code, string direction, int yearOfAdmission, List<Discipline> disciplines)
    {
        Code = code;
        Direction = direction;
        YearOfAdmission = yearOfAdmission;
        Disciplines = disciplines;

        BasicSort();
    }

    public string Code { get; }
    public string Direction { get; }
    public int YearOfAdmission { get; }
    public List<Discipline> Disciplines { get; set; }

    public void SortBySemester(int semester)
    {
        foreach (var discipline in Disciplines)
            discipline.DisciplineData = discipline.DisciplineData.Where(dictionary => dictionary.Semester == semester).ToList();
        Disciplines = Disciplines.Where(discipline => discipline.DisciplineData.Count != 0).ToList();
    }

    public void SortByCourse(int course)
    {
        foreach (var discipline in Disciplines)
            discipline.DisciplineData = discipline.DisciplineData.Where(dictionary => dictionary.Course == course).ToList();
        Disciplines = Disciplines.Where(discipline => discipline.DisciplineData.Count != 0).ToList();
    }

    public void SortByCourseAndSemester(int course, int semester, bool unify)
    {
        foreach (var discipline in Disciplines)
            discipline.DisciplineData = discipline.DisciplineData
                .Where(dictionary => dictionary.Course == course && dictionary.Semester == semester).ToList();
        Disciplines = Disciplines.Where(discipline => discipline.DisciplineData.Count != 0).ToList();
        
        var match = Disciplines.Where(x => !string.IsNullOrEmpty(x.ParentCode)).ToList();
        var parentCodes = match.Select(x => x.ParentCode).Distinct().ToList();
        if (match.Count != 0 && unify)
        {
            /*foreach (var parentCode in parentCodes)
            {
                var newDiscipline = match.First(x => x.ParentCode == parentCode);
                foreach (var disc in match.Where(x => x.ParentCode == parentCode))
                {
                    if (unifiedName == string.Empty)
                    {
                        unifiedName = disc.Name;
                    }
                    else
                    {
                        unifiedName += "/" + disc.Name;
                    }
                }
                newDiscipline.Name = unifiedName;
                Disciplines.RemoveAll(x => x.ParentCode == parentCode);
                Disciplines.RemoveAll(x => x.OwnCode == parentCode);
                Disciplines.Add(newDiscipline);
            }*/
            for (int i = 0; i < parentCodes.Count; i++)
            {
                var matchingCodeDisciplines = match.Where(discipline => discipline.ParentCode == parentCodes[i]).ToList();
                var unifiedDiscipline = matchingCodeDisciplines.First();
                if (matchingCodeDisciplines.Count > 1)
                {
                    for (int j = 1; j < matchingCodeDisciplines.Count; j++)
                    {
                        unifiedDiscipline.Name += $"/{matchingCodeDisciplines[j].Name}";
                    }
                }

                Disciplines.RemoveAll(x => x.ParentCode == parentCodes[i]);
                Disciplines.RemoveAll(x => x.OwnCode == parentCodes[i]);
                Disciplines.Add(unifiedDiscipline);
            }
        }
        else if (unify == false)
        {
            for (int i = 0; i < parentCodes.Count; i++)
            {
                Disciplines.RemoveAll(x => x.OwnCode == parentCodes[i]);
            }
        }
    }

    private void BasicSort()
    {
        foreach (var discipline in Disciplines)
        {
            var sortedData = new List<DisciplineData>();
            foreach (var dictionary in discipline.DisciplineData)
            {
                string[] firstMatch =
                {
                    "Зачет",
                    "Зачет с оценкой",
                    "Экзамен",
                    "Курсовой проект"
                };
                string[] secondMatch =
                {
                    "Лекционные занятия",
                    "Лабораторные занятия",
                    "Практические занятия",
                    "Семинарские занятия"
                };

                if (firstMatch.Contains(dictionary.Type) &&
                    dictionary.Hours != 0 &&
                    dictionary.Course != 0 &&
                    dictionary.Semester != 0) // TODO: Remove magical number
                    sortedData.Add(dictionary);
                else if (secondMatch.Contains(dictionary.Type) &&
                         dictionary.Hours > 2.0 &&
                         dictionary.Hours != 0 &&
                         dictionary.Course != 0 &&
                         dictionary.Semester != 0) // TODO: Remove magical number
                    sortedData.Add(dictionary);
            }

            discipline.DisciplineData = sortedData;
        }

        Disciplines = Disciplines.Where(d => d.DisciplineData.Count > 0).ToList();
        Disciplines = Disciplines.Where(
                discipline => discipline.DisciplineData.Where(
                    dict => 
                        dict.Type is "Лекционные занятия" or "Лабораторные занятия" or "Практические занятия" or "Семинарские занятия")
                    .ToList().Count > 0)
            .ToList();
    }
}

public class Discipline
{
    public Discipline(string name, string ownCode, string parentCode, string departmentCode, string departmentName, List<DisciplineData> disciplineData)
    {
        Name = name;
        OwnCode = ownCode;
        ParentCode = parentCode;
        DepartmentCode = departmentCode;
        DepartmentName = departmentName;
        DisciplineData = disciplineData;
    }

    public string Name { get; set; }
    public string OwnCode { get; }
    public string ParentCode { get; }
    public string DepartmentCode { get; }
    public string DepartmentName { get; }
    public List<DisciplineData> DisciplineData { get; set; }
}

public class DisciplineData
{
    public DisciplineData(string type, double hours, int semester, int course)
    {
        Type = type;
        Hours = hours;
        Semester = semester;
        Course = course;
    }
    
    public string Type { get; set; }
    public double Hours { get; }
    public int Semester { get; }
    public int Course { get; }
}