namespace Plans.Manager.BLL.Objects;

public class Discipline
{
    public Discipline(string name, string code, string parentCode, List<DisciplineData> data, Department department)
    {
        Name = name.Trim(); // TODO: Check for exceptions
        Code = code;
        ParentCode = parentCode;
        Data = data;
        Department = department;
        
        Sort();
    }

    public string Name { get; set; }
    public string Code { get; }
    public string ParentCode { get; }
    public Department Department { get; }
    
    public bool IsTrack { get; } // TODO: Include this property
    
    public bool IsElective { get; } // TODO: Include this property
    public List<DisciplineData> Data { get; internal set; }
    
    private void Sort()
    {
        List<DisciplineData> sortedDisciplineData = new List<DisciplineData>();
        foreach (DisciplineData disciplineData in Data)
        {
            string[] controlTypeVariants =
            {
                "Зачет",
                "Зачет с оценкой",
                "Экзамен",
                "Курсовой проект",
                "Курсовая работа"
            }; // TODO: Move to configuration file instead
            string[] hoursTypeVariants =
            {
                "Лекционные занятия",
                "Лабораторные занятия",
                "Практические занятия",
                "Семинарские занятия"
            }; // TODO: Move to configuration file instead

            if (controlTypeVariants.Contains(disciplineData.Type) &&
                disciplineData.Hours != 0 &&
                disciplineData.Course != 0 &&
                disciplineData.Semester != 0) // TODO: Remove magical number
                sortedDisciplineData.Add(disciplineData);
            else if (hoursTypeVariants.Contains(disciplineData.Type) &&
                     disciplineData.Hours >= 4.0 &&
                     disciplineData.Course != 0 &&
                     disciplineData.Semester != 0) // TODO: Remove magical number
                sortedDisciplineData.Add(disciplineData);
        }

        Data = sortedDisciplineData;
    }
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