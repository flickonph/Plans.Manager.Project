namespace Plans.Manager.BLL.Objects;

public class Plan
{
    public Plan(string path, PlanData data)
    {
        Path = path;
        Data = data;
    }

    public string Path { get; }

    public PlanData Data { get; }
}

public class PlanData
{
    public PlanData(string code, string direction, int yearOfAdmission, List<Discipline> disciplines)
    {
        Code = code;
        Direction = direction;
        YearOfAdmission = yearOfAdmission;
        Disciplines = disciplines;
        
        Sort();
    }

    public string Code { get; }
    public string Direction { get; }
    public int YearOfAdmission { get; }
    public List<Discipline> Disciplines { get; private set; }

    private void Sort()
    {
        string[] hoursTypeVariants =
        {
            "Лекционные занятия",
            "Лабораторные занятия",
            "Практические занятия",
            "Семинарские занятия"
        }; // TODO: Move to configuration file instead
        
        Disciplines = Disciplines.Where(
                discipline => discipline.Data.Where(
                        disciplineData => hoursTypeVariants.Contains(disciplineData.Type)).ToList().Count > 0).ToList();
        Disciplines = Disciplines.Where(d => d.Data.Count > 0).ToList();
    }

    public void SortBySemester(int semester)
    {
        foreach (Discipline discipline in Disciplines)
            discipline.Data = discipline.Data.Where(dictionary => dictionary.Semester == semester).ToList();
        Disciplines = Disciplines.Where(discipline => discipline.Data.Count != 0).ToList();
    }

    public void SortByCourse(int course)
    {
        foreach (Discipline discipline in Disciplines)
            discipline.Data = discipline.Data.Where(dictionary => dictionary.Course == course).ToList();
        Disciplines = Disciplines.Where(discipline => discipline.Data.Count != 0).ToList();
    }

    public void SortByCourseAndSemester(int course, int semester, bool unify)
    {
        foreach (Discipline discipline in Disciplines)
            discipline.Data = discipline.Data
                .Where(dictionary => dictionary.Course == course && dictionary.Semester == semester).ToList();
        Disciplines = Disciplines.Where(discipline => discipline.Data.Count != 0).ToList();
        
        List<Discipline> match = Disciplines.Where(x => !string.IsNullOrEmpty(x.ParentCode)).ToList();
        List<string> parentCodes = match.Select(x => x.ParentCode).Distinct().ToList();
        if (match.Count != 0 && unify)
        {
            foreach (var parentCode in parentCodes)
            {
                List<Discipline> matchingCodeDisciplines = match.Where(discipline => discipline.ParentCode == parentCode).ToList();
                Discipline unifiedDiscipline = matchingCodeDisciplines.First();
                if (matchingCodeDisciplines.Count > 1)
                {
                    for (int j = 1; j < matchingCodeDisciplines.Count; j++)
                    {
                        unifiedDiscipline.Name += $"/{matchingCodeDisciplines[j].Name}";
                    }
                }

                Disciplines.RemoveAll(x => x.ParentCode == parentCode);
                Disciplines.RemoveAll(x => x.Code == parentCode);
                Disciplines.Add(unifiedDiscipline);
            }
        }
        else if (unify == false)
        {
            foreach (var parentCode in parentCodes)
            {
                Disciplines.RemoveAll(discipline => discipline.Code == parentCode);
            }
        }
    }
}