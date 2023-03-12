namespace Plans.Manager.BLL.Objects;

public class Group
{
    public Group(string name, string directionCode, int years, string facultyName, string direction, int numberOfStudents, Department department, Department additionalDepartment)
    {
        Name = name;
        Years = years;
        FacultyName = facultyName;

        Direction = direction;
        DirectionCode = directionCode;
        
        NumberOfStudents = numberOfStudents;
        
        Department = department;
        AdditionalDepartment = additionalDepartment;
    }

    public string Name { get; }
    public int Years { get; }
    public string FacultyName { get; }
    
    public Department Department { get; }
    public Department AdditionalDepartment { get; }

    public string Direction { get; }
    public string DirectionCode { get; }
    
    public int NumberOfStudents { get; set; }

    public override string ToString()
    {
        return $"Группа: {Name}\n" +
               $"Направление: {DirectionCode}\n" +
               $"Факультет: {FacultyName}\n" +
               $"Кафедра: {Department.Name}\n" +
               $"Доп. кафедра: {AdditionalDepartment.Name}\n" +
               $"Профиль: {Direction}\n" +
               $"Количество студентов: {NumberOfStudents}";
    }
}