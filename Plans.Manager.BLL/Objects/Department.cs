namespace Plans.Manager.BLL.Objects;

public class Department
{
    public Department(string code, string department)
    {
        Code = code.Trim();
        Name = department.Trim();
    }
    
    public string Code { get; }
    public string Name { get; }

    public override string ToString()
    {
        return $"Название кафедры: {Name}, код:{Code}";
    }
}