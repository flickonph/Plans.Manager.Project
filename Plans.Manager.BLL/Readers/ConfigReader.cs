using System.Xml;

namespace Plans.Manager.BLL.Readers;

public static class ConfigReader
{ 
    public static List<EducationalGroup> GetGroups()
    {
        var xmlElements = GetDocument(Directory.Groups).DocumentElement!.ChildNodes.OfType<XmlElement>();

        return xmlElements.Select(xmlElement =>
                new EducationalGroup(
                    xmlElement.GetAttribute("Name"),
                    xmlElement.GetAttribute("Code"),
                     Convert.ToInt32(xmlElement.GetAttribute("Years")),
                    xmlElement.GetAttribute("FacultyName"),
                    xmlElement.GetAttribute("DepCode"),
                    xmlElement.GetAttribute("AddDepCode"),
                    xmlElement.GetAttribute("Direction"),
                    Convert.ToInt32(xmlElement.GetAttribute("NumberOfStudents"))))
            .ToList();
    }

    public static List<EducationalGroup> GetGroups(int year)
    {
        var xmlElements = (List<XmlElement>)GetDocument(Directory.Groups).DocumentElement!.ChildNodes.OfType<XmlElement>()
            .Where(x => x.GetAttribute("Years") == year.ToString());
        return xmlElements.Select(xmlElement =>
                new EducationalGroup(
                    xmlElement.GetAttribute("Name"),
                    xmlElement.GetAttribute("Code"),
                    Convert.ToInt32(xmlElement.GetAttribute("Years")),
                    xmlElement.GetAttribute("FacultyName"),
                    xmlElement.GetAttribute("DepCode"),
                    xmlElement.GetAttribute("AddDepCode"),
                    xmlElement.GetAttribute("Direction"),
                    Convert.ToInt32(xmlElement.GetAttribute("NumberOfStudents"))))
            .ToList();
    }

    public static List<string> GetAutoCompleteDisciplines()
    {
        var xmlElements = GetDocument(Directory.AutoComplete).DocumentElement!.ChildNodes.OfType<XmlElement>();
        return xmlElements.Select(xmlElement => xmlElement.GetAttribute("Name")).ToList();
    }

    public static Dictionary<string, string> GetDepartments()
    {
        var xmlElements = GetDocument(Directory.Departments).DocumentElement!.ChildNodes.OfType<XmlElement>();
        return xmlElements.ToDictionary(element => element.GetAttribute("DepCode"),
            element => element.GetAttribute("DepName"));
    }

    public static XmlDocument GetDocument(Directory directory)
    {
        var currentXmlDocument = new XmlDocument();
        var requiredDocPath = GetDocumentPath(directory);
        currentXmlDocument.Load(requiredDocPath);
        
        return currentXmlDocument;
    }

    public static string GetDocumentPath(Directory directory)
    {
        const string settingsFilePath = @"Config/settings.xml";
        const string xmlElementName = "Paths";
        
        var currentXmlDocument = new XmlDocument();
        currentXmlDocument.Load(settingsFilePath);
        var xmlNodeList = currentXmlDocument.GetElementsByTagName(xmlElementName);
        var xmlElements = xmlNodeList.OfType<XmlElement>().First();
        
        var requiredDocPath = xmlElements.GetAttribute(directory.Value);
        return requiredDocPath;
    }
}

public class EducationalGroup
{
    public EducationalGroup(string name, string code, int years, string facultyName, string depCode,
        string addDepCode, string direction, int numberOfStudents)
    {
        Name = name;
        Code = code;
        Years = years;
        FacultyName = facultyName;
        DepCode = depCode;
        AddDepCode = addDepCode;
        Direction = direction;
        NumberOfStudents = numberOfStudents;
    }

    public string Name { get; set; }
    public string Code { get; }
    public int Years { get; }
    public string FacultyName { get; }
    public string DepCode { get; }
    public string AddDepCode { get; }
    public string Direction { get; }
    public int NumberOfStudents { get; set; }

    public override string ToString()
    {
        return $"Название: {Name}\n" +
               $"Направление: {Code}\n" +
               $"Факультет: {FacultyName}\n" +
               $"Кафедра: {GetDepartmentName(DepCode)}\n" +
               $"Доп. кафедра: {GetDepartmentName(AddDepCode)}\n" +
               $"Профиль: {Direction}\n" +
               $"Количество студентов: {NumberOfStudents}";
    }

    public string AllData()
    {
        return $"Направление: {Code}\n" +
               $"Факультет: {FacultyName}\n" +
               $"Кафедра: {GetDepartmentName(DepCode)}\n" +
               $"Доп. кафедра: {GetDepartmentName(AddDepCode)}\n" +
               $"Профиль: {Direction}\n" +
               $"Количество студентов: {NumberOfStudents}";
    }

    private static string GetDepartmentName(string depCode)
    {
        return !string.IsNullOrEmpty(depCode) ? ConfigReader.GetDepartments()[depCode] : "";
    }
}

public class Directory
{
    private Directory(string value) { Value = value; }
    public string Value { get; private set; }
    public static Directory AutoComplete => new("AutoComplete");
    public static Directory Departments => new("Departments");
    public static Directory Groups => new("Groups");
    public static Directory ConnectionString => new("ConnectionString");

    public override string ToString()
    {
        return Value;
    }
}