using System.Xml;
using Plans.Manager.BLL.Objects;

namespace Plans.Manager.BLL.Readers;

public static class XmlReader
{
    private static XmlDocument GetDocument(DirectoryEnum directoryEnum)
    {
        XmlDocument currentXmlDocument = new XmlDocument();
        currentXmlDocument.Load(GetDocumentPath(directoryEnum)); // TODO: UserUnhandledException
        
        return currentXmlDocument;
    }

    public static string GetDepartmentName(string departmentCode)
    {
        IEnumerable<XmlElement> xmlElements = GetDocument(DirectoryEnum.Departments).DocumentElement!.ChildNodes.OfType<XmlElement>();
        return xmlElements.Where(element => element.GetAttribute("DepCode") == departmentCode).FirstOrDefault()?.GetAttribute("DepName") ?? string.Empty; // TODO: Rename "DepCode" & "DepName"
    }
    
    public static List<Group> GetGroups()
    {
        IEnumerable<XmlElement> xmlElements = GetDocument(DirectoryEnum.Groups).DocumentElement!.ChildNodes.OfType<XmlElement>();

        return xmlElements.Select(element =>
                new Group(
                    element.GetAttribute("Name"),
                    element.GetAttribute("Code"),
                     Convert.ToInt32(element.GetAttribute("Years")),
                    element.GetAttribute("FacultyName"),
                    element.GetAttribute("Direction"),
                    Convert.ToInt32(element.GetAttribute("NumberOfStudents")),
                    new Department(element.GetAttribute("DepCode"), GetDepartmentName(element.GetAttribute("DepCode"))), // TODO: Rename "DepCode" & "DepName"
                    new Department(element.GetAttribute("AddDepCode"), GetDepartmentName(element.GetAttribute("DepCode"))))).ToList(); // TODO: Rename "DepCode" & "DepName"
    }

    public static List<Group> GetGroups(int year)
    {
        IEnumerable<XmlElement> xmlElements = GetDocument(DirectoryEnum.Groups).DocumentElement!.ChildNodes.OfType<XmlElement>()
            .Where(element => element.GetAttribute("Years") == year.ToString());
        
        return xmlElements.Select(element =>
            new Group(
                element.GetAttribute("Name"),
                element.GetAttribute("Code"),
                Convert.ToInt32(element.GetAttribute("Years")),
                element.GetAttribute("FacultyName"),
                element.GetAttribute("Direction"),
                Convert.ToInt32(element.GetAttribute("NumberOfStudents")),
                new Department(element.GetAttribute("DepCode"), GetDepartmentName(element.GetAttribute("DepCode"))), // TODO: Rename "DepCode" & "DepName"
                new Department(element.GetAttribute("AddDepCode"), GetDepartmentName(element.GetAttribute("DepCode"))))).ToList(); // TODO: Rename "DepCode" & "DepName"
    }

    public static List<string> GetAutoCompleteDisciplines()
    {
        IEnumerable<XmlElement> xmlElements = GetDocument(DirectoryEnum.AutoComplete).DocumentElement!.ChildNodes.OfType<XmlElement>();
        return xmlElements.Select(xmlElement => xmlElement.GetAttribute("Name")).ToList();
    }

    public static List<Department> GetDepartments()
    {
        IEnumerable<XmlElement> xmlElements = GetDocument(DirectoryEnum.Departments).DocumentElement!.ChildNodes.OfType<XmlElement>();
        return xmlElements.Select(element => new Department(element.GetAttribute("DepCode"), GetDepartmentName(element.GetAttribute("DepCode")))).ToList(); // TODO: Rename "DepCode" & "DepName"
    }

    public static string GetDocumentPath(DirectoryEnum directoryEnum)
    {
        const string settingsFilePath = @"Config/Settings.xml"; // TODO: Remove this
        const string xmlElementName = "Paths"; // TODO: Remove this
        
        XmlDocument currentXmlDocument = new XmlDocument();
        currentXmlDocument.Load(settingsFilePath);
        XmlNodeList xmlNodeList = currentXmlDocument.GetElementsByTagName(xmlElementName);
        XmlElement xmlElements = xmlNodeList.OfType<XmlElement>().First();
        
        string requiredDocPath = xmlElements.GetAttribute(directoryEnum.Value);
        return requiredDocPath;
    }
}

public class DirectoryEnum
{
    private DirectoryEnum(string value) { Value = value; }
    public string Value { get; }
    public static DirectoryEnum AutoComplete => new("AutoComplete");
    public static DirectoryEnum Departments => new("Departments");
    public static DirectoryEnum Groups => new("Groups");
    public static DirectoryEnum ConnectionString => new("ConnectionString");

    public override string ToString()
    {
        return Value;
    }
}