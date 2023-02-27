using System.Xml;

namespace Plans.Manager.BLL.Readers;

public abstract class AReader
{
    private readonly string _path;
    private readonly string? _extension;
    private readonly XmlDocument _currentDocument;

    protected AReader(string path)
    {
        _extension = "plx";
        _path = path;
        _currentDocument = new XmlDocument();
        LoadDocument();
    }

    protected List<XmlElement> GetXmlElements(string xmlElementName)
    {
        var xmlNodeList = _currentDocument.GetElementsByTagName(xmlElementName);
        var xmlElements = xmlNodeList.OfType<XmlElement>().ToList();
        return xmlElements;
    }

    private void LoadDocument()
    {
        CheckFile();
        _currentDocument.Load(_path);
    }

    private void CheckFile()
    {
        if (_extension is null or "") throw new Exception("Extension for check method is not set");

        if (!File.Exists(_path)) throw new FileNotFoundException("The file does not exist at the current path.");

        var fileInfo = new FileInfo(_path);

        if (fileInfo.Extension.ToLower() != $".{_extension}")
            throw new InvalidDataException("The file extension is not correct.");
    }
}