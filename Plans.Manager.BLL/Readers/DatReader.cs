using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Plans.Manager.BLL.Readers;

public class DatReader
{ 
    private readonly string _pathToFile;
    
    public DatReader(string pathToFile)
    {
        _pathToFile = pathToFile;
    }
    
    public XmlDocument DatToXml()
    {
        IEnumerable<string> deps = GetDatStrings(_pathToFile);
        MD5 md5 = MD5.Create();
        FileStream fileStream = File.OpenRead(_pathToFile);
        string hash = BitConverter.ToString(md5.ComputeHash(fileStream)).Replace("-", string.Empty);

        XElement document = new XElement("DepNames", new XAttribute("DateOfCreation", $"{DateTime.Now}"),
            new XAttribute("HashCode", $"{hash}"));
        int counter = 1;
        foreach (string dep in deps)
        {
            document.Add(new XElement("Department", new XAttribute("DepCode", $"{counter}"),
                new XAttribute("DepName", $"{dep}")));
            counter++;
        }

        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(document.CreateReader());
        return xmlDoc;
    }
    
    private static IEnumerable<string> GetDatStrings(string pathToFile)
    {
        int startIndex = 0;
        const int length = 100;
        List<string> deps = new List<string>();

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        StreamReader objInput = new StreamReader(
            pathToFile,
            Encoding.GetEncoding("windows-1251"));

        string objString = objInput.ReadToEnd();
        int totalLength = objString.Length;

        while (startIndex < totalLength)
        {
            string chars = objString.Substring(startIndex, length);
            string dep = chars.Trim();
            deps.Add(dep);
            startIndex += 100;
        }

        return deps;
    }
}