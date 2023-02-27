using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Plans.Manager.BLL.Services;

public class DatHandler
{
    private readonly string _pathToFile;

    /// <summary>
    ///     Initializes a new instance of the <see cref="DatHandler" /> class.
    /// </summary>
    /// <param name="pathToFile"> Local path to file to work with. </param>
    /// <example>
    ///     var configHandler = new ConfigHandler(string pathToFile);
    ///     ... list = configHandler.GetRequiredAttributes(option);
    /// </example>
    public DatHandler(string pathToFile)
    {
        _pathToFile = pathToFile;
    }

    /// <summary>
    ///     Converts collection of departments to XML document.
    /// </summary>
    /// <returns> XML document with names of departments with numbers. </returns>
    /// <exception cref="NullReferenceException"> Thrown if document is null before being returned. </exception>
    public XmlDocument DatToXml()
    {
        Check(_pathToFile);
        var deps = GetDatStrings(_pathToFile);
        var md5 = MD5.Create();
        var fileStream = File.OpenRead(_pathToFile);
        var hash = BitConverter.ToString(md5.ComputeHash(fileStream)).Replace("-", string.Empty);

        var document = new XElement("DepNames", new XAttribute("DateOfCreation", $"{DateTime.Now}"),
            new XAttribute("HashCode", $"{hash}"));
        var counter = 1;
        foreach (var dep in deps)
        {
            document.Add(new XElement("Department", new XAttribute("DepCode", $"{counter}"),
                new XAttribute("DepName", $"{dep}")));
            counter++;
        }

        var xmlDoc = new XmlDocument();
        xmlDoc.Load(document.CreateReader());
        return xmlDoc;
    }

    /// <summary>
    ///     Gets departments names from "DepNames.dat" file.
    /// </summary>
    /// <param name="pathToFile"> Path to configuration file. </param>
    /// <returns> Collection of departments. </returns>
    private static IEnumerable<string> GetDatStrings(string pathToFile)
    {
        var startIndex = 0;
        const int length = 100;
        var deps = new List<string>();

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var objInput = new StreamReader(
            pathToFile,
            Encoding.GetEncoding("windows-1251"));

        var objString = objInput.ReadToEnd();
        var totalLength = objString.Length;

        while (startIndex < totalLength)
        {
            var chars = objString.Substring(startIndex, length);
            var dep = chars.Trim();
            deps.Add(dep);
            startIndex += 100;
        }

        return deps;
    }

    private static void Check(params string?[] parameters)
    {
        if (parameters.Any(p => string.IsNullOrEmpty(p) && string.IsNullOrWhiteSpace(p)))
            throw new NullReferenceException();
    }
}