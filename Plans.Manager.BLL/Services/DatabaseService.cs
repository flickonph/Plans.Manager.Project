using System.Security.Cryptography;
using System.Xml;
using Microsoft.EntityFrameworkCore;
using Plans_Manager_DAL.Context;
using Plans_Manager_Shared.Tables;
using Plans.Manager.BLL.Readers;
using XmlReader = Plans.Manager.BLL.Readers.XmlReader;

namespace Plans.Manager.BLL.Services;

public class DatabaseService
{
    private readonly List<string> _paths;

    public DatabaseService(List<string> paths)
    {
        _paths = paths;
    }
    
    public async void LoadToDatabase()
    {
        bool unused = await Task.Factory.StartNew(Worker).ConfigureAwait(true);
    }
    
    private bool Worker()
    {
        foreach (string path in _paths)
            try
            {
                DatabaseHandler dbServExt = new DatabaseHandler(XmlReader.GetDocumentPath(DirectoryEnum.ConnectionString));
                DbContextOptionsBuilder<Context> options = new DbContextOptionsBuilder<Context>();
                options.UseNpgsql(
                    $"Host={dbServExt.GetConnectionStringParams("Host")};" +
                    $"Username={dbServExt.GetConnectionStringParams("Username")};" +
                    $"Password={dbServExt.GetConnectionStringParams("Password")};" +
                    $"Database={dbServExt.GetConnectionStringParams("Database")};");

                Context context = new Context(options.Options);
                context.TablePrints.Add(new TablePrint(
                    dbServExt.GetHash(path),
                    dbServExt.GetCurrentDate(),
                    Path.GetFileNameWithoutExtension(path),
                    dbServExt.GetFileBytes(path)));
                context.SaveChanges();
            }
            catch
            {
                // TODO: Logger
            }

        return true;
    }
}

public class DatabaseHandler
{
    private readonly string _pathToConfigFile;

    public DatabaseHandler(string pathToConfigFile)
    {
        _pathToConfigFile = pathToConfigFile;
    }
    
    public string GetConnectionStringParams(string option)
    {
        XmlDocument doc = new XmlDocument();
        doc.Load(_pathToConfigFile);
        XmlElement? root = doc.DocumentElement;

        if (root == null || !root.HasAttribute($"{option}"))
            throw new ArgumentException($"{option} is not provided by configuration file");

        string attribute = root.GetAttribute($"{option}");
        return attribute;
    }
    
    public byte[] GetFileBytes(string pathToFile)
    {
        byte[] byteArr = File.ReadAllBytes(pathToFile);
        return byteArr;
    }
    
    public string GetHash(string pathToFile)
    {
        MD5 md5 = MD5.Create();
        string hash = BitConverter.ToString(md5.ComputeHash(GetFileStream(pathToFile))).Replace("-", string.Empty);
        return hash;
    }
    
    public DateOnly GetCurrentDate()
    {
        DateOnly now = DateOnly.FromDateTime(DateTime.Now);
        return now;
    }
    
    private FileStream GetFileStream(string pathToFile)
    {
        FileStream stream = File.OpenRead(pathToFile);
        return stream;
    }
}