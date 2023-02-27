using System.Security.Cryptography;
using System.Xml;
using Microsoft.EntityFrameworkCore;
using Plans_Manager_DAL.Context;
using Plans_Manager_Shared.Tables;
using Plans.Manager.BLL.Readers;
using Directory = Plans.Manager.BLL.Readers.Directory;

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
        var unused = await Task.Factory.StartNew(Worker).ConfigureAwait(true);
    }
    
    private bool Worker()
    {
        foreach (var path in _paths)
            try
            {
                var dbServExt = new DatabaseHandler(ConfigReader.GetDocumentPath(Directory.ConnectionString));
                var options = new DbContextOptionsBuilder<Context>();
                options.UseNpgsql(
                    $"Host={dbServExt.GetConnectionStringParams("Host")};" +
                    $"Username={dbServExt.GetConnectionStringParams("Username")};" +
                    $"Password={dbServExt.GetConnectionStringParams("Password")};" +
                    $"Database={dbServExt.GetConnectionStringParams("Database")};");

                var context = new Context(options.Options);
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
        var doc = new XmlDocument();
        doc.Load(_pathToConfigFile);
        var root = doc.DocumentElement;

        if (root == null || !root.HasAttribute($"{option}"))
            throw new ArgumentException($"{option} is not provided by configuration file");

        var attribute = root.GetAttribute($"{option}");
        return attribute;
    }
    
    public byte[] GetFileBytes(string pathToFile)
    {
        var byteArr = File.ReadAllBytes(pathToFile);
        return byteArr;
    }
    
    public string GetHash(string pathToFile)
    {
        var md5 = MD5.Create();
        var hash = BitConverter.ToString(md5.ComputeHash(GetFileStream(pathToFile))).Replace("-", string.Empty);
        return hash;
    }
    
    public DateOnly GetCurrentDate()
    {
        var now = DateOnly.FromDateTime(DateTime.Now);
        return now;
    }
    
    private FileStream GetFileStream(string pathToFile)
    {
        var stream = File.OpenRead(pathToFile);
        return stream;
    }
}