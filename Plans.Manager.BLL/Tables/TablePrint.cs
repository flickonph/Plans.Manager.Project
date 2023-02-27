namespace Plans_Manager_Shared.Tables;

public class TablePrint
{
    public TablePrint(string hashcode, DateOnly date, string name, byte[] file)
    {
        Id = Guid.Empty;
        Hashcode = hashcode;
        Date = date;
        Name = name;
        File = file;
    }

#pragma warning disable CS8618
    private TablePrint()
    {
    }
#pragma warning restore CS8618

    public string Hashcode { get; }
    public DateOnly Date { get; }
    public string Name { get; }
    public byte[] File { get; }
    public Guid Id { get; set; }
}