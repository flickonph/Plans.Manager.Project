using Microsoft.EntityFrameworkCore;
using Plans_Manager_Shared.Tables;

namespace Plans_Manager_DAL.Context;

public sealed class Context : DbContext
{
#pragma warning disable CS8618
    public Context(DbContextOptions<Context> options) : base(options)
#pragma warning restore CS8618
    {
        /*Database.EnsureDeleted();*/
        Database.EnsureCreated();
    }

    public DbSet<TablePrint> TablePrints { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<TablePrint>().HasKey(x => x.Id);
        modelBuilder.Entity<TablePrint>().Property(x => x.Date)
            .HasConversion(
                x => x.ToString(),
                x => DateOnly.Parse(x)
            );
        modelBuilder.Entity<TablePrint>().HasIndex(x => x.Hashcode).IsUnique();
        base.OnModelCreating(modelBuilder);
    }
}