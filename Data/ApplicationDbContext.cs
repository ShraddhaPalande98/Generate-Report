using DataTables.Models;
using Microsoft.EntityFrameworkCore;

namespace DataTables.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }

        public DbSet<Sells> Sells { get; set; }
    }
}
