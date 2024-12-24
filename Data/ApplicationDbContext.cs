using Expense_Tracker.Models;
using Microsoft.EntityFrameworkCore;


namespace Expense_Tracker.Data
{
    public class ApplicationDbContext : DbContext
    {

        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }

        public DbSet<User> Users { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Map 'User' entity to 'tbl_User' table
            modelBuilder.Entity<User>().ToTable("tbl_User");
        }
    }
}
