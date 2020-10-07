using System;
using System.Data.Entity;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;

namespace MCQFeedImport
{
  
    public partial class Blancco : DbContext
    {
        public Blancco()
            : base("name=Blancco")
        {
        }

        public virtual DbSet<MCQFeed> MCQFeed { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<MCQFeed>()
                .Property(e => e.Swap_Assure_Eligible)
                .IsFixedLength();

            modelBuilder.Entity<MCQFeed>()
                .Property(e => e.Make)
                .IsUnicode(false);

            modelBuilder.Entity<MCQFeed>()
                .Property(e => e.Model)
                .IsUnicode(false);

            modelBuilder.Entity<MCQFeed>()
                .Property(e => e.IMEI)
                .IsUnicode(false);

            modelBuilder.Entity<MCQFeed>()
                .Property(e => e.Customer_Contract_Start_Date)
                .IsUnicode(false);
        }
    }
}
