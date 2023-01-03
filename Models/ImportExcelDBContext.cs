using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace ReadImportSQL_ExcelFile.Models
{
    public partial class ImportExcelDBContext : DbContext
    {
        public ImportExcelDBContext()
        {
        }

        public ImportExcelDBContext(DbContextOptions<ImportExcelDBContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Contacto> Contactos { get; set; } = null!;

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
                optionsBuilder.UseSqlServer("Server=.; DataBase=ImportExcelDB;user id=sa;password=xansiety;Integrated Security=true; Encrypt=false;MultipleActiveResultSets=true;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Contacto>(entity =>
            {
                entity.HasKey(e => e.IdContacto)
                    .HasName("PK__CONTACTO__A4D6BBFA8D94D861");

                entity.ToTable("CONTACTO");

                entity.Property(e => e.Apellido)
                    .HasMaxLength(40)
                    .IsUnicode(false);

                entity.Property(e => e.Correo)
                    .HasMaxLength(40)
                    .IsUnicode(false);

                entity.Property(e => e.Nombre)
                    .HasMaxLength(40)
                    .IsUnicode(false);

                entity.Property(e => e.Telefono)
                    .HasMaxLength(40)
                    .IsUnicode(false);
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
