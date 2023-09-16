using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace BGC_Test_WebAPI.Models;

public partial class BgcTestDbContext : DbContext
{
    public BgcTestDbContext()
    {
    }

    public BgcTestDbContext(DbContextOptions<BgcTestDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Order> Orders { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Order>(entity =>
        {
            entity.HasKey(e => e.Id);

            entity.Property(e => e.Id).HasColumnName("ID");
            entity.Property(e => e.Category)
                .HasMaxLength(255)
                .IsUnicode(false);
            entity.Property(e => e.City)
                .HasMaxLength(255)
                .IsUnicode(false);
            entity.Property(e => e.OrderDate).HasColumnType("date");
            entity.Property(e => e.Product)
                .HasMaxLength(255)
                .IsUnicode(false);
            entity.Property(e => e.Region)
                .HasMaxLength(255)
                .IsUnicode(false);
            entity.Property(e => e.TotalPrice).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.UnitPrice).HasColumnType("decimal(10, 2)");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
