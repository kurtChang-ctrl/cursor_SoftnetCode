using System;
using System.Collections.Generic;
using Base.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace SoftNetWebII.Tables
{
    public partial class MyContext : DbContext
    {
        public MyContext()
        {
        }

        public MyContext(DbContextOptions<MyContext> options)
            : base(options)
        {
        }

        public virtual DbSet<User> User { get; set; }
        public virtual DbSet<XpCode> XpCode { get; set; }
        public virtual DbSet<XpFlow> XpFlow { get; set; }
        public virtual DbSet<XpFlowLine> XpFlowLine { get; set; }
        public virtual DbSet<XpFlowNode> XpFlowNode { get; set; }
        public virtual DbSet<XpFlowSign> XpFlowSign { get; set; }
  
        public virtual DbSet<XpProg> XpProg { get; set; }



        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer(_Fun.Config.Db);
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<User>(entity =>
            {
                entity.Property(e => e.Id)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.Account)
                    .IsRequired()
                    .HasMaxLength(20)
                    .IsUnicode(false);

                entity.Property(e => e.DeptId)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasMaxLength(20);

                entity.Property(e => e.PhotoFile).HasMaxLength(100);

                entity.Property(e => e.Pwd)
                    .IsRequired()
                    .HasMaxLength(32)
                    .IsUnicode(false);
            });

            modelBuilder.Entity<XpCode>(entity =>
            {
                entity.HasKey(e => new { e.Type, e.Value });

                entity.Property(e => e.Type)
                    .HasMaxLength(20)
                    .IsUnicode(false);

                entity.Property(e => e.Value)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.Ext)
                    .HasMaxLength(30)
                    .IsUnicode(false);

                entity.Property(e => e.Name_enUS).HasMaxLength(30);

                entity.Property(e => e.Name_zhCN).HasMaxLength(30);

                entity.Property(e => e.Name_zhTW)
                    .IsRequired()
                    .HasMaxLength(30);

                entity.Property(e => e.Note).HasMaxLength(255);
            });

            modelBuilder.Entity<XpFlow>(entity =>
            {
                entity.Property(e => e.Id)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.Code)
                    .IsRequired()
                    .HasMaxLength(20)
                    .IsUnicode(false);

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasMaxLength(30);
            });

            modelBuilder.Entity<XpFlowLine>(entity =>
            {
                entity.Property(e => e.Id)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.CondStr)
                    .HasMaxLength(255)
                    .IsUnicode(false);

                entity.Property(e => e.EndNode)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.FlowId)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.StartNode)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);
            });

            modelBuilder.Entity<XpFlowNode>(entity =>
            {
                entity.Property(e => e.Id)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.FlowId)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasMaxLength(30);

                entity.Property(e => e.NodeType)
                    .IsRequired()
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .IsFixedLength();

                entity.Property(e => e.PassType)
                    .IsRequired()
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .IsFixedLength();

                entity.Property(e => e.SignerType)
                    .HasMaxLength(2)
                    .IsUnicode(false);

                entity.Property(e => e.SignerValue)
                    .HasMaxLength(30)
                    .IsUnicode(false);
            });

            modelBuilder.Entity<XpFlowSign>(entity =>
            {
                entity.Property(e => e.Id)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.FlowId)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.NodeName)
                    .IsRequired()
                    .HasMaxLength(30);

                entity.Property(e => e.Note).HasMaxLength(255);

                entity.Property(e => e.SignStatus)
                    .IsRequired()
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .IsFixedLength();

                entity.Property(e => e.SignTime).HasColumnType("datetime");

                entity.Property(e => e.SignerId)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.SignerName)
                    .IsRequired()
                    .HasMaxLength(20);

                entity.Property(e => e.SourceId)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsUnicode(false);
            });

            modelBuilder.Entity<XpProg>(entity =>
            {
                entity.Property(e => e.Id)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.Code)
                    .IsRequired()
                    .HasMaxLength(30)
                    .IsUnicode(false);

                entity.Property(e => e.Icon)
                    .HasMaxLength(20)
                    .IsUnicode(false);

                entity.Property(e => e.Name)
                    .IsRequired()
                    .HasMaxLength(30);

                entity.Property(e => e.Url)
                    .HasMaxLength(100)
                    .IsUnicode(false);
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }

    public partial class User
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Account { get; set; }
        public string Pwd { get; set; }
        public string DeptId { get; set; }
        public string PhotoFile { get; set; }
        public bool Status { get; set; }
    }
    public partial class XpFlow
    {
        public string Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public bool Portrait { get; set; }
        public bool Status { get; set; }
    }
    public partial class XpFlowLine
    {
        public string Id { get; set; }
        public string FlowId { get; set; }
        public string StartNode { get; set; }
        public string EndNode { get; set; }
        public string CondStr { get; set; }
        public short Sort { get; set; }
    }
    public partial class XpFlowNode
    {
        public string Id { get; set; }
        public string FlowId { get; set; }
        public string Name { get; set; }
        public string NodeType { get; set; }
        public short PosX { get; set; }
        public short PosY { get; set; }
        public string SignerType { get; set; }
        public string SignerValue { get; set; }
        public string PassType { get; set; }
        public short? PassNum { get; set; }
    }
    public partial class XpFlowSign
    {
        public string Id { get; set; }
        public string FlowId { get; set; }
        public string SourceId { get; set; }
        public string NodeName { get; set; }
        public short FlowLevel { get; set; }
        public short TotalLevel { get; set; }
        public string SignerId { get; set; }
        public string SignerName { get; set; }
        public string SignStatus { get; set; }
        public DateTime? SignTime { get; set; }
        public string Note { get; set; }
    }
}
