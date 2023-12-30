using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    public class IntDBApplicationDbContext : DbContext
    {
        public IntDBApplicationDbContext(DbContextOptions<IntDBApplicationDbContext> options) : base(options)
        {

        }

        public DbSet<SapEquipment> SapEquipment { get; set; }
        public DbSet<SapUnitOfMeasure> SapUnitOfMeasure { get; set; }
        public DbSet<MesUnitOfMeasure> MesUnitOfMeasure { get; set; }
        public DbSet<MesMaterial> MesMaterial { get; set; }
        public DbSet<SapMaterial> SapMaterial { get; set; }
        public DbSet<CorrectionReason> CorrectionReason { get; set; }

        public DbSet<MesParamSourceType> MesParamSourceType { get; set; }
        public DbSet<DataType> DataType { get; set; }
        public DbSet<DataSource> DataSource { get; set; }
        public DbSet<ReportTemplateType> ReportTemplateType { get; set; }
        public DbSet<LogEventType> LogEventType { get; set; }
        public DbSet<Settings> Settings { get; set; }

        public DbSet<UnitOfMeasureSapToMesMapping> UnitOfMeasureSapToMesMapping { get; set; }
        public DbSet<SapToMesMaterialMapping> SapToMesMaterialMapping { get; set; }
        public DbSet<MesDepartment> MesDepartment { get; set; }
        public DbSet<MesParam> MesParam { get; set; }

        public DbSet<User> User { get; set; }
        public DbSet<Role> Role { get; set; }
        public DbSet<ADGroup> ADGroup { get; set; }
        public DbSet<UserToRole> UserToRole { get; set; }
        public DbSet<RoleToADGroup> RoleToADGroup { get; set; }
        public DbSet<RoleToDepartment> RoleToDepartment { get; set; }

        public DbSet<ReportTemplate> ReportTemplate { get; set; }
        public DbSet<ReportTemplateTypeTоRole> ReportTemplateTypeTоRole { get; set; }
        public DbSet<ReportEntity> ReportEntity { get; set; }
        public DbSet<ReportEntityLog> ReportEntityLog { get; set; }
        public DbSet<Smena> Smena { get; set; }
        public DbSet<Version> Version { get; set; }
        public DbSet<Scheduler> Scheduler { get; set; }
        public DbSet<MesNdoStocks> MesNdoStocks { get; set; }
        public DbSet<SapNdoOUT> SapNdoOUT { get; set; }

        public DbSet<MesMovements> MesMovements { get; set; }
        public DbSet<SapMovementsIN> SapMovementsIN { get; set; }
        public DbSet<SapMovementsOUT> SapMovementsOUT { get; set; }

        public DbSet<LogEvent> LogEvent { get; set; }

        public DbSet<MesMovementsComment> MesMovementsComment { get; set; }





        //protected override void OnModelCreating(ModelBuilder modelBuilder)
        //{
        //    base.OnModelCreating(modelBuilder);
        //    //modelBuilder.Entity<SapEquipment>()
        //    //    .HasAlternateKey(x => new { x.ErpPlantId, x.ErpId })
        //    //    .HasName("SapEquipmentDestFK")
        //    //    .HasName("SapEquipmentSourceFK");
        //    modelBuilder.Entity<SapMaterial>()
        //        .HasIndex(x => x.Code);
        //        //.HasPrincipalKey(x => x.Code);
        //        //.HasAlternateKey(x => new { x.Code })                
        //        //.HasName("FK_SapMovementsOUT_SapMaterial");
        //}

    }

}
