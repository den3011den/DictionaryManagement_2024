using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("MesMovements", Schema = "dbo")]
    public class MesMovements
    {

        [Key]
        [Required]
        public Guid Id { get; set; }

        [Required]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Required]
        public Guid AddUserId { get; set; }
        [ForeignKey("AddUserId")]
        public User AddUserFK { get; set; }

        [Required]
        public int MesParamId { get; set; }
        [ForeignKey("MesParamId")]
        public MesParam MesParamFK { get; set; }

        [Required]
        public DateTime ValueTime { get; set; }

        [Required]
        public decimal Value { get; set; } = decimal.Zero;

        public Guid? SapMovementOutId { get; set; }
        [ForeignKey("SapMovementOutId")]
        public SapMovementsOUT? SapMovementsOUTFK { get; set; }

        public string? SapMovementInId { get; set; }
        [ForeignKey("SapMovementInId")]
        public SapMovementsIN? SapMovementsINFK { get; set; }

        public int? DataSourceId { get; set; }
        [ForeignKey("DataSourceId")]
        public DataSource? DataSourceFK { get; set; }

        public int? DataTypeId { get; set; }
        [ForeignKey("DataTypeId")]
        public DataType? DataTypeFK { get; set; }

        public Guid? ReportGuid { get; set; }
        [ForeignKey("ReportGuid")]
        public ReportEntity? ReportEntityFK { get; set; }

        public Guid? PreviousRecordId { get; set; }
        [ForeignKey("PreviousRecordId")]
        public MesMovements? MesMovementsFK { get; set; }

        public bool? MesGone { get; set; }

        public DateTime? MesGoneTime { get; set; }

        public virtual ICollection<MesMovementsComment>? MesMovementsCommentList { get; set; }

        public bool? NeedWriteToSap { get; set; }

    }
}
