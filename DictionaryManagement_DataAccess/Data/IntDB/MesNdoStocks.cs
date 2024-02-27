using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("MesNdoStocks", Schema = "dbo")]
    public class MesNdoStocks
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        [Required]
        public int MesParamId { get; set; }
        [ForeignKey("MesParamId")]
        public MesParam MesParamFK { get; set; }

        [Required]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Required]
        public Guid AddUserId { get; set; }
        [ForeignKey("AddUserId")]
        public User AddUserFK { get; set; }

        [Required]
        public DateTime ValueTime { get; set; }

        [Required]
        [Precision(18, 6)]
        public decimal Value { get; set; } = decimal.Zero;

        [Precision(18, 6)]
        public decimal? ValueDifference { get; set; } = decimal.Zero;

        public Guid? ReportGuid { get; set; }
        [ForeignKey("ReportGuid")]
        public ReportEntity? ReportEntityFK { get; set; }

        public Int64? SapNdoOutId { get; set; }
        [ForeignKey("SapNdoOutId")]
        public SapNdoOUT? SapNdoOUTFK { get; set; }

    }
}
