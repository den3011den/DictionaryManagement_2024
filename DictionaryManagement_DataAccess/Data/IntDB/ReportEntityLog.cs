using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportEntityLog", Schema = "dbo")]
    public class ReportEntityLog
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        [Required]
        public DateTime LogTime { get; set; }

        [Required]
        public Guid ReportEntityId { get; set; }
        [ForeignKey("ReportEntityId")]
        public ReportEntity ReportEntityFK { get; set; }

        public string LogMessage { get; set; }

        public string LogType { get; set; }

        public bool IsError { get; set; } = false;

    }
}
