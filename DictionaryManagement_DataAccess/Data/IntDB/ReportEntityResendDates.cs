using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportEntityResendDates", Schema = "dbo")]
    public class ReportEntityResendDates
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        [Required]
        public Guid ReportEntityId { get; set; }

        [ForeignKey("ReportEntityId")]
        public virtual ReportEntity ReportEntityFK { get; set; }

        [Required]
        public DateTime ResendDate { get; set; }

    }
}
