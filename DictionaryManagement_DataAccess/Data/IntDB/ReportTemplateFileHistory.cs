using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportTemplateFileHistory", Schema = "dbo")]
    public class ReportTemplateFileHistory
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        public Guid ReportTemplateId { get; set; }
        [ForeignKey("ReportTemplateId")]
        public ReportTemplate ReportTemplateFK { get; set; }

        [Required]
        public DateTime AddTime { get; set; }

        public Guid AddUserId { get; set; }
        [ForeignKey("AddUserId")]
        public User AddUserFK { get; set; }

        public string? PreviousFileName { get; set; }

        [Required]
        public string CurrentFileName { get; set; }

    }
}
