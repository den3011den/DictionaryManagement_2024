using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportTemplateType", Schema = "dbo")]
    public class ReportTemplateType
    {
        [Key]
        [Required]
        public int Id { get; set; }

        [Required]
        [MaxLength(1000)]
        [MinLength(1)]
        public string Name { get; set; } = string.Empty;

        public bool? NeedAutoCalc { get; set; }

        public bool IsArchive { get; set; }

        public bool? CanAutoCalc { get; set; }
        public string? VbaScriptFileName { get; set; } = string.Empty;
        public string? SampleFileName { get; set; } = string.Empty;
    }
}
