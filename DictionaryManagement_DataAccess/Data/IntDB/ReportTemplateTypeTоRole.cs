using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportTemplateTypeTоRole", Schema = "dbo")]
    public class ReportTemplateTypeTоRole
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int Id { get; set; }

        [Required]
        public int ReportTemplateTypeId { get; set; }

        [ForeignKey("ReportTemplateTypeId")]
        public ReportTemplateType ReportTemplateTypeFK { get; set; }

        [Required]
        public Guid RoleId { get; set; }

        [ForeignKey("RoleId")]
        public Role RoleFK { get; set; }

        public bool CanDownload { get; set; } = true;
        public bool CanUpload { get; set; } = true;


    }

}
