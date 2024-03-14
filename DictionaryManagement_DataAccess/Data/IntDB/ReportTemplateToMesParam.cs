using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportTemplateToMesParam", Schema = "dbo")]
    public class ReportTemplateToMesParam
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int Id { get; set; }

        public Guid ReportTemplateId { get; set; }

        [ForeignKey("ReportTemplateId")]
        public ReportTemplate? ReportTemplateFK { get; set; }

        public int? MesParamId { get; set; }

        [ForeignKey("MesParamId")]
        public MesParam? MesParamFK { get; set; }

        public string? MesParamCode { get; set; }

        public string SheetName { get; set; } = "";

    }

}
