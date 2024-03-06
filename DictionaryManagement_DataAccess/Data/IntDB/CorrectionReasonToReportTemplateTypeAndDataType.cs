using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("CorrectionReasonToReportTemplateTypeAndDataType", Schema = "dbo")]
    public class CorrectionReasonToReportTemplateTypeAndDataType
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int Id { get; set; }

        public int CorrectionReasonId { get; set; }

        [ForeignKey("CorrectionReasonId")]
        public CorrectionReason? CorrectionReasonFK { get; set; }

        public int ReportTemplateTypeId { get; set; }

        [ForeignKey("ReportTemplateTypeId")]
        public ReportTemplateType? ReportTemplateTypeFK { get; set; }

        public int DataTypeId { get; set; }

        [ForeignKey("DataTypeId")]
        public DataType? DataTypeFK { get; set; }

    }

}
