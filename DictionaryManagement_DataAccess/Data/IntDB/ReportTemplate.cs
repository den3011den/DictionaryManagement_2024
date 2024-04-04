using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportTemplate", Schema = "dbo")]
    public class ReportTemplate
    {

        [Key]
        [Required]
        public Guid Id { get; set; }

        [Required]
        public DateTime AddTime { get; set; }

        public string? Description { get; set; }

        [Required]
        public Guid AddUserId { get; set; }

        [ForeignKey("AddUserId")]
        public User AddUserFK { get; set; }

        [Required]
        public int ReportTemplateTypeId { get; set; }

        [ForeignKey("ReportTemplateTypeId")]
        public ReportTemplateType ReportTemplateTypeFK { get; set; }



        [Required]
        public int DestDataTypeId { get; set; }

        [ForeignKey("DestDataTypeId")]
        public DataType DestDataTypeFK { get; set; }

        [Required]
        public int DepartmentId { get; set; }

        [ForeignKey("DepartmentId")]
        public MesDepartment MesDepartmentFK { get; set; }

        [Required]
        public string TemplateFileName { get; set; }

        public bool IsArchive { get; set; }
        public bool? NeedAutoCalc { get; set; }
        public int? AutoCalcOrder { get; set; }
        public int? AutoCalcNumber { get; set; }

        public virtual ICollection<ReportTemplateFileHistory>? ReportTemplateFileHistoryList { get; set; }
        public virtual ICollection<ReportTemplateToMesParam>? ReportTemplateToMesParamList { get; set; }
    }

}
