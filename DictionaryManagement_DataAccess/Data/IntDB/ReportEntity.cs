using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("ReportEntity", Schema = "dbo")]
    public class ReportEntity
    {

        [Key]
        [Required]
        public Guid Id { get; set; }

        [Required]
        public Guid ReportTemplateId { get; set; }
        [ForeignKey("ReportTemplateId")]
        public ReportTemplate ReportTemplateFK { get; set; }

        public DateTime? ReportTimeStart { get; set; }

        public DateTime? ReportTimeEnd { get; set; }

        public int? ReportDepartmentId { get; set; }
        [ForeignKey("ReportDepartmentId")]
        public MesDepartment? ReportDepartmentFK { get; set; }

        public DateTime? DownloadTime { get; set; }

        public Guid DownloadUserId { get; set; }
        [ForeignKey("DownloadUserId")]
        public User DownloadUserFK { get; set; }

        public string? DownloadReportFileName { get; set; }

        public bool? DownloadSuccessFlag { get; set; }

        public DateTime? UploadTime { get; set; }

        public Guid? UploadUserId { get; set; }
        [ForeignKey("UploadUserId")]
        public User? UploadUserFK { get; set; }

        public string? UploadReportFileName { get; set; }

        public bool? UploadSuccessFlag { get; set; }
    }
}
