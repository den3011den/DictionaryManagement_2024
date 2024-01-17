using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportEntityDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [Required(ErrorMessage = "ИД шаблона отчёта обязательно")]
        [Display(Name = "ИД шаблона отчёта")]
        public Guid ReportTemplateId { get; set; }

        [Required(ErrorMessage = "Шаблон отчёта обязателен")]
        [Display(Name = "Шаблон отчёта")]
        public ReportTemplateDTO ReportTemplateDTOFK { get; set; }

        [Display(Name = "Начало периода")]
        public DateTime? ReportTimeStart { get; set; }

        [Display(Name = "Окончание периода")]
        public DateTime? ReportTimeEnd { get; set; }

        [Display(Name = "ИД производства")]
        public int? ReportDepartmentId { get; set; }

        [Display(Name = "Производство")]
        public MesDepartmentDTO? ReportDepartmentDTOFK { get; set; }

        [Display(Name = "Когда скачали")]
        public DateTime? DownloadTime { get; set; }

        [Required(ErrorMessage = "ИД скачавшего пользователя обязателен")]
        [Display(Name = "ИД скачавшего пользователя")]
        public Guid DownloadUserId { get; set; }

        [Required(ErrorMessage = "Скачавший пользователь обязателен")]
        [Display(Name = "Кто скачал")]
        public UserDTO DownloadUserDTOFK { get; set; }

        [Display(Name = "Имя скачаного файла")]
        public string? DownloadReportFileName { get; set; }

        [Display(Name = "Успешно скачан")]
        public bool DownloadSuccessFlag { get; set; } = true;

        [Display(Name = "Когда загружен")]
        public DateTime? UploadTime { get; set; }

        [Display(Name = "ИД кто загрузил")]
        public Guid? UploadUserId { get; set; }

        [Display(Name = "Кто загрузил")]
        public UserDTO? UploadUserDTOFK { get; set; }

        [Display(Name = "Имя загруженного файла")]
        public string? UploadReportFileName { get; set; }

        [Display(Name = "Успешно загружен")]
        public bool UploadSuccessFlag { get; set; } = false;

        public bool ResendMode { get; set; } = false;

        [Display(Name = "Даты переотправки данных в SAP")]
        public IEnumerable<ReportEntityResendDatesDTO>? ReportEntityResendDatesListDTO { get; set; }

        [NotMapped]
        [Display(Name = "Ид записи")]
        public string ToStringId
        {
            get
            {
                return Id.ToString();
            }
            set
            {
                ToStringId = value;
            }
        }

    }
}
