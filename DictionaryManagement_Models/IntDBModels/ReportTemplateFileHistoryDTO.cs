using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportTemplateFileHistoryDTO
    {


        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Int64 Id { get; set; }

        [Required(ErrorMessage = "ИД шаблона отчёта обязателен")]
        [Display(Name = "ИД шаблона отчёта")]
        public Guid ReportTemplateId { get; set; }

        [Required(ErrorMessage = "Шаблон отчёта обязателен")]
        [Display(Name = "Шаблон отчёта")]
        public ReportTemplateDTO ReportTemplateDTOFK { get; set; }

        [Required(ErrorMessage = "Время добавления записи обязателено")]
        [Display(Name = "Время добавления")]
        public DateTime AddTime { get; set; }

        [Required(ErrorMessage = "ИД пользователя обязателен")]
        [Display(Name = "ИД пользователя")]
        public Guid AddUserId { get; set; }

        [Required(ErrorMessage = "Пользователь обязателен")]
        [Display(Name = "Пользователь")]
        public UserDTO AddUserDTOFK { get; set; }

        [Display(Name = "Имя предшествующего файла")]
        public string? PreviousFileName { get; set; }

        [Required(ErrorMessage = "Имя текущего файла обязательно")]
        [Display(Name = "Имя текущего файла")]
        public string CurrentFileName { get; set; }


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
