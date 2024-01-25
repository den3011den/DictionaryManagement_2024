using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportEntityLogDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Int64 Id { get; set; }

        [Display(Name = "Дата/время")]
        [Required(ErrorMessage = "Дата/время обязательны для заполнения")]
        public DateTime LogTime { get; set; }

        [Display(Name = "ИД экземпляра отчёта")]
        [Required(ErrorMessage = "ИД экземпляра отчёта обязателен")]
        public Guid ReportEntityId { get; set; }

        [Display(Name = "Экземпляр отчёта")]
        [Required(ErrorMessage = "Экземпляр отчёта обязателен")]
        public ReportEntityDTO ReportEntityDTOFK { get; set; }

        [Display(Name = "Сообщение")]
        [Required(ErrorMessage = "Сообщение обязательно к заполнению")]
        public string LogMessage { get; set; }

        [Display(Name = "Тип сообщения")]
        [Required(ErrorMessage = "Тип сообщения обязателен к заполнению")]
        public string LogType { get; set; }

        [Display(Name = "Флаг ошибки")]
        [Required(ErrorMessage = "Флаг ошибки обязателен к заполнению")]
        public bool IsError { get; set; } = false;

        [NotMapped]
        [Display(Name = "Тэг СИР")]
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

        [NotMapped]
        [Display(Name = "Ошибка")]
        public bool IsErrorBool
        {
            get
            {

                return IsError == null ? false : (bool)IsError;

            }
            set
            {
                IsErrorBool = value;
            }
        }
    }
}
