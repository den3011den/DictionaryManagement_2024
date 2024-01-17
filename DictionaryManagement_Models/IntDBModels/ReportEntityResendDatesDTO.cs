using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportEntityResendDatesDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Int64 Id { get; set; }


        [Display(Name = "Ид экземпляра отчёта")]
        [Required(ErrorMessage = "Ид экземпляра отчёта обязателен")]
        public Guid ReportEntityId { get; set; }

        [Display(Name = "Экземпляр отчёта")]
        [Required(ErrorMessage = "Экземпляр отчёта обязателен")]
        public ReportEntityDTO ReportEntityDTOFK { get; set; }

        [Display(Name = "Дата")]
        [Required(ErrorMessage = "Дата обязательна")]
        public DateTime ResendDate { get; set; }

    }
}

