using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

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

        [Display(Name = "Дата СИР")]
        [Required(ErrorMessage = "Дата обязательна")]
        public DateTime ResendDateSIR { get; set; }

        [NotMapped]
        [Display(Name = "Пользовательская дата")]
        [Required(ErrorMessage = "Дата обязательна")]
        public DateTime ResendDateUser
        {
            get
            {
                return ResendDateSIR.AddDays(-1);
            }
            set
            {
                ResendDateUser = value;
            }
        }
    }
}

