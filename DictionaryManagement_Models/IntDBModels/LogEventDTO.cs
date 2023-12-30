using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class LogEventDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Int64 Id { get; set; }


        [Display(Name = "Ид типа события")]
        [Required(ErrorMessage = "ИД типа события обязателен")]
        public int LogEventTypeId { get; set; }

        [Display(Name = "Тип события")]
        [Required(ErrorMessage = "Тип события обязателен")]
        public LogEventTypeDTO LogEventTypeDTOFK { get; set; }

        [Required(ErrorMessage = "Старое значение")]
        public string? OldValue { get; set; }
        [Required(ErrorMessage = "Новое значение")]
        public string? NewValue { get; set; }

        [Required]
        public DateTime EventTime { get; set; } = DateTime.Now;

        [Display(Name = "ИД пользователя")]
        [Required(ErrorMessage = "ИД пользователя обязателен")]
        public Guid UserId { get; set; }

        [Display(Name = "Пользователь")]
        [Required(ErrorMessage = "Пользователь обязателен")]
        public UserDTO UserDTOFK { get; set; }

        [Display(Name = "Описание")]
        [Required(ErrorMessage = "Описание обязателено")]
        public string Description { get; set; }

        [Display(Name = "Крит")]
        public bool? IsCritical { get; set; } = false;

        [Display(Name = "Ошибка")]
        public bool? IsError { get; set; } = false;

        [Display(Name = "Предупреждение")]
        public bool? IsWarning { get; set; } = false;

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
        [Display(Name = "Крит.")]
        public bool IsCriticalBool
        {
            get
            {

                return IsCritical == null ? false : (bool)IsCritical;

            }
            set
            {
                IsCriticalBool = value;
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

        [NotMapped]
        [Display(Name = "Ошибка")]
        public bool IsWarningBool
        {
            get
            {
                return IsWarning == null ? false : (bool)IsWarning;
            }
            set
            {
                IsWarningBool = value;
            }
        }

    }
}

