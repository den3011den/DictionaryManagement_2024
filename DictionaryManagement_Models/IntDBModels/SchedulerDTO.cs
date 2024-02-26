using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace DictionaryManagement_Models.IntDBModels
{
    public class SchedulerDTO
    {
        [ForLogAttribute(NameProperty = "поле \"Ид группы\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД записи обязательно")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование модуля\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Наименование модуля")]
        [Required(ErrorMessage = "Наименование модуля обязательно")]
        public string ModuleName { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время начала\"")]
        [Display(Name = "Время старта задания")]
        [Required(ErrorMessage = "Время старта задания обязательно")]
        public TimeSpan StartTime { get; set; }

        [NotMapped]
        [Display(Name = "Время старта задания")]
        [Required(ErrorMessage = "Время старта задания обязательно")]
        public DateTime StartTimeDateTime { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время последнего выполнения\"")]
        [Display(Name = "Время последнего выполнения")]
        public DateTime? LastExecuted { get; set; } = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

        [ForLogAttribute(NameProperty = "поле \"Сейчас выполняется\"")]
        [Display(Name = "Сейчас выполняется")]
        public bool IsRunningNow { get; set; } = false;

        public override string ToString()
        {
            return $"Модуль: {ModuleName} Старт: {StartTime.ToString()}";
        }
    }
}
