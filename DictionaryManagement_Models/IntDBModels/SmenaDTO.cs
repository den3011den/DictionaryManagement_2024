using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SmenaDTO
    {
        [ForLogAttribute(NameProperty = "поле \"Ид записи\"")]
        [Display(Name = "Ид записи")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Наименование")]
        [Required(ErrorMessage = "Наименование смены обязательно")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Ид производства обязательно")]
        [Display(Name = "Ид производства")]
        public int DepartmentId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Производство\"")]
        [Required(ErrorMessage = "Выбор производства обязателен")]
        [Display(Name = "Производство")]
        public MesDepartmentDTO DepartmentDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время начала\"")]
        [Display(Name = "Время начала смены")]
        [Required(ErrorMessage = "Время начала смены обязательно")]
        public TimeSpan StartTime { get; set; }

        [NotMapped]
        [Display(Name = "Время начала смены")]
        [Required(ErrorMessage = "Время начала смены обязательно")]
        public DateTime StartTimeDateTime { get; set; }


        [ForLogAttribute(NameProperty = "поле \"Продолжительность (в часах)\"")]
        [Display(Name = "Длительность смены (ч.)")]
        [Required(ErrorMessage = "Длительность смены обязательна")]
        [Range(1, 24, ErrorMessage = "Продолжительность смены не может быть меньше одного часа и больше 24 часов")]
        public byte HoursDuration { get; set; }

        [Display(Name = "В архиве")]
        public bool? IsArchive { get; set; }

        public override string ToString()
        {
            string ret_var = Name + " по пр-ву " + DepartmentDTOFK.ToString() + " Начало: " + StartTime.ToString() +
                " Продолжительность: " + HoursDuration.ToString();
            return ret_var;
        }
    }
}
