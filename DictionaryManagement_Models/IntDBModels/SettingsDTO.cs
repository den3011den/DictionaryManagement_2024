using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SettingsDTO
    {
        [ForLogAttribute(NameProperty = "поле \"Ид записи\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код настройки является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование настройки является обязательным для заполнения полем")]
        [Display(Name = "Наименование настройки")]
        [MaxLength(100, ErrorMessage = "Наименование настройки не может быть больше 100 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Описание\"")]
        [Display(Name = "Описание настройки")]
        [MaxLength(300, ErrorMessage = "Описание настройки не может быть больше 300 символов")]
        public string Description { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Значение настройки\"")]
        [Display(Name = "Значение настройки")]
        public string Value { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"{Name}";
        }
    }
}
