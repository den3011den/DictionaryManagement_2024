using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class VersionDTO
    {

        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        [Display(Name = "Версия БД")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Версия БД обязательна")]
        [MaxLength(20, ErrorMessage = "Значение не должно быть больше 20-ти символов")]
        public string version { get; set; }
    }
}

