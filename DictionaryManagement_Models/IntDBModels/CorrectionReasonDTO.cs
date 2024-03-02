using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class CorrectionReasonDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД записи причины корректировки является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Наименование причины корректировки является обязательным для заполнения полем")]
        [Display(Name = "Наименование причины корректировки")]
        [MaxLength(250, ErrorMessage = "Наименование причины корректировки не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        public override string ToString()
        {
            return $"{Name}";
        }
    }
}
