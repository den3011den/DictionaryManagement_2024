using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MaterialDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД записи является обязательным для заполнения полем")]
        public int Id { get; set; }


        [ForLogAttribute(NameProperty = "поле \"Код материала\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Код материала является обязательным для заполнения полем")]
        [Display(Name = "Код материала")]
        [MaxLength(100, ErrorMessage = "Код материала не может быть больше 100 символов")]
        public string Code { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Наименование материала является обязательным для заполнения полем")]
        [Display(Name = "Наименование материала SAP")]
        [MaxLength(250, ErrorMessage = "Наименование материала не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Сокр. наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Сокращённое наименование материала является обязательным для заполнения полем")]
        [Display(Name = "Сокращённое наименование материала")]
        [MaxLength(100, ErrorMessage = "Сокращённое наименование материала не может быть больше 100 символов")]
        public string ShortName { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

    }
}
