using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapUnitOfMeasureDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД записи ед. изм. MES является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование ед. изм. MES является обязательным для заполнения полем")]
        [Display(Name = "Наименование материала SAP")]
        [MaxLength(250, ErrorMessage = "Наименование материала SAP не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Сокр. наименование\"")]
        [Required(ErrorMessage = "Сокращённое наименование ед. изм. является обязательным для заполнения полем")]
        [Display(Name = "Сокращённое наименование материала SAP")]
        [MaxLength(100, ErrorMessage = "Сокращённое наименование  ед. изм. не может быть больше 100 символов")]
        public string ShortName { get; set; } = string.Empty;


        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        public override string ToString()
        {
            return $"{ShortName}";
        }
    }
}
