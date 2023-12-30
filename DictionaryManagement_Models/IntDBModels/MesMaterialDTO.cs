using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesMaterialDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД записи является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код материала\"")]
        [Required(ErrorMessage = "Код материала MES является обязательным для заполнения полем")]
        [Display(Name = "Код материала MES")]
        [MaxLength(100, ErrorMessage = "Код материала MES не может быть больше 100 символов")]
        public string Code { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование материала MES является обязательным для заполнения полем")]
        [Display(Name = "Наименование материала MES")]
        [MaxLength(250, ErrorMessage = "Наименование материала MES не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Сокр. наименование\"")]
        [Required(ErrorMessage = "Сокращённое наименование материала MES является обязательным для заполнения полем")]
        [Display(Name = "Сокращённое наименование материала MES")]
        [MaxLength(100, ErrorMessage = "Сокращённое наименование материала MES не может быть больше 100 символов")]
        public string ShortName { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [NotMapped]
        public string ToStringValue { get; set; } = string.Empty;

        public override string ToString()
        {
            ToStringValue = $"{Code} {ShortName}";
            return ToStringValue;
        }
        //public string ToStringForLog()
        //{
        //    return $"{Code} {ShortName}";
        //}
    }
}
