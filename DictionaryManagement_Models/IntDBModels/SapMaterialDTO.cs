using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapMaterialDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД записи является обязательным для заполнения полем")]
        public int Id { get; set; }


        [ForLogAttribute(NameProperty = "поле \"Код материала\"")]
        [Required(ErrorMessage = "Код материала SAP является обязательным для заполнения полем")]
        [Display(Name = "Код материала SAP")]
        [MaxLength(100, ErrorMessage = "Код материала SAP не может быть больше 100 символов")]
        public string Code { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование материала SAP является обязательным для заполнения полем")]
        [Display(Name = "Наименование материала SAP")]
        [MaxLength(250, ErrorMessage = "Наименование материала SAP не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Сокр. наименование\"")]
        [Required(ErrorMessage = "Сокращённое наименование материала SAP является обязательным для заполнения полем")]
        [Display(Name = "Сокращённое наименование материала SAP")]
        [MaxLength(100, ErrorMessage = "Сокращённое наименование материала SAP не может быть больше 100 символов")]
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

        [NotMapped]
        public string ToStringCodeName
        {
            get
            {
                string retVar = Code + " " + ShortName;
                return retVar;
            }
            set
            {
                ToStringCodeName = value;
            }
        }

    }
}
