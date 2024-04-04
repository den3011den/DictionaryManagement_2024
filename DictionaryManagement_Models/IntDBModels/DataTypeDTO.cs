using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class DataTypeDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "ИД записи")]
        [Required(ErrorMessage = "ИД записи вида данных является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Наименование вида данных является обязательным для заполнения полем")]
        [Display(Name = "Наименование вида данных")]
        [MaxLength(250, ErrorMessage = "Наименование вида данных не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Приоритет передачи в SAP\"")]
        [Display(Name = "Приоритет передачи в SAP")]
        public int? Priority { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Результирующий тип авторасчёта\"")]
        [Display(Name = "Результирующий тип авторасчёта")]
        public bool? IsAutoCalcDestDataType { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Нельзя переименовывать\"")]
        [Display(Name = "Нельзя переименовывать")]
        public bool? CantChangeName { get; set; }

        [NotMapped]
        public bool? OldCantChangeName { get; set; }

        [NotMapped]
        public string ToStringValue { get; set; } = string.Empty;

        public override string ToString()
        {
            ToStringValue = $"{Id} {Name}";
            return ToStringValue;
        }
    }
}
