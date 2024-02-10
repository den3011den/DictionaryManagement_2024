using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesParamSourceTypeDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Ид записи является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Наименование типа тэга является обязательным для заполнения полем")]
        [Display(Name = "Наименование типа параметра")]
        [MaxLength(250, ErrorMessage = "Наименование типа тэга не может быть больше 250 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; } = false;

        [NotMapped]
        public string ToStringValue
        {
            get
            {
                return Name;
            }
            set
            {
                ToStringValue = value;
            }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
