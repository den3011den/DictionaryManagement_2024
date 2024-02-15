using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportTemplateTypeDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код типа шаблона отчёта является обязательным для заполнения полем")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Наименование типа шаблона отчёта является обязательным для заполнения полем")]
        [Display(Name = "Наименование типа шаблона отчёта")]
        [MaxLength(100, ErrorMessage = "Наименование типа шаблона отчёта не может быть больше 100 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [Display(Name = "Требуется авторасчёт")]
        public bool? NeedAutoCalc { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Доступен для настройки авторасчёта\"")]
        [Display(Name = "Доступен для настройки авторасчёта")]
        public bool? CanAutoCalc { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Файл-скрипт VBA\"")]
        [Display(Name = "Файл-скрипт VBA")]
        public string? VbaScriptFileName { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Файл-образец шаблона отчёта\"")]
        [Display(Name = "Файл-образец шаблона отчёта")]
        public string? SampleFileName { get; set; } = string.Empty;

        [NotMapped]
        [Display(Name = "Чтение")]
        public bool CanDownload { get; set; }

        [NotMapped]
        [Display(Name = "Запись")]
        public bool CanUpload { get; set; }

        [NotMapped]
        [Display(Name = "Ид записи")]
        public string ToStringId
        {
            get
            {
                return Id.ToString();
            }
            set
            {
                ToStringId = value;
            }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
