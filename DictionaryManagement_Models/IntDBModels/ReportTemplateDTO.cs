using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportTemplateDTO
    {

        [ForLogAttribute(NameProperty = "поле \"Ид шаблона отчёта\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время добавления\"")]
        [Required(ErrorMessage = "Время добавления обязательно")]
        [Display(Name = "Время добавления")]
        public DateTime AddTime { get; set; }

        //[Required(ErrorMessage = "ИД добавившего пользователя обязательно")]        
        [Display(Name = "Ид добавившего пользователя")]
        public Guid AddUserId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Добавил\"")]
        //[Required(ErrorMessage = "Выбор пользователя обязателен")]
        [Display(Name = "Пользователь")]
        public UserDTO AddUserDTOFK { get; set; }

        [Display(Name = "Описание")]
        public string Description { get; set; } = "Шаблон типа: \"\" c вых данными: \"\" для производства: \"\"";

        [Required(ErrorMessage = "Ид типа отчёта обязательно")]
        [Display(Name = "Ид типа отчёта")]
        public int ReportTemplateTypeId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тип отчёта\"")]
        [Required(ErrorMessage = "Выбор типа отчёта обязателен")]
        [Display(Name = "Тип отчёта")]
        public ReportTemplateTypeDTO ReportTemplateTypeDTOFK { get; set; }

        [Required(ErrorMessage = "Ид типа данных обязательно")]
        [Display(Name = "Ид типа данных")]
        public int DestDataTypeId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тип выходных данных\"")]
        [Required(ErrorMessage = "Выбор типа данных обязателен")]
        [Display(Name = "Тип данных")]
        public DataTypeDTO DestDataTypeDTOFK { get; set; }

        [Required(ErrorMessage = "Ид производства обязательно")]
        [Display(Name = "Ид производства")]
        public int DepartmentId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Производство\"")]
        [Required(ErrorMessage = "Выбор производства обязателен")]
        [Display(Name = "Производство")]
        public MesDepartmentDTO MesDepartmentDTOFK { get; set; }

        [Display(Name = "Имя файла")]
        [Required(ErrorMessage = "Выбор файла обязателен")]
        public string TemplateFileName { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Авторасчёт\"")]
        [Display(Name = "Авторасчёт")]
        public bool? NeedAutoCalc { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Порядок авторасчёта\"")]
        [Display(Name = "Порядок авторасчёта")]
        [Range(1, 1000000, ErrorMessage = "Порядок авторасчёта должен быть числом от {1} до {2}")]
        public int? AutoCalcOrder { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Кол-во прогонов авторасчёта\"")]
        [Display(Name = "Кол-во прогонов авторасчёта")]
        [Range(1, 5, ErrorMessage = "Кол-во прогонов авторасчёта должно быть от {1} до {2}")]
        public int? AutoCalcNumber { get; set; }

        [Display(Name = "История файлов шаблона отчёта")]
        public IEnumerable<ReportTemplateFileHistoryDTO>? ReportTemplateFileHistoryListDTO { get; set; }

        [Display(Name = "Тэги шаблона отчёта")]
        public IEnumerable<ReportTemplateToMesParamDTO>? ReportTemplateToMesParamListDTO { get; set; }

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
            return $"{Id.ToString().ToUpper()} {Description}";
        }
    }
}

