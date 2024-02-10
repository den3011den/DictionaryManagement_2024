using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesParamDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД тэга СИР\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код тэга СИР\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Код обязателен для заполнения")]
        [StringLength(100, MinimumLength = 1, ErrorMessage = "Код может быть от 1 до 100 символов")]
        [Display(Name = "Код тэга СИР")]
        public string Code { get; set; }

        //[Required(ErrorMessage = "Наименование тэга СИР обязательно для заполнения")]
        //[StringLength(250, MinimumLength = 3, ErrorMessage = "Наименование может быть от 3 до 250 символов")]
        [ForLogAttribute(NameProperty = "поле \"Наименование тэга СИР\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Наименование")]
        public string? Name { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Описание\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [StringLength(100, MinimumLength = 3, ErrorMessage = "Описание может быть от 3 до 100 символов")]
        [Display(Name = "Описание")]
        public string? Description { get; set; }

        [Display(Name = "Ид типа источника параметра")]
        public int? MesParamSourceType { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Источник\"")]
        [Display(Name = "Источник")]
        public MesParamSourceTypeDTO? MesParamSourceTypeDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тэг/ИД источника\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Тэг источника")]
        [StringLength(300, ErrorMessage = "Тэг источника может быть до 300 символов")]
        public string? MesParamSourceLink { get; set; } = "";

        [Display(Name = "Ид производства")]
        public int? DepartmentId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Производство\"")]
        [Display(Name = "Производство")]
        public MesDepartmentDTO? MesDepartmentDTOFK { get; set; }

        [Display(Name = "Ид источника SAP")]
        public int? SapEquipmentIdSource { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Источник SAP\"")]
        [Display(Name = "Источник SAP")]
        public SapEquipmentDTO? SapEquipmentSourceDTOFK { get; set; }

        [Display(Name = "Ид приёмника SAP")]
        public int? SapEquipmentIdDest { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Приёмник SAP\"")]
        [Display(Name = "Приёмник SAP")]
        public SapEquipmentDTO? SapEquipmentDestDTOFK { get; set; }

        [Display(Name = "Ид материала MES")]
        public int? MesMaterialId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Материал MES\"")]
        [Display(Name = "Материал MES")]
        public MesMaterialDTO? MesMaterialDTOFK { get; set; }

        [Display(Name = "Ид материала SAP")]
        public int? SapMaterialId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Материал SAP\"")]
        [Display(Name = "Материал SAP")]
        public SapMaterialDTO? SapMaterialDTOFK { get; set; }

        [Display(Name = "Ид ед.изм. MES")]
        public int? MesUnitOfMeasureId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Ед.изм. MES\"")]
        [Display(Name = "Ед.изм. MES")]
        public MesUnitOfMeasureDTO? MesUnitOfMeasureDTOFK { get; set; }

        [Display(Name = "Ид ед.изм. SAP")]
        public int? SapUnitOfMeasureId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Ед.изм. SAP\"")]
        [Display(Name = "Ед.изм. SAP")]
        public SapUnitOfMeasureDTO? SapUnitOfMeasureDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Глубина опроса (в днях)\"")]
        [Display(Name = "Глубина опроса (в днях)")]
        public int? DaysRequestInPast { get; set; } = 45;

        [ForLogAttribute(NameProperty = "поле \"Точка измерения\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Точка измерения")]
        [StringLength(maximumLength: 100, ErrorMessage = "Точка измерения не может быть длиннее 100 символов")]
        public string? TI { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование точки измерения\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Наименование точки измерения")]
        [StringLength(maximumLength: 250, ErrorMessage = "Наименование точки измерения не может быть длиннее 250 символов")]
        public string? NameTI { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Технологическое место\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Технологическое место")]
        [StringLength(maximumLength: 100, ErrorMessage = "Технологическое место не может быть длиннее 100 символов")]
        public string? TM { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование технологического места\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Наименование технологического места")]
        [StringLength(maximumLength: 250, ErrorMessage = "Наименование технологического места не может быть длиннее 250 символов")]
        public string? NameTM { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Коэффициент пересчёта данных по тэгу из ед. изм. MES в ед. изм. СИР\"")]
        [Display(Name = "Коэффициент пересчёта данных по тэгу из ед. изм. MES в ед. изм. СИР")]
        public decimal? MesToSirUnitOfMeasureKoef { get; set; } = decimal.One;

        [ForLogAttribute(NameProperty = "поле \"Передавать в SAP\"")]
        [Display(Name = "Передавать в SAP")]
        public bool? NeedWriteToSap { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Читать из SAP\"")]
        [Display(Name = "Читать из SAP")]
        public bool? NeedReadFromSap { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Читать из MES\"")]
        [Display(Name = "Читать из MES")]
        public bool? NeedReadFromMes { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Передавать в MES\"")]
        [Display(Name = "Передавать в MES")]
        public bool? NeedWriteToMes { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Параметр НДО\"")]
        [Display(Name = "Параметр НДО")]
        public bool? IsNdo { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [NotMapped]
        [Display(Name = "Тэг СИР")]
        public string ToStringCodeName
        {
            get
            {
                return $"{Code} {Name}";
            }
            set
            {
                ToStringCodeName = value;
            }
        }

        [NotMapped]
        [Display(Name = "НДО")]
        public bool IsNdoBool
        {
            get
            {
                return IsNdo == null ? false : (bool)IsNdo;
            }
            set
            {
                IsNdoBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Передавать в SAP")]
        public bool NeedWriteToSapBool
        {
            get
            {
                return NeedWriteToSap == null ? false : (bool)NeedWriteToSap;
            }
            set
            {
                NeedWriteToSapBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Читать из SAP")]
        public bool NeedReadFromSapBool
        {
            get
            {
                return NeedReadFromSap == null ? false : (bool)NeedReadFromSap;
            }
            set
            {
                NeedReadFromSapBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Читать из MES")]
        public bool NeedReadFromMesBool
        {
            get
            {
                return NeedReadFromMes == null ? false : (bool)NeedReadFromMes;
            }
            set
            {
                NeedReadFromMesBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Передавать в MES")]
        public bool NeedWriteToMesBool
        {
            get
            {
                return NeedWriteToMes == null ? false : (bool)NeedWriteToMes;
            }
            set
            {
                NeedWriteToMesBool = value;
            }
        }


        [NotMapped]
        [Display(Name = "ИД")]
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
            return $"{Code} {Name}";
        }
    }
}

