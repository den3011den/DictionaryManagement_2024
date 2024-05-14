using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportTemplateToMesParamDTO
    {
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код обязателен для заполнения")]
        public int Id { get; set; }

        [Required(ErrorMessage = "Шаблон отчёта обязателен")]
        [Display(Name = "ИД шаблона отчёта")]
        public Guid ReportTemplateId { get; set; }

        [Display(Name = "Шаблон отчёта")]
        [Required(ErrorMessage = "Шаблон отчёта обязателен")]
        public ReportTemplateDTO? ReportTemplateDTOFK { get; set; }

        [Display(Name = "ИД тэга СИР")]
        public int? MesParamId { get; set; }

        [Display(Name = "Тэг СИР")]
        public MesParamDTO? MesParamDTOFK { get; set; }

        [Display(Name = "Код тэга СИР")]
        public string? MesParamCode { get; set; }

        [Display(Name = "Имя листа файла шаблона отчёта")]
        [Required(ErrorMessage = "Имя листа файла шаблона отчёта обязательно к заполнению")]
        public string SheetName { get; set; } = "";

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

        [NotMapped]
        private string? _toStringMesParamId;

        [NotMapped]
        [Display(Name = "ИД тэга в СИР")]
        public string? ToStringMesParamId
        {
            get
            {
                _toStringMesParamId = MesParamDTOFK == null ? "Нет" : MesParamDTOFK.Id.ToString();
                return _toStringMesParamId;
            }
            set
            {
                _toStringMesParamId = value;
            }
        }

        [NotMapped]
        private string? _toStringMesParamCode;

        [NotMapped]
        [Display(Name = "Код тэга в СИР")]
        public string? ToStringMesParamCode
        {
            get
            {
                _toStringMesParamCode = MesParamDTOFK == null ? "Нет" : MesParamDTOFK.Code;
                return _toStringMesParamCode;
            }
            set
            {
                _toStringMesParamCode = value;
            }
        }

        [NotMapped]
        private string? _toStringMesParamSourceType;

        [NotMapped]
        [Display(Name = "Источник тэга в СИР")]
        public string? ToStringMesParamSourceType
        {
            get
            {
                _toStringMesParamSourceType = MesParamDTOFK == null ? "" : MesParamDTOFK.MesParamSourceTypeDTOFK == null ? "" : MesParamDTOFK.MesParamSourceTypeDTOFK.Name;
                return _toStringMesParamSourceType;
            }
            set
            {
                _toStringMesParamSourceType = value;
            }
        }

        [NotMapped]
        private string? _toStringSapEquipmentSource;

        [NotMapped]
        [Display(Name = "Источник SAP тэга в СИР")]
        public string? ToStringSapEquipmentSource
        {
            get
            {
                _toStringSapEquipmentSource = MesParamDTOFK == null ? "" : MesParamDTOFK.SapEquipmentSourceDTOFK == null ? "" : MesParamDTOFK.SapEquipmentSourceDTOFK.ToStringErpPlantIdErpIdName;
                return _toStringSapEquipmentSource;
            }
            set
            {
                _toStringSapEquipmentSource = value;
            }
        }

        [NotMapped]
        private string? _toStringSapEquipmentDest;
        [NotMapped]
        [Display(Name = "Приёмник SAP тэга в СИР")]
        public string? ToStringSapEquipmentDest
        {
            get
            {
                _toStringSapEquipmentDest = MesParamDTOFK == null ? "" : MesParamDTOFK.SapEquipmentDestDTOFK == null ? "" : MesParamDTOFK.SapEquipmentDestDTOFK.ToStringErpPlantIdErpIdName;
                return _toStringSapEquipmentDest;
            }
            set
            {
                _toStringSapEquipmentDest = value;
            }
        }

        [NotMapped]
        private string? _toStringSapMaterial;
        [NotMapped]
        [Display(Name = "Материал SAP тэга в СИР")]
        public string? ToStringSapMaterial
        {
            get
            {
                _toStringSapMaterial = MesParamDTOFK == null ? "" : MesParamDTOFK.SapMaterialDTOFK == null ? "" : MesParamDTOFK.SapMaterialDTOFK.ToStringCodeName;
                return _toStringSapMaterial;
            }
            set
            {
                _toStringSapMaterial = value;
            }
        }

        [NotMapped]
        private string? _toStringSapUnitOfMeasure;
        [NotMapped]
        [Display(Name = "Ед.изм. SAP тэга в СИР")]
        public string? ToStringSapUnitOfMeasure
        {
            get
            {
                _toStringSapUnitOfMeasure = MesParamDTOFK == null ? "" : MesParamDTOFK.SapUnitOfMeasureDTOFK == null ? "" : MesParamDTOFK.SapUnitOfMeasureDTOFK.NameAndShortName;
                return _toStringSapUnitOfMeasure;
            }
            set
            {
                _toStringSapUnitOfMeasure = value;
            }
        }

        [NotMapped]
        private bool _needWriteToSapBool;
        [NotMapped]
        [Display(Name = "Передавать в SAP")]
        public bool NeedWriteToSapBool
        {
            get
            {
                _needWriteToSapBool = MesParamDTOFK == null ? false : MesParamDTOFK.NeedWriteToSapBool;
                return _needWriteToSapBool;
            }
            set
            {
                _needWriteToSapBool = value;
            }
        }

        [NotMapped]
        private bool _needReadFromSapBool;
        [NotMapped]
        [Display(Name = "Читать из SAP")]
        public bool NeedReadFromSapBool
        {
            get
            {
                _needReadFromSapBool = MesParamDTOFK == null ? false : MesParamDTOFK.NeedReadFromSapBool;
                return _needWriteToSapBool;
            }
            set
            {
                _needReadFromSapBool = value;
            }
        }

        [NotMapped]
        private bool _needReadFromMesBool;
        [NotMapped]
        [Display(Name = "Читать из Mes")]
        public bool NeedReadFromMesBool
        {
            get
            {
                _needReadFromMesBool = MesParamDTOFK == null ? false : MesParamDTOFK.NeedReadFromMesBool;
                return _needReadFromMesBool;
            }
            set
            {
                _needReadFromMesBool = value;
            }
        }

        [NotMapped]
        private bool _needWriteToMesBool;
        [NotMapped]
        [Display(Name = "Передавать в Mes")]
        public bool NeedWriteToMesBool
        {
            get
            {
                _needWriteToMesBool = MesParamDTOFK == null ? false : MesParamDTOFK.NeedWriteToMesBool;
                return _needWriteToMesBool;
            }
            set
            {
                _needWriteToMesBool = value;
            }
        }

        [NotMapped]
        private string _toStringReportTemplateId;
        [NotMapped]
        [Display(Name = "Ид шаблона отчёта")]
        public string ToStringReportTemplateId
        {
            get
            {
                _toStringReportTemplateId = ReportTemplateId.ToString();
                return _toStringReportTemplateId;
            }
            set
            {
                _toStringReportTemplateId = value;
            }
        }

        [NotMapped]
        private bool _isArchiveBool;
        [NotMapped]
        [Display(Name = "Архив")]
        public bool IsArchiveBool
        {
            get
            {
                _isArchiveBool = MesParamDTOFK == null ? false : MesParamDTOFK.IsArchive;
                return _isArchiveBool;
            }
            set
            {
                _isArchiveBool = value;
            }
        }


        public override string ToString()
        {
            return "Шаблон: \"{ReportTemplateDTOFK.Id.ToString}\" " + "Тэг: \"" + MesParamDTOFK != null ? MesParamDTOFK.Id.ToString() : ""
                + "\" Код тэга: \"" + MesParamCode != null ? MesParamCode : "" + "Лист: \"" + SheetName + "\"";
        }

    }

}
