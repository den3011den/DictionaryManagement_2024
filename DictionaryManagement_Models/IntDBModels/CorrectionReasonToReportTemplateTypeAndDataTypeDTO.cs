using DictionaryManagement_Models.IntDBModels;
using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    public class CorrectionReasonToReportTemplateTypeAndDataTypeDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код обязателен для заполнения")]
        public int Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"ИД причины корректировки\"")]
        [Required(ErrorMessage = "Причина корректировки обязательна для заполнения")]
        [Display(Name = "ИД причины корректировки")]
        public int CorrectionReasonId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Причина корректировки\"")]
        [Display(Name = "Причина корректировки")]
        [Required(ErrorMessage = "Причина корректировки обязательна для заполнения")]
        public CorrectionReasonDTO? CorrectionReasonDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"ИД типа шаблона отчёта\"")]
        [Display(Name = "ИД типа шаблона отчёта")]
        [Required(ErrorMessage = "Тип шаблона отчёта обязателен для заполнения")]
        public int ReportTemplateTypeId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тип шаблона отчёта\"")]
        [Required(ErrorMessage = "Тип данных обязателен для заполнения")]
        [Display(Name = "Тип шаблона отчёта")]
        public ReportTemplateTypeDTO? ReportTemplateTypeDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"ИД типа данных\"")]
        [Display(Name = "ИД типа данных")]
        [Required(ErrorMessage = "Тип данных обязателен для заполнения")]
        public int DataTypeId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тип данных\"")]
        [Display(Name = "Тип данных")]
        [Required(ErrorMessage = "Тип данных обязателен для заполнения")]
        public DataTypeDTO? DataTypeDTOFK { get; set; }

    }

}
