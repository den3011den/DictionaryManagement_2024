using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapMovementsINDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public string ErpId { get; set; }

        [Display(Name = "Время добавления")]
        [Required(ErrorMessage = "Время добавления обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Display(Name = "Время ввода документа в SAP")]
        [Required(ErrorMessage = "Время ввода документа в SAP обязательно")]
        public DateTime SapDocumentEnterTime { get; set; }

        [Display(Name = "Время проведения документа в SAP")]
        [Required(ErrorMessage = "Время проведения документа в SAP обязательно")]
        public DateTime SapDocumentPostTime { get; set; }

        [Display(Name = "Номер партии")]
        public string? BatchNo { get; set; }

        [Display(Name = "Код материала SAP")]
        [Required(ErrorMessage = "Код материала SAP обязателен")]
        public string SapMaterialCode { get; set; }

        [Display(Name = "Материал SAP")]
        public SapMaterialDTO? SapMaterialDTOFK { get; set; }

        [Display(Name = "Код завода-источника SAP")]
        [Required(ErrorMessage = "Код завода-источника SAP обязателен")]
        public string ErpPlantIdSource { get; set; }

        [Display(Name = "Код ресурса/склада-источника SAP")]
        [Required(ErrorMessage = "Код ресурса/склада-источника SAP обязателен")]
        public string ErpIdSource { get; set; }

        [Display(Name = "Источник SAP")]
        public SapEquipmentDTO? SapEquipmentSourceDTOFK { get; set; }

        [Display(Name = "Склад-источник SAP")]
        public bool? IsWarehouseSource { get; set; }

        [Display(Name = "Код завода-приемника SAP")]
        [Required(ErrorMessage = "Код завода-приемника SAP обязателен")]
        public string ErpPlantIdDest { get; set; }

        [Display(Name = "Код ресурса/склада-приёмника SAP")]
        [Required(ErrorMessage = "Код ресурса/склада-приёмника SAP обязателен")]
        public string ErpIdDest { get; set; }

        [Display(Name = "Приёмник SAP")]
        public SapEquipmentDTO? SapEquipmentDestDTOFK { get; set; }

        [Display(Name = "Склад-приёмник SAP")]
        public bool? IsWarehouseDest { get; set; }

        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [Display(Name = "Ед. изм. SAP")]
        [Required(ErrorMessage = "Ед. изм. SAP обязательно")]
        public string SapUnitOfMeasure { get; set; }

        [Display(Name = "Сторно?")]
        public bool? IsStorno { get; set; }

        [Display(Name = "Ушло в архив данных СИР?")]
        public bool? MesGone { get; set; }

        [Display(Name = "Время ухода в архив данных СИР")]
        public DateTime? MesGoneTime { get; set; }

        [Display(Name = "Ошибка")]
        public bool? MesError { get; set; }

        [Display(Name = "Сообщение об ошибке")]
        public string? MesErrorMessage { get; set; }

        [Display(Name = "ИД записи в архиве данных СИР")]
        public Guid? MesMovementId { get; set; }

        [Display(Name = "Запись в архиве данных СИР")]
        public MesMovementsDTO? MesMovementDTOFK { get; set; }

        [Display(Name = "ИД предыдущей записи")]
        public string? PreviousErpId { get; set; }

        [Display(Name = "Предыдущая запись")]
        public SapMovementsINDTO? PreviousRecordDTOFK { get; set; }

        [Display(Name = "Тип движения")]
        public string? MoveType { get; set; }

        [Display(Name = "Ед.изм. SAP")]
        public SapUnitOfMeasureDTO? SapUnitOfMeasureDTOFK { get; set; }


        [NotMapped]
        [Display(Name = "Значение")]
        public string ToStringValue
        {
            get
            {
                return Value.ToString();
            }
            set
            {
                ToStringValue = value;
            }
        }


        [NotMapped]
        [Display(Name = "ИД записи в архиве данных")]
        public string? ToStringMesMovementId
        {
            get
            {
                return MesMovementId.ToString();
            }
            set
            {
                ToStringMesMovementId = value;
            }
        }
    }
}

