using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapMovementsINDTO
    {

        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public string ErpId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время добавления\"")]
        [Display(Name = "Время добавления")]
        [Required(ErrorMessage = "Время добавления обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [ForLogAttribute(NameProperty = "поле \"Время ввода документа в SAP\"")]
        [Display(Name = "Время ввода документа в SAP")]
        [Required(ErrorMessage = "Время ввода документа в SAP обязательно")]
        public DateTime SapDocumentEnterTime { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время проведения документа в SAP\"")]
        [Display(Name = "Время проведения документа в SAP")]
        [Required(ErrorMessage = "Время проведения документа в SAP обязательно")]
        public DateTime SapDocumentPostTime { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Номер партии\"")]
        [Display(Name = "Номер партии")]
        public string? BatchNo { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код материала SAP\"")]
        [Display(Name = "Код материала SAP")]
        [Required(ErrorMessage = "Код материала SAP обязателен")]
        public string SapMaterialCode { get; set; }

        [Display(Name = "Материал SAP")]
        public SapMaterialDTO? SapMaterialDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код завода-источника SAP\"")]
        [Display(Name = "Код завода-источника SAP")]
        [Required(ErrorMessage = "Код завода-источника SAP обязателен")]
        public string ErpPlantIdSource { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код ресурса/склада-источника SAP\"")]
        [Display(Name = "Код ресурса/склада-источника SAP")]
        [Required(ErrorMessage = "Код ресурса/склада-источника SAP обязателен")]
        public string ErpIdSource { get; set; }

        [Display(Name = "Источник SAP")]
        public SapEquipmentDTO? SapEquipmentSourceDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Склад-источник SAP\"")]
        [Display(Name = "Склад-источник SAP")]
        public bool? IsWarehouseSource { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код завода-приемника SAP\"")]
        [Display(Name = "Код завода-приемника SAP")]
        [Required(ErrorMessage = "Код завода-приемника SAP обязателен")]
        public string ErpPlantIdDest { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Код ресурса/склада-приёмника SAP\"")]
        [Display(Name = "Код ресурса/склада-приёмника SAP")]
        [Required(ErrorMessage = "Код ресурса/склада-приёмника SAP обязателен")]
        public string ErpIdDest { get; set; }

        [Display(Name = "Приёмник SAP")]
        public SapEquipmentDTO? SapEquipmentDestDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Склад-приёмник SAP\"")]
        [Display(Name = "Склад-приёмник SAP")]
        public bool? IsWarehouseDest { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Значение\"")]
        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [ForLogAttribute(NameProperty = "поле \"Ед. изм. SAP\"")]
        [Display(Name = "Ед. изм. SAP")]
        [Required(ErrorMessage = "Ед. изм. SAP обязательно")]
        public string SapUnitOfMeasure { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Сторно\"")]
        [Display(Name = "Сторно?")]
        public bool? IsStorno { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Mes забрал\"")]
        [Display(Name = "Ушло в архив данных СИР?")]
        public bool? MesGone { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время Mes забрал\"")]
        [Display(Name = "Время ухода в архив данных СИР")]
        public DateTime? MesGoneTime { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Mes ошибка\"")]
        [Display(Name = "Ошибка")]
        public bool? MesError { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Mes ошибка сообщение\"")]
        [Display(Name = "Сообщение об ошибке")]
        public string? MesErrorMessage { get; set; }

        [ForLogAttribute(NameProperty = "поле \"ИД записи в архиве данных СИР\"")]
        [Display(Name = "ИД записи в архиве данных СИР")]
        public Guid? MesMovementId { get; set; }

        [Display(Name = "Запись в архиве данных СИР")]
        public MesMovementsDTO? MesMovementDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"ИД предыдущей записи\"")]
        [Display(Name = "ИД предыдущей записи")]
        public string? PreviousErpId { get; set; }

        [Display(Name = "Предыдущая запись")]
        public SapMovementsINDTO? PreviousRecordDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тип движения\"")]
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

        [NotMapped]
        [Display(Name = "Материал SAP")]
        public string? ToStringSapMaterialDTOFK
        {
            get
            {
                if (SapMaterialDTOFK != null)
                    return SapMaterialDTOFK.ToStringCodeName;
                else
                    return "НЕ НАЙДЕН";
            }
            set
            {
                ToStringSapMaterialDTOFK = value;
            }
        }

        [NotMapped]
        [Display(Name = "Источник в СИР")]
        public string? ToStringSapEquipmentSourceDTOFK
        {
            get
            {
                if (SapEquipmentSourceDTOFK != null)
                    return SapEquipmentSourceDTOFK.ToStringErpPlantIdErpIdName;
                else
                    return "НЕ НАЙДЕН";
            }
            set
            {
                ToStringSapEquipmentSourceDTOFK = value;
            }
        }


        [NotMapped]
        [Display(Name = "Приёмник в СИР")]
        public string? ToStringSapEquipmentDestDTOFK
        {
            get
            {
                if (SapEquipmentDestDTOFK != null)
                    return SapEquipmentDestDTOFK.ToStringErpPlantIdErpIdName;
                else
                    return "НЕ НАЙДЕН";
            }
            set
            {
                ToStringSapEquipmentDestDTOFK = value;
            }
        }

        public override string ToString()
        {
            return $"{ErpId}";
        }
    }
}

