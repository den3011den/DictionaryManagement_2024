using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapMovementsOUTDTO
    {
        [ForLogAttribute(NameProperty = "поле \"ИД записи\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время добавления\"")]
        [Display(Name = "Время добавления")]
        [Required(ErrorMessage = "Время добавления обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [ForLogAttribute(NameProperty = "поле \"Номер партии\"")]
        [Display(Name = "Номер партии")]
        public string? BatchNo { get; set; }


        [Display(Name = "Код материала SAP")]
        [Required(ErrorMessage = "Код материала SAP обязателен")]
        public string SapMaterialCode { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Материал SAP\"")]
        [Display(Name = "Материал SAP")]
        [Required(ErrorMessage = "Материал SAP обязателен")]
        public SapMaterialDTO SapMaterialDTOFK { get; set; }

        [Display(Name = "Код завода-источника SAP")]
        public string? ErpPlantIdSource { get; set; }

        [Display(Name = "Код ресурса/склада-источника SAP")]
        public string? ErpIdSource { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Источник SAP\"")]
        [Display(Name = "Источник SAP")]
        public virtual SapEquipmentDTO SapEquipmentSourceDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Склад-источник SAP\"")]
        [Display(Name = "Склад-источник SAP")]
        public bool? IsWarehouseSource { get; set; }

        [Display(Name = "Код завода-приемника SAP")]
        public string? ErpPlantIdDest { get; set; }


        public string? ErpIdDest { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Приёмник SAP\"")]
        [Display(Name = "Приёмник SAP")]
        public virtual SapEquipmentDTO SapEquipmentDestDTOFK { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Склад-приемник SAP\"")]
        [Display(Name = "Склад-приемник SAP")]
        public bool? IsWarehouseDest { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время значения\"")]
        [Display(Name = "Время значения")]
        [Required(ErrorMessage = "Время значения обязательно")]
        public DateTime ValueTime { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Значение\"")]
        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [ForLogAttribute(NameProperty = "поле \"Корректировка\"")]
        [Display(Name = "Корректировка")]
        [Required(ErrorMessage = "Корректировка обязательна")]
        public decimal Correction2Previous { get; set; } = decimal.Zero;

        [ForLogAttribute(NameProperty = "поле \"Согласовано\"")]
        [Display(Name = "Согласовано")]
        public bool? IsReconciled { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Ед. изм. SAP\"")]
        [Display(Name = "Ед. изм. SAP")]
        public string? SapUnitOfMeasure { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Ушло в SAP\"")]
        [Display(Name = "Ушло в SAP")]
        public bool? SapGone { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Время ушло в SAP\"")]
        [Display(Name = "Время ушло в SAP")]
        public DateTime? SapGoneTime { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Ошибка SAP\"")]
        [Display(Name = "Ошибка SAP")]
        public bool? SapError { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Ошибка SAP сообщение\"")]
        [Display(Name = "Ошибка SAP сообщение")]
        public string? SapErrorMessage { get; set; }

        [Display(Name = "ИД записи в архиве данных")]
        public Guid? MesMovementId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Запись в архиве данных\"")]
        [Display(Name = "ИД записи в архиве данных")]
        public MesMovementsDTO? MesMovementsDTOFK { get; set; }

        [Display(Name = "ИД предыдущей записи")]
        public Guid? PreviousRecordId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Предыдущая запись\"")]
        [Display(Name = "Предыдущая запись")]
        public SapMovementsOUTDTO? PreviousRecordDTOFK { get; set; }

        [Display(Name = "ИД тэга СИР")]
        [Required(ErrorMessage = "ИД тэга СИР обязательно")]
        public int MesParamId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Тэг СИР\"")]
        [Display(Name = "Тэг СИР")]
        [Required(ErrorMessage = "Тэг СИР обязателен")]
        public MesParamDTO MesParamDTOFK { get; set; }

        [NotMapped]
        private string _toStringId;

        [NotMapped]
        [Display(Name = "Ид записи")]
        public string ToStringId
        {
            get
            {
                _toStringId = Id.ToString();
                return _toStringId;
            }
            set
            {
                _toStringId = value;
            }
        }

        [NotMapped]
        private string _toStringCorrection2Previous;

        [NotMapped]
        [Display(Name = "Корректировка")]
        public string ToStringCorrection2Previous
        {
            get
            {
                _toStringCorrection2Previous = Correction2Previous.ToString();
                return _toStringCorrection2Previous;
            }
            set
            {
                _toStringCorrection2Previous = value;
            }
        }

        [NotMapped]
        public string _toStringValue;

        [NotMapped]
        [Display(Name = "Значение")]
        public string ToStringValue
        {
            get
            {
                _toStringValue = Value.ToString();
                return _toStringValue;
            }
            set
            {
                _toStringValue = value;
            }
        }

        [NotMapped]
        private string? _toStringMesMovementId;

        [NotMapped]
        [Display(Name = "ИД записи в архиве данных")]
        public string? ToStringMesMovementId
        {
            get
            {
                _toStringMesMovementId = MesMovementId == null ? "" : MesMovementId.ToString();
                return _toStringMesMovementId;
            }
            set
            {
                _toStringMesMovementId = value;
            }
        }

        [NotMapped]
        private string? _toStringPreviousRecordId;

        [NotMapped]
        [Display(Name = "ИД предыдущей записи")]
        public string? ToStringPreviousRecordId
        {
            get
            {
                _toStringPreviousRecordId = PreviousRecordId == null ? "" : PreviousRecordId.ToString();
                return _toStringPreviousRecordId;
            }
            set
            {
                _toStringPreviousRecordId = value;
            }
        }

        [NotMapped]
        private bool _sapErrorBool;

        [NotMapped]
        [Display(Name = "Sap ошибка")]
        public bool SapErrorBool
        {
            get
            {
#pragma warning disable IDE0075 // Simplify conditional expression
                _sapErrorBool = SapError == null ? false : (bool)SapError;
#pragma warning restore IDE0075 // Simplify conditional expression
                return _sapErrorBool;
            }
            set
            {
                _sapErrorBool = value;
            }
        }

        [NotMapped]
        private bool _sapGoneBool;

        [NotMapped]
        [Display(Name = "Sap обработал")]
        public bool SapGoneBool
        {
            get
            {
#pragma warning disable IDE0075 // Simplify conditional expression
                _sapGoneBool = SapGone == null ? false : (bool)SapGone;
#pragma warning restore IDE0075 // Simplify conditional expression
                return _sapGoneBool;
            }
            set
            {
                _sapGoneBool = value;
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
#pragma warning disable IDE0071 // Simplify interpolation
            return $"{Id.ToString()}";
#pragma warning restore IDE0071 // Simplify interpolation
        }
    }
}

