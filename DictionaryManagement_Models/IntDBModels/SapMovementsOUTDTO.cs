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

        [NotMapped]
        [Display(Name = "Корректировка")]
        public string ToStringCorrection2Previous
        {
            get
            {
                return Correction2Previous.ToString();
            }
            set
            {
                ToStringCorrection2Previous = value;
            }
        }

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
        [Display(Name = "ИД предыдущей записи")]
        public string? ToStringPreviousRecordId
        {
            get
            {
                return PreviousRecordId.ToString();
            }
            set
            {
                ToStringPreviousRecordId = value;
            }
        }

        [NotMapped]
        [Display(Name = "Sap ошибка")]
        public bool SapErrorBool
        {
            get
            {
                return SapError == null ? false : (bool)SapError;
            }
            set
            {
                SapErrorBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Sap забрал")]
        public bool SapGoneBool
        {
            get
            {
                return SapGone == null ? false : (bool)SapGone;
            }
            set
            {
                SapGoneBool = value;
            }
        }

        public override string ToString()
        {
            return $"{Id.ToString()}";
        }
    }
}

