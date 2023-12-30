using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapMovementsOUTDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [Display(Name = "Время добавления")]
        [Required(ErrorMessage = "Время добавления обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Display(Name = "Номер партии")]
        public string? BatchNo { get; set; }

        [Display(Name = "Код материала SAP")]
        [Required(ErrorMessage = "Код материала SAP обязателен")]
        public string SapMaterialCode { get; set; }

        [Display(Name = "Материал SAP")]
        [Required(ErrorMessage = "Материал SAP обязателен")]
        public SapMaterialDTO SapMaterialDTOFK { get; set; }

        [Display(Name = "Код завода-источника SAP")]
        public string? ErpPlantIdSource { get; set; }

        [Display(Name = "Код ресурса/склада-источника SAP")]
        public string? ErpIdSource { get; set; }

        [Display(Name = "Источник SAP")]
        public virtual SapEquipmentDTO SapEquipmentSourceDTOFK { get; set; }

        [Display(Name = "Склад-источник SAP")]
        public bool? IsWarehouseSource { get; set; }

        [Display(Name = "Код завода-приемника SAP")]
        public string? ErpPlantIdDest { get; set; }

        [Display(Name = "Код ресурса/склада-приемника SAP")]
        public string? ErpIdDest { get; set; }

        [Display(Name = "Приёмник SAP")]
        public virtual SapEquipmentDTO SapEquipmentDestDTOFK { get; set; }


        [Display(Name = "Склад-приемник SAP")]
        public bool? IsWarehouseDest { get; set; }

        [Display(Name = "Время значения")]
        [Required(ErrorMessage = "Время значения обязательно")]
        public DateTime ValueTime { get; set; }

        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [Display(Name = "Корректировка")]
        [Required(ErrorMessage = "Корректировка обязательна")]
        public decimal Correction2Previous { get; set; } = decimal.Zero;

        [Display(Name = "Согласовано")]
        public bool? IsReconciled { get; set; }

        [Display(Name = "Ед. изм. SAP")]
        public string? SapUnitOfMeasure { get; set; }

        [Display(Name = "Ушло в SAP")]
        public bool? SapGone { get; set; }

        [Display(Name = "Время ушло в SAP")]
        public DateTime? SapGoneTime { get; set; }

        [Display(Name = "Ошибка SAP")]
        public bool? SapError { get; set; }

        [Display(Name = "Ошибка SAP сообщение")]
        public string? SapErrorMessage { get; set; }

        [Display(Name = "ИД записи в архиве данных")]
        public Guid? MesMovementId { get; set; }

        [Display(Name = "ИД записи в архиве данных")]
        public MesMovementsDTO? MesMovementsDTOFK { get; set; }

        [Display(Name = "ИД предыдущей записи")]
        public Guid? PreviousRecordId { get; set; }

        [Display(Name = "Предыдущая запись")]
        public SapMovementsOUTDTO? PreviousRecordDTOFK { get; set; }

        [Display(Name = "ИД тэга СИР")]
        [Required(ErrorMessage = "ИД тэга СИР обязательно")]
        public int MesParamId { get; set; }

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


    }
}

