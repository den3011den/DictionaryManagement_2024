using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesNdoStocksDTO
    {
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Int64 Id { get; set; }

        [Display(Name = "ИД тэга СИР")]
        [Required(ErrorMessage = "ИД тэга СИР обязателен")]
        public int MesParamId { get; set; }

        [Display(Name = "Тэг СИР")]
        [Required(ErrorMessage = "Выбор тэга СИР обязателен")]
        public MesParamDTO MesParamDTOFK { get; set; }

        [Display(Name = "Время добавления записи")]
        [Required(ErrorMessage = "Время добавления записи обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Display(Name = "ИД пользователя")]
        [Required(ErrorMessage = "ИД пользователя добавившего запись обязателен")]
        public Guid AddUserId { get; set; }

        [Display(Name = "Пользователь")]
        [Required(ErrorMessage = "Пользователя добавивший запись обязателен")]
        public UserDTO AddUserDTOFK { get; set; }

        [Display(Name = "Время значения")]
        [Required(ErrorMessage = "Время значения обязательно")]
        public DateTime ValueTime { get; set; }

        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [Required(ErrorMessage = "Разность с предыдущим")]
        public decimal? ValueDifference { get; set; } = decimal.Zero;

        [Display(Name = "ИД экземпляра отчёта")]
        public Guid? ReportGuid { get; set; }

        [Display(Name = "Экземпляр отчёта")]
        public ReportEntityDTO? ReportEntityDTOFK { get; set; }

        [Display(Name = "ИД записи в зеркале SAP")]
        public Int64? SapNdoOutId { get; set; }

        public SapNdoOUTDTO? SapNdoOUTDTOFK { get; set; }

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
        [Display(Name = "Разность")]
        public string? ToStringValueDifference
        {
            get
            {
                return ValueDifference.ToString();
            }
            set
            {
                ToStringValueDifference = value;
            }
        }

        [NotMapped]
        [Display(Name = "Sap забрал")]
        public bool SapGoneBool
        {
            get
            {
                if (SapNdoOUTDTOFK == null)
                {
                    return false;
                }
                else
                {
                    return SapNdoOUTDTOFK.SapGone == null ? false : SapNdoOUTDTOFK.SapGone;
                }
            }
            set
            {
                SapGoneBool = value;
            }
        }
    }
}

