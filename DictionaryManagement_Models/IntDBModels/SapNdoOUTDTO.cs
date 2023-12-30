using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapNdoOUTDTO
    {
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД записи обязателен")]
        public Int64 Id { get; set; }

        [Display(Name = "Время добавления записи")]
        [Required(ErrorMessage = "Время добавления записи обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Display(Name = "Имя тэга")]
        [Required(ErrorMessage = "Имя тэга обязательно")]
        public string TagName { get; set; }

        [Display(Name = "Время значения")]
        [Required(ErrorMessage = "Время значения обязательно")]
        public DateTime ValueTime { get; set; }

        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [Display(Name = "Признак передачи в SAP")]
        [Required(ErrorMessage = "Признак передачи в SAP обязателен")]
        public bool SapGone { get; set; } = false;

        [Display(Name = "Признак передачи в SAP")]
        public DateTime? SapGoneTime { get; set; }

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
    }
}

