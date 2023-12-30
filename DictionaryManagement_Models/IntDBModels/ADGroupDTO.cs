using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ADGroupDTO
    {

        [ForLogAttribute(NameProperty = "поле \"Ид группы\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование обязательно для заполнения")]
        [StringLength(300, MinimumLength = 1, ErrorMessage = "Наименование может быть от 1 до 300 символов")]
        [Display(Name = "Наименование")]
        public string Name { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Описание\"")]
        [StringLength(300, MinimumLength = 1, ErrorMessage = "Описание может быть от 1 до 300 символов")]
        [Display(Description = "Описание")]
        public string? Description { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; } = false;

        [NotMapped]
        [Display(Name = " ")]
        public bool Checked { get; set; } = false;

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
        public override string ToString()
        {
            return $"{Name}";
        }
    }
}

