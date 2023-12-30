using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesDepartmentVMDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        [Required(ErrorMessage = "Код обязателен для заполнения")]
        [Range(1, int.MaxValue, ErrorMessage = "Код может быть от {1} до {2}")]
        [Display(Name = "Код")]
        public int? MesCode { get; set; }

        [Required(ErrorMessage = "Наименование обязательно для заполнения")]
        [Display(Name = "Наименование")]
        [MaxLength(500, ErrorMessage = "Длина наименования не может быть больше 500 символов")]
        public string Name { get; set; } = string.Empty;

        [Required(ErrorMessage = "Сокр. Наименование обязательно для заполнения")]
        [Display(Name = "Сокр. наименование")]
        [MaxLength(500, ErrorMessage = "Длина сокр. наименование не может быть больше 500 символов")]
        public string ShortName { get; set; } = string.Empty;

        [Display(Name = "Ид родителя")]
        public int? ParentDepartmentId { get; set; }

        [Display(Name = "Родитель")]
        public MesDepartmentVMDTO? DepartmentParentVMDTO { get; set; }

        [Display(Name = "Дети")]
        public IEnumerable<MesDepartmentVMDTO>? ChildrenDTO { get; set; }


        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [NotMapped]
        public string ToStringValue { get; set; } = string.Empty;

        public override string ToString()
        {
            ToStringValue = $"{ShortName}";
            return ToStringValue;
        }

        [NotMapped]
        [Display(Name = " ")]
        public bool Checked { get; set; } = false;

        [NotMapped]
        [Display(Name = "Уровень")]
        public int DepLevel { get; set; }

    }
}

