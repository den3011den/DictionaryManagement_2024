using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class RoleToADGroupDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        [Required(ErrorMessage = "ИД роли обязателен")]
        [Display(Name = "Ид роли")]
        public Guid RoleId { get; set; }

        [Required(ErrorMessage = "Выбор роли обязателен")]
        [Display(Name = "Роль")]
        public RoleDTO RoleDTOFK { get; set; }

        [Required(ErrorMessage = "ИД группы AD обязателен")]
        [Display(Name = "Ид группы AD")]
        public Guid ADGroupId { get; set; }

        [Required(ErrorMessage = "Выбор группы AD обязателен")]
        [Display(Name = "AD группа")]
        public ADGroupDTO ADGroupDTOFK { get; set; }

    }
}

