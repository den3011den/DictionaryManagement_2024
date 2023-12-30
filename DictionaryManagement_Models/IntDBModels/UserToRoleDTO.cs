using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class UserToRoleDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        [Required(ErrorMessage = "ИД пользователя обязателен")]
        [Display(Name = "Ид пользователя")]
        public Guid UserId { get; set; }

        [Required(ErrorMessage = "Выбор пользователя обязателен")]
        [Display(Name = "Прользователь")]
        public UserDTO UserDTOFK { get; set; }

        [Required(ErrorMessage = "ИД роли обязателен")]
        [Display(Name = "Ид роли")]
        public Guid RoleId { get; set; }

        [Required(ErrorMessage = "Выбор роли обязателен")]
        [Display(Name = "Роль")]
        public RoleDTO RoleDTOFK { get; set; }

    }
}

