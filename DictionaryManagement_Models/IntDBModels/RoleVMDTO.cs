using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

/* ViewModel для списка ролей для отображения в справочнике и управления привязками к шаблонам отчётов и пользователям */

namespace DictionaryManagement_Models.IntDBModels
{
    public class RoleVMDTO
    {
        [ForLogAttribute(NameProperty = "поле \"Ид роли\"")]
        [Display(Name = "Ид роли")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование обязательно для заполнения")]
        [StringLength(250, MinimumLength = 1, ErrorMessage = "Наименование может быть от 1 до 100 символов")]
        [Display(Name = "Наименование")]
        public string Name { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Описание\"")]
        [Display(Name = "Описание")]
        public string? Description { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Использование Админки чтение/запись\"")]
        [Display(Name = "Использование Админки чтение/запись")]
        public bool? IsAdmin { get; set; } = false;

        [ForLogAttribute(NameProperty = "поле \"Использование Админки только чтение\"")]
        [Display(Name = "Использование Админки только чтение")]
        public bool? IsAdminReadOnly { get; set; } = false;

        [NotMapped]
        public string AdminString
        {
            get
            {
                return IsAdmin == true ? "(Админка - чтение и запись)" : (IsAdminReadOnly == true ? "(Админка - только чтение)" : "");
            }
            set
            {
                AdminString = value;
            }
        }

        [NotMapped]
        public string NameWithAdminString
        {
            get
            {
                return (Name + (IsAdmin == true ? " (Админка - чтение и запись)" : (IsAdminReadOnly == true ? " (Админка - только чтение)" : "")));
            }
            set
            {
                NameWithAdminString = value;
            }
        }

        [NotMapped]
        public IEnumerable<UserToRoleDTO>? UserToRoleDTOs { get; set; }
        [NotMapped]
        public IEnumerable<ReportTemplateTypeTоRoleDTO>? ReportTemplateTypeTоRoleDTOs { get; set; }

        [NotMapped]
        public IEnumerable<RoleToADGroupDTO>? RoleToADGroupDTOs { get; set; }

        [NotMapped]
        public IEnumerable<RoleToDepartmentDTO>? RoleToDepartmentDTOs { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; } = false;

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
        [Display(Name = "Для фильтра по пользователям")]
        public string UsersLoginAndNameString
        {
            get
            {
                string retVar = "";
                foreach (var userToRoleItem in UserToRoleDTOs)
                {
                    retVar = retVar + userToRoleItem.UserDTOFK.Login + " " + userToRoleItem.UserDTOFK.UserName + " ";
                }
                return retVar;
            }
            set
            {
                UsersLoginAndNameString = value;
            }
        }

        [NotMapped]
        [Display(Name = "Для фильтра по типам шаблонов отчётов")]
        public string ReportTemplateTypeString
        {
            get
            {
                string retVar = "";
                foreach (var reportTemplateTypeTоRoleItem in ReportTemplateTypeTоRoleDTOs)
                {
                    retVar = retVar + reportTemplateTypeTоRoleItem.ReportTemplateTypeDTOFK.Name + " ";
                }
                return retVar;
            }
            set
            {
                ReportTemplateTypeString = value;
            }
        }

        [NotMapped]
        [Display(Name = "Для фильтра по группам AD")]
        public string RoleToADGroupString
        {
            get
            {
                string retVar = "";
                foreach (var roleToADGroupItem in RoleToADGroupDTOs)
                {
                    retVar = retVar + roleToADGroupItem.ADGroupDTOFK.Name + " ";
                }
                return retVar;
            }
            set
            {
                RoleToADGroupString = value;
            }
        }

        [NotMapped]
        [Display(Name = "Для фильтра по производствам")]
        public string RoleToDepartmentString
        {
            get
            {
                string retVar = "";
                foreach (var roleToDepartmentItem in RoleToDepartmentDTOs)
                {
                    retVar = retVar + roleToDepartmentItem.DepartmentDTOFK.ShortName + " ";
                }
                return retVar;
            }
            set
            {
                RoleToDepartmentString = value;
            }
        }

        public override string ToString()
        {
            return $"{Name}";
        }

    }
}

