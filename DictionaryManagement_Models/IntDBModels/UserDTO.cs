using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class UserDTO
    {

        [ForLogAttribute(NameProperty = "поле \"ИД пользователя\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Логин\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "Логин обязателен для заполнения")]
        [StringLength(250, MinimumLength = 1, ErrorMessage = "Логин может быть от 1 до 250 символов")]
        [Display(Name = "Логин")]
        public string Login { get; set; }

        [ForLogAttribute(NameProperty = "поле \"ФИО\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Required(ErrorMessage = "ФИО обязательно для заполнения")]
        [StringLength(250, MinimumLength = 3, ErrorMessage = "ФИО может быть от 3 до 250 символов")]
        [Display(Name = "ФИО")]
        public string? UserName { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Описание\"")]
        [CheckControlSymbols]
        [CheckLeadingAndTrailingSpaces]
        [Display(Name = "Описание")]
        public string? Description { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Автомат\"")]
        [Required(ErrorMessage = "Параметр является обязательным")]
        [Display(Name = "Синх с AD")]
        public bool IsSyncWithAD { get; set; } = true;

        [ForLogAttribute(NameProperty = "поле \"Время посл. синх. AD\"")]
        [Required(ErrorMessage = "Параметр является обязательным")]
        [Display(Name = "Время последней синхронизации с группами AD")]
        public DateTime SyncWithADGroupsLastTime { get; set; } = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

        [ForLogAttribute(NameProperty = "поле \"Сервисный пользователь\"")]
        [Display(Name = "Сервисный пользователь")]
        public bool IsServiceUser { get; set; } = false;


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

        [NotMapped]
        [Display(Name = "Имя и логин")]
        public string UserNameAndLogin
        {
            get
            {
                return UserName + " (" + Login + ")";
            }
            set
            {
                UserNameAndLogin = value;
            }
        }

        public override string ToString()
        {
            return $"{UserName}";
        }
    }
}

