using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class RoleDTO
    {
        [ForLogAttribute(NameProperty = "поле \"Ид роли\"")]
        [Display(Name = "Ид записи")]
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

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; } = false;

        [ForLogAttribute(NameProperty = "поле \"Использование Админки чтение/запись\"")]
        [Display(Name = "Использование Админки чтение/запись")]
        public bool IsAdmin { get; set; } = false;

        [ForLogAttribute(NameProperty = "поле \"Использование Админки только чтение\"")]
        [Display(Name = "Использование Админки только чтение")]
        public bool IsAdminReadOnly { get; set; } = false;

        public override string ToString()
        {
            return $"{Name}";
        }
    }
}

