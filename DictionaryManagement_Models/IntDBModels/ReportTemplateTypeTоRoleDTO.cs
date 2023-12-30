using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportTemplateTypeTоRoleDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        [Required(ErrorMessage = "ИД типа отчёта обязателен")]
        [Display(Name = "Ид тип отчёта")]
        public int ReportTemplateTypeId { get; set; }

        [Required(ErrorMessage = "Выбор типа отчёта обязателен")]
        [Display(Name = "Тип отчёта")]
        public ReportTemplateTypeDTO ReportTemplateTypeDTOFK { get; set; }

        [Required(ErrorMessage = "ИД роли обязателен")]
        [Display(Name = "Ид роли")]
        public Guid RoleId { get; set; }

        [Required(ErrorMessage = "Выбор роли обязателен")]
        [Display(Name = "Роль")]
        public RoleDTO RoleDTOFK { get; set; }

        [Display(Name = "Может скачивать")]
        public bool CanDownload { get; set; }

        [Display(Name = "Может загружать")]
        public bool CanUpload { get; set; }

    }
}

