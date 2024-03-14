using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ReportTemplateToMesParamDTO
    {
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код обязателен для заполнения")]
        public int Id { get; set; }

        [Required(ErrorMessage = "Шаблон отчёта обязателен")]
        [Display(Name = "ИД шаблона отчёта")]
        public Guid ReportTemplateId { get; set; }

        [Display(Name = "Шаблон отчёта")]
        [Required(ErrorMessage = "Шаблон отчёта обязателен")]
        public ReportTemplateDTO? ReportTemplateDTOFK { get; set; }

        [Display(Name = "ИД тэга СИР")]
        public int? MesParamId { get; set; }

        [Display(Name = "Тэг СИР")]
        public MesParamDTO? MesParamDTOFK { get; set; }

        [Display(Name = "Код тэга СИР")]
        public string? MesParamCode { get; set; }

        [Display(Name = "Имя листа файла шаблона отчёта")]
        [Required(ErrorMessage = "Имя листа файла шаблона отчёта обязательно к заполнению")]
        public string SheetName { get; set; } = "";

        [NotMapped]
        [Display(Name = "ИД")]
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
            return "Шаблон: \"{ReportTemplateDTOFK.Id.ToString}\" " + "Тэг: \"" + MesParamDTOFK != null ? MesParamDTOFK.Id.ToString() : ""
                + "\" Код тэга: \"" + MesParamCode != null ? MesParamCode : "" + "Лист: \"" + SheetName + "\"";
        }

    }

}
