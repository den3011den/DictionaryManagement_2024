using DictionaryManagement_Models.IntDBModels;
using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    public class SapToMesMaterialMappingDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код обязателен для заполнения")]
        public int Id { get; set; }

        [Required]
        [Display(Name = "Код материала SAP")]
        public int SapMaterialId { get; set; }

        [Required]
        [Display(Name = "Материал MES")]
        public SapMaterialDTO? SapMaterialDTO { get; set; }

        [Required]
        [Display(Name = "Код материала MES")]
        public int MesMaterialId { get; set; }

        [Required]
        [Display(Name = "Материал MES")]
        public MesMaterialDTO? MesMaterialDTO { get; set; }

    }

}
