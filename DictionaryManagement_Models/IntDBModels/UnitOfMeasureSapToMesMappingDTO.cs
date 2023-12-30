using DictionaryManagement_Models.IntDBModels;
using System.ComponentModel.DataAnnotations;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    public class UnitOfMeasureSapToMesMappingDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "Код обязателен для заполнения")]
        public int Id { get; set; }

        [Required]
        [Display(Name = "Код единицы измерения SAP")]
        public int SapUnitId { get; set; }

        [Required]
        [Display(Name = "Единица измерения SAP")]
        public SapUnitOfMeasureDTO? SapUnitOfMeasureDTO { get; set; }

        [Required]
        [Display(Name = "Код единицы измерения MES")]
        public int MesUnitId { get; set; }

        [Required]
        [Display(Name = "Единица измерения MES")]
        public MesUnitOfMeasureDTO? MesUnitOfMeasureDTO { get; set; }

        [Required]
        [Range(0.0001, 1000000000, ErrorMessage = "Значение должно быть между {1} and {2}")]
        [Display(Name = "Коэф. пересчёта ед. изм. SAP в MES")]
        public decimal SapToMesTransformKoef { get; set; } = decimal.One;


    }

}
