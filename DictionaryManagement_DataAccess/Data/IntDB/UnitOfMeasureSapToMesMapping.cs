using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("UnitOfMeasureSapToMesMapping", Schema = "dbo")]
    public class UnitOfMeasureSapToMesMapping
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int Id { get; set; }

        public int SapUnitId { get; set; }

        [ForeignKey("SapUnitId")]
        public SapUnitOfMeasure? SapUnitOfMeasure { get; set; }

        public int MesUnitId { get; set; }

        [ForeignKey("MesUnitId")]
        public MesUnitOfMeasure? MesUnitOfMeasure { get; set; }

        [Required]
        [Range(0.0001, 1000000000, ErrorMessage = "Значение должно быть между {1} and {2}")]
        public decimal SapToMesTransformKoef { get; set; } = decimal.One;

    }

}
