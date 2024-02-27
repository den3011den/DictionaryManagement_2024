using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("MesParam", Schema = "dbo")]
    public class MesParam
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int Id { get; set; }

        public string Code { get; set; }

        public string? Name { get; set; }

        public string? Description { get; set; }

        public int? MesParamSourceType { get; set; }

        [ForeignKey("MesParamSourceType")]
        public MesParamSourceType? MesParamSourceTypeFK { get; set; }

        public string? MesParamSourceLink { get; set; }

        public int? DepartmentId { get; set; }

        [ForeignKey("DepartmentId")]
        public MesDepartment? MesDepartmentFK { get; set; }

        public int? SapEquipmentIdSource { get; set; }

        [ForeignKey("SapEquipmentIdSource")]
        public SapEquipment? SapEquipmentSourceFK { get; set; }

        public int? SapEquipmentIdDest { get; set; }
        [ForeignKey("SapEquipmentIdDest")]
        public SapEquipment? SapEquipmentDestFK { get; set; }

        public int? MesMaterialId { get; set; }
        [ForeignKey("MesMaterialId")]
        public MesMaterial? MesMaterialFK { get; set; }

        public int? SapMaterialId { get; set; }
        [ForeignKey("SapMaterialId")]
        public SapMaterial? SapMaterialFK { get; set; }

        public int? MesUnitOfMeasureId { get; set; }
        [ForeignKey("MesUnitOfMeasureId")]
        public MesUnitOfMeasure? MesUnitOfMeasureFK { get; set; }

        public int? SapUnitOfMeasureId { get; set; }
        [ForeignKey("SapUnitOfMeasureId")]
        public SapUnitOfMeasure? SapUnitOfMeasureFK { get; set; }

        public int? DaysRequestInPast { get; set; }

        public string? TI { get; set; }
        public string? NameTI { get; set; }
        public string? TM { get; set; }
        public string? NameTM { get; set; }

        [Precision(12, 6)]
        public decimal? MesToSirUnitOfMeasureKoef { get; set; } = decimal.One;

        public bool? NeedWriteToSap { get; set; }
        public bool? NeedReadFromSap { get; set; }
        public bool? NeedReadFromMes { get; set; }
        public bool? NeedWriteToMes { get; set; }
        public bool? IsNdo { get; set; }
        //[Display(Name = "Это аннотация архива")]
        public bool IsArchive { get; set; }
    }

}
