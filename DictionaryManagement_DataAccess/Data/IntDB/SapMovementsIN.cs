using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("SapMovementsIN", Schema = "dbo")]
    public class SapMovementsIN
    {

        [Key]
        [Required]
        public string ErpId { get; set; }

        [Required]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Required]
        public DateTime SapDocumentEnterTime { get; set; }

        [Required]
        public DateTime SapDocumentPostTime { get; set; }

        public string? BatchNo { get; set; }

        [Required]
        public string SapMaterialCode { get; set; }

        [NotMapped]
        public SapMaterial? SapMaterialFK { get; set; }

        [Required]
        public string ErpPlantIdSource { get; set; }

        [Required]
        public string ErpIdSource { get; set; }

        [NotMapped]
        public SapEquipment? SapEquipmentSourceFK { get; set; }

        public bool? IsWarehouseSource { get; set; }

        [Required]
        public string ErpPlantIdDest { get; set; }

        [Required]
        public string ErpIdDest { get; set; }

        [NotMapped]
        public SapEquipment? SapEquipmentDestFK { get; set; }

        public bool? IsWarehouseDest { get; set; }

        [Required]
        public decimal Value { get; set; } = decimal.Zero;

        [Required]
        public string SapUnitOfMeasure { get; set; }

        public bool? IsStorno { get; set; }

        public bool? MesGone { get; set; }

        public DateTime? MesGoneTime { get; set; }

        public bool? MesError { get; set; }

        public string? MesErrorMessage { get; set; }

        public Guid? MesMovementId { get; set; }
        [ForeignKey("MesMovementId")]
        public MesMovements? MesMovementFK { get; set; }

        public string? PreviousErpId { get; set; }
        [ForeignKey("PreviousErpId")]
        public SapMovementsIN? PreviousRecordFK { get; set; }

        public string? MoveType { get; set; }

        [NotMapped]
        public SapUnitOfMeasure? SapUnitOfMeasureFK { get; set; }
    }
}
