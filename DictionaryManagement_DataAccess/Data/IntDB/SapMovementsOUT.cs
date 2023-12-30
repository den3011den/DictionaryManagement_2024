using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("SapMovementsOUT", Schema = "dbo")]
    public class SapMovementsOUT
    {

        [Key]
        [Required]
        public Guid Id { get; set; }

        [Required]
        public DateTime AddTime { get; set; } = DateTime.Now;

        public string? BatchNo { get; set; }

        [Required]
        public string SapMaterialCode { get; set; }

        [NotMapped]
        public SapMaterial? SapMaterialFK { get; set; }

        public string? ErpPlantIdSource { get; set; }

        public string? ErpIdSource { get; set; }

        [NotMapped]
        public SapEquipment? SapEquipmentSourceFK { get; set; }

        public bool? IsWarehouseSource { get; set; }

        public string? ErpPlantIdDest { get; set; }
        public string? ErpIdDest { get; set; }

        [NotMapped]
        public SapEquipment? SapEquipmentDestFK { get; set; }

        public bool? IsWarehouseDest { get; set; }

        [Required]
        public DateTime ValueTime { get; set; }

        [Required]
        public decimal Value { get; set; } = decimal.Zero;

        public decimal Correction2Previous { get; set; } = decimal.Zero;

        public bool? IsReconciled { get; set; }

        public string? SapUnitOfMeasure { get; set; }

        public bool? SapGone { get; set; }

        public DateTime? SapGoneTime { get; set; }

        public bool? SapError { get; set; }

        public string? SapErrorMessage { get; set; }

        public Guid? MesMovementId { get; set; }
        [ForeignKey("MesMovementId")]
        public MesMovements? MesMovementsFK { get; set; }

        public Guid? PreviousRecordId { get; set; }
        [ForeignKey("PreviousRecordId")]
        public SapMovementsOUT? PreviousRecordFK { get; set; }

        public int MesParamId { get; set; }
        [ForeignKey("MesParamId")]
        public MesParam MesParamFK { get; set; }


    }
}
