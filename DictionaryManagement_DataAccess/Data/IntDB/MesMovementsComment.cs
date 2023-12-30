using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("MesMovementsComment", Schema = "dbo")]
    public class MesMovementsComment
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        [Required]
        public Guid MesMovementsId { get; set; }
        [ForeignKey("MesMovementsId")]
        public virtual MesMovements MesMovementsFK { get; set; }

        public int? CorrectionReasonId { get; set; }
        [ForeignKey("CorrectionReasonId")]
        public CorrectionReason? CorrectionReasonFK { get; set; }

        public string? CorrectionComment { get; set; }

    }
}
