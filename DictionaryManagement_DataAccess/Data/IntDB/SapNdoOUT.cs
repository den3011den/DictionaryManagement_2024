using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("SapNdoOUT", Schema = "dbo")]
    public class SapNdoOUT
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        [Required]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Required]
        public string TagName { get; set; }

        [Required]
        public DateTime ValueTime { get; set; }

        [Required]
        public decimal Value { get; set; } = decimal.Zero;

        [Required]
        public bool SapGone { get; set; } = false;

        public DateTime? SapGoneTime { get; set; }

    }
}
