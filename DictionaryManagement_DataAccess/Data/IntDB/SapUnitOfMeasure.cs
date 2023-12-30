using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("SapUnitOfMeasure", Schema = "dbo")]
    public class SapUnitOfMeasure
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [MaxLength(250)]
        [MinLength(1)]
        public string Name { get; set; } = string.Empty;

        [Required]
        [MaxLength(100)]
        [MinLength(1)]
        public string ShortName { get; set; } = string.Empty;

        public bool IsArchive { get; set; }
    }
}
