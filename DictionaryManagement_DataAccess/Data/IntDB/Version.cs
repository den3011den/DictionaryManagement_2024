using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{

    [Table("Version", Schema = "dbo")]
    public class Version
    {
        [Key]
        [Required]
        public int Id { get; set; }

        [Required]
        [Column("Version")]
        public string? version { get; set; }

    }
}
