using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("DataType", Schema = "dbo")]
    public class DataType
    {
        [Key]
        [Required]
        public int Id { get; set; }

        [Required]
        [MaxLength(250)]
        [MinLength(1)]
        public string Name { get; set; } = string.Empty;

        public int? Priority { get; set; }

        public bool IsArchive { get; set; }

        public bool? IsAutoCalcDestDataType { get; set; }
    }
}
