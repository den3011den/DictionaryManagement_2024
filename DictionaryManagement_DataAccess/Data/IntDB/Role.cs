using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("Role", Schema = "dbo")]
    public class Role
    {

        [Key]
        [Required]
        public Guid Id { get; set; }

        [Required]
        public string Name { get; set; }

        public string? Description { get; set; }

        public bool IsArchive { get; set; } = false;

        public bool? IsAdmin { get; set; } = false;

        public bool? IsAdminReadOnly { get; set; } = false;
    }

}
