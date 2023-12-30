using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("User", Schema = "dbo")]
    public class User
    {

        [Key]
        [Required]
        public Guid Id { get; set; }

        [Required]
        public string Login { get; set; }

        public string? UserName { get; set; }

        public string? Description { get; set; }

        public bool IsArchive { get; set; } = false;

        public bool? IsSyncWithAD { get; set; } = true;

        public DateTime? SyncWithADGroupsLastTime { get; set; } = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

        public bool? IsServiceUser { get; set; } = false;
    }

}
