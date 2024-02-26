using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("Scheduler", Schema = "dbo")]
    public class Scheduler
    {

        [Key]
        [Required]
        public int Id { get; set; }

        [Required]
        public string ModuleName { get; set; }

        [Required]
        public TimeSpan StartTime { get; set; }

        public DateTime? LastExecuted { get; set; } = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

        public bool? IsRunningNow { get; set; } = false;
    }
}
