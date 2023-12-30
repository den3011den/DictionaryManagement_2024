using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("LogEvent", Schema = "dbo")]
    public class LogEvent
    {

        [Key]
        [Required]
        public Int64 Id { get; set; }

        [Required]
        public int LogEventTypeId { get; set; }

        [ForeignKey("LogEventTypeId")]
        public LogEventType LogEventTypeFK { get; set; }

        public string? OldValue { get; set; }
        public string? NewValue { get; set; }

        [Required]
        public DateTime EventTime { get; set; } = DateTime.Now;

        [Required]
        public Guid UserId { get; set; }
        [ForeignKey("UserId")]
        public User UserFK { get; set; }

        public string Description { get; set; }

        public bool? IsCritical { get; set; } = false;
        public bool? IsError { get; set; } = false;
        public bool? IsWarning { get; set; } = false;

    }
}
