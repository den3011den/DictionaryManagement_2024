using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("Smena", Schema = "dbo")]
    public class Smena
    {

        [Key]
        [Required]
        public int Id { get; set; }

        [Required]
        public string Name { get; set; }

        [Required]
        public int DepartmentId { get; set; }

        [ForeignKey("DepartmentId")]
        public MesDepartment DepartmentFK { get; set; }

        [Required]
        public TimeSpan StartTime { get; set; }

        [Required]
        public byte HoursDuration { get; set; }

        public bool? IsArchive { get; set; }
    }
}
