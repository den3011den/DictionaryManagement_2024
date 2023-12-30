using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_DataAccess.Data.IntDB
{
    [Table("MesDepartment", Schema = "dbo")]
    public class MesDepartment
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int Id { get; set; }

        public int? MesCode { get; set; }

        public string Name { get; set; }

        public string ShortName { get; set; }

        public int? ParentDepartmentId { get; set; }

        [ForeignKey("ParentDepartmentId")]
        public MesDepartment? DepartmentParent { get; set; }

        public bool IsArchive { get; set; }
    }

}
