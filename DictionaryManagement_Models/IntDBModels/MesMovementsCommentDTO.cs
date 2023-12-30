using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace DictionaryManagement_Models.IntDBModels
{
    public class MesMovementsCommentDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Int64 Id { get; set; }

        [Display(Name = "Ид записи архива данных")]
        [Required(ErrorMessage = "Ид записи архива данных обязателен")]
        public Guid MesMovementsId { get; set; }

        [Display(Name = "Запись архива данных")]
        [Required(ErrorMessage = "Запись архива данных обязателена")]
        public MesMovementsDTO MesMovementsDTOFK { get; set; }

        [Display(Name = "ИД записи причины корректировки")]
        public int? CorrectionReasonId { get; set; }

        [Display(Name = "Запись причины корректировки")]
        public CorrectionReasonDTO? CorrectionReasonDTOFK { get; set; }

        [Display(Name = "Причина корректировки")]
        public string? CorrectionComment { get; set; }


        [NotMapped]
        [Display(Name = "Ид записи")]
        public string ToStringId
        {
            get
            {
                return Id.ToString();
            }
            set
            {
                ToStringId = value;
            }
        }

    }
}
