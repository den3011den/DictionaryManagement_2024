using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesMovementsDTO
    {

        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public Guid Id { get; set; }

        [Display(Name = "Время добавления")]
        [Required(ErrorMessage = "Время добавления обязательно")]
        public DateTime AddTime { get; set; } = DateTime.Now;

        [Display(Name = "ИД кто добавил")]
        [Required(ErrorMessage = "ИД кто добавил обязателен")]
        public Guid AddUserId { get; set; }

        [Display(Name = "Кто добавил")]
        [Required(ErrorMessage = "Кто добавил выбрать обязательно")]
        public UserDTO AddUserDTOFK { get; set; }

        [Display(Name = "ИД тэга СИР")]
        [Required(ErrorMessage = "ИД тэга СИР обязателен")]
        public int MesParamId { get; set; }

        [Display(Name = "Тэг СИР")]
        [Required(ErrorMessage = "Тэг СИРа обязателен")]
        public MesParamDTO MesParamDTOFK { get; set; }

        [Display(Name = "Время значения")]
        [Required(ErrorMessage = "Время значения обязательно")]
        public DateTime ValueTime { get; set; }

        [Display(Name = "Значение")]
        [Required(ErrorMessage = "Значение обязательно")]
        public decimal Value { get; set; } = decimal.Zero;

        [Display(Name = "ИД записи в витрине SAP (выход)")]
        public Guid? SapMovementOutId { get; set; }

        [Display(Name = "Запись в витрине SAP (выход)")]
        public SapMovementsOUTDTO? SapMovementsOUTDTOFK { get; set; }

        [Display(Name = "ИД записи в витрине SAP (вход)")]
        public string? SapMovementInId { get; set; }

        [Display(Name = "Запись в витрине SAP (вход)")]
        public SapMovementsINDTO? SapMovementsINDTOFK { get; set; }

        [Display(Name = "ИД источника данных")]
        public int? DataSourceId { get; set; }

        [Display(Name = "Источник данных")]
        public DataSourceDTO? DataSourceDTOFK { get; set; }

        [Display(Name = "ИД типа данных")]
        public int? DataTypeId { get; set; }

        [Display(Name = "Тип данных")]
        public DataTypeDTO? DataTypeDTOFK { get; set; }

        [Display(Name = "ИД экземпляра отчёта")]
        public Guid? ReportGuid { get; set; }

        [Display(Name = "Экземпляр отчёта")]
        public ReportEntityDTO? ReportEntityDTOFK { get; set; }

        [Display(Name = "ИД предыдущей записи")]
        public Guid? PreviousRecordId { get; set; }

        [Display(Name = "Предыдущая запись")]
        public MesMovementsDTO? MesMovementsDTOFK { get; set; }

        [Display(Name = "Ушло из СИРа в MES")]
        public bool? MesGone { get; set; }

        [Display(Name = "Время ушло из СИРа в MES")]
        public DateTime? MesGoneTime { get; set; }

        [Display(Name = "Передавать в SAP")]
        public bool? NeedWriteToSap { get; set; }

        [Display(Name = "Комментарии")]
        public IEnumerable<MesMovementsCommentDTO>? MesMovementsCommentListDTO { get; set; }

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

        [NotMapped]
        [Display(Name = "Значение")]
        public string ToStringValue
        {
            get
            {
                return Value.ToString();
            }
            set
            {
                ToStringValue = value;
            }
        }

        [NotMapped]
        [Display(Name = "ИД записи в витрине SAP (выход)")]
        public string? ToStringSapMovementOutId
        {
            get
            {
                return SapMovementOutId.ToString();
            }
            set
            {
                ToStringSapMovementOutId = value;
            }
        }

        [NotMapped]
        [Display(Name = "ИД предыдущей записи")]
        public string? ToStringPreviousRecordId
        {
            get
            {
                return PreviousRecordId.ToString();
            }
            set
            {
                ToStringPreviousRecordId = value;
            }
        }

        [NotMapped]
        [Display(Name = "Причина корректировки")]
        public string ToStringCorrectionReason
        {
            get
            {
                MesMovementsCommentDTO? commentVar = MesMovementsCommentListDTO.FirstOrDefault();

                if (commentVar != null)
                {
                    var retVar = commentVar.CorrectionReasonDTOFK.Name;
                    if (retVar != null)
                        return retVar;
                    else
                        return "";
                }
                return "";
            }
            set
            {
                ToStringValue = value;
            }
        }

        [NotMapped]
        [Display(Name = "Комментарий корректировки")]
        public string ToStringCorrectionComment
        {
            get
            {
                MesMovementsCommentDTO? commentVar = MesMovementsCommentListDTO.FirstOrDefault();
                if (commentVar != null)
                {
                    var retVar = commentVar.CorrectionComment;
                    if (retVar != null)
                        return retVar;
                    else
                        return "";
                }
                else
                    return "";
            }
            set
            {
                ToStringValue = value;
            }
        }

        [NotMapped]
        [Display(Name = "Sap забрал")]
        public bool SapGoneBool
        {
            get
            {
                if (SapMovementsOUTDTOFK == null)
                {
                    return false;
                }
                else
                {
                    return SapMovementsOUTDTOFK.SapGone == null ? false : (bool)SapMovementsOUTDTOFK.SapGone;
                }
            }
            set
            {
                SapGoneBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Mes забрал")]
        public bool MesGoneBool
        {
            get
            {

                return MesGone == null ? false : (bool)MesGone;

            }
            set
            {
                MesGoneBool = value;
            }
        }

        [NotMapped]
        [Display(Name = "Передавать в SAP")]
        public bool NeedWriteToSapBool
        {
            get
            {
                return NeedWriteToSap == null ? false : (bool)NeedWriteToSap;
            }
            set
            {
                NeedWriteToSapBool = value;
            }
        }
    }
}

