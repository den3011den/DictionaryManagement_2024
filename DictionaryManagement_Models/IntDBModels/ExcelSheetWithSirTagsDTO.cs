using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class ExcelSheetWithSirTagsDTO
    {
        public string Column1 { get; set; }
        public string Column2 { get; set; }
        public string Column3 { get; set; }
        public string Column4 { get; set; }
        public string Column5 { get; set; }
        public string Column6 { get; set; }
        public string Column7 { get; set; }
        public string Column8 { get; set; }
        public string Column9 { get; set; }
        public string Column10 { get; set; }
        public string Column11 { get; set; }
        public string Column12 { get; set; }
        public MesParamDTO? MesParamDTOFK { get; set; }

        public bool MesParamFoundFlag { get; set; }

        [NotMapped]
        public string MesParamIdToString
        {
            get
            {
                return MesParamDTOFK == null ? "" : MesParamDTOFK.ToStringId;
            }
            set
            {
                MesParamIdToString = value;
            }
        }

        [NotMapped]
        public string MesParamCodeNameToString
        {
            get
            {
                string retVar = "НЕ НАЙДЕН";
                if (MesParamDTOFK != null)
                {
                    retVar = MesParamDTOFK.Code + " " + MesParamDTOFK.Name;
                }
                return retVar;
            }
            set
            {
                MesParamIdToString = value;
            }
        }

    }
}
