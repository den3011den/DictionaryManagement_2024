using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class SapMaterialDTO : MaterialDTO
    {

        public SapMaterialDTO()
        {

        }
        public SapMaterialDTO(MaterialDTO obj)
        {
            Id = obj.Id;
            Code = obj.Code;
            Name = obj.Name;
            ShortName = obj.ShortName;
            IsArchive = obj.IsArchive;
        }

        [NotMapped]
        public string ToStringValue { get; set; } = string.Empty;

        public override string ToString()
        {
            ToStringValue = $"{Code} {ShortName}";
            return ToStringValue;
        }

        [NotMapped]
        public string ToStringCodeName
        {
            get
            {
                string retVar = Code + " " + ShortName;
                return retVar;
            }
            set
            {
                ToStringCodeName = value;
            }
        }

    }
}
