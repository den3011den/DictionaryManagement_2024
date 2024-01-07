using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesMaterialDTO : MaterialDTO
    {
        public MesMaterialDTO()
        {

        }

        public MesMaterialDTO(MaterialDTO obj)
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
        //public string ToStringForLog()
        //{
        //    return $"{Code} {ShortName}";
        //}
    }
}
