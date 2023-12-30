using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DictionaryManagement_Models.IntDBModels
{
    public class MesDepartmentDTO
    {

        [ForLogAttribute(NameProperty = "поле \"ИД\"")]
        [Display(Name = "Ид записи")]
        [Required(ErrorMessage = "ИД обязателен")]
        public int Id { get; set; }

        //[Required(ErrorMessage = "Код обязателен для заполнения")]
        //[Range(1, int.MaxValue, ErrorMessage = "Код может быть от {1} до {2}")]
        [ForLogAttribute(NameProperty = "поле \"Код\"")]
        [Display(Name = "Код")]
        public int? MesCode { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Наименование\"")]
        [Required(ErrorMessage = "Наименование обязательно для заполнения")]
        [Display(Name = "Наименование")]
        [MaxLength(500, ErrorMessage = "Длина наименования не может быть больше 500 символов")]
        public string Name { get; set; } = string.Empty;

        [ForLogAttribute(NameProperty = "поле \"Сокр. наименование\"")]
        [Required(ErrorMessage = "Сокр. Наименование обязательно для заполнения")]
        [Display(Name = "Сокр. наименование")]
        [MaxLength(500, ErrorMessage = "Длина сокр. наименование не может быть больше 500 символов")]
        public string ShortName { get; set; } = string.Empty;

        [Display(Name = "Ид родителя")]
        public int? ParentDepartmentId { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Родитель\"")]
        [Display(Name = "Родитель")]
        public MesDepartmentDTO? DepartmentParentDTO { get; set; }

        [ForLogAttribute(NameProperty = "поле \"Архив\"")]
        [Display(Name = "В архиве")]
        public bool IsArchive { get; set; }

        //[NotMapped]
        //public string ToStringValue { get; set; } = string.Empty;

        //public override string ToString()
        //{
        //    ToStringValue = $"{ShortName}";
        //    return ToStringValue;
        //}

        [NotMapped]
        public string ToStringShortName
        {
            get
            {
                return ShortName;
            }
            set
            {
                ToStringShortName = value;
            }
        }

        [NotMapped]
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
        public string ToStringMesCode
        {
            get
            {
                return (MesCode == null ? "" : MesCode.ToString());
            }
            set
            {
                ToStringMesCode = value;
            }
        }

        [NotMapped]
        [Display(Name = " ")]
        public bool Checked { get; set; } = false;

        [NotMapped]
        [Display(Name = "Уровень")]
        public int DepLevel { get; set; }

        [NotMapped]
        public string ToStringHierarchyShortName
        {
            //get
            //{
            //    string ret_var = ShortName;
            //    MesDepartmentDTO mesDepartmentDTO = this;
            //    while(mesDepartmentDTO.DepartmentParentDTO != null)
            //    {
            //        ret_var = mesDepartmentDTO.DepartmentParentDTO.ShortName + " - " + ret_var;
            //        mesDepartmentDTO = mesDepartmentDTO.DepartmentParentDTO;
            //    }
            //    return ret_var;
            //}
            //set
            //{
            //    ToStringShortName = value;
            //}

            get
            {
                string ret_var = "";
                MesDepartmentDTO mesDepartmentDTO = this;
                while (mesDepartmentDTO.DepartmentParentDTO != null)
                {
                    if (string.IsNullOrEmpty(ret_var))
                        ret_var = mesDepartmentDTO.ShortName;
                    else
                        ret_var = mesDepartmentDTO.ShortName + " - " + ret_var;
                    mesDepartmentDTO = mesDepartmentDTO.DepartmentParentDTO;
                }
                if (string.IsNullOrEmpty(ret_var))
                    ret_var = mesDepartmentDTO.ShortName;
                else
                    ret_var = mesDepartmentDTO.ShortName + " - " + ret_var;
                return ret_var;
            }
            set
            {
                ToStringHierarchyShortName = value;
            }

        }
        public override string ToString()
        {
            return $"{ToStringHierarchyShortName}";
        }



    }
}

