using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IMesDepartmentRepository
    {
        public Task<MesDepartmentDTO> GetById(int mesDepartmentId);
        public Task<IEnumerable<MesDepartmentDTO>> GetChildList(int? mesDepartmentId);
        public Task<bool> HasChild(int mesDepartmentId);
        //public bool HasChild(int mesDepartmentId);
        public Task<IEnumerable<MesDepartmentDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<IEnumerable<MesDepartmentDTO>> GetAllTopLevel();
        public Task<MesDepartmentDTO> Update(MesDepartmentDTO objDTO);
        public Task<MesDepartmentDTO> Create(MesDepartmentDTO objectToAddDTO);
        public Task<int> Delete(int mesDepartmentId, UpdateMode updateMode = UpdateMode.Update);
        public Task<MesDepartmentDTO> GetByCode(int? mesCode = 0);
        public Task<MesDepartmentDTO> GetByName(string name = "");
        public Task<MesDepartmentDTO> GetByShortName(string shortName = "");
        public Task<Tuple<IEnumerable<MesDepartmentVMDTO>, int>> GetAllDepartmentWithChildren(int? mesDepartmentRootId, int level, int maxLevel, MesDepartmentVMDTO? departmentParentVMDTO);
    }
}
