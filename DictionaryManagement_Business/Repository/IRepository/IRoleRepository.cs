using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IRoleRepository
    {
        public Task<RoleDTO> GetById(Guid Id);
        public Task<IEnumerable<RoleDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<RoleDTO> Update(RoleDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<RoleDTO> Create(RoleDTO objectToAddDTO);
        public Task<RoleDTO> GetByName(string name = "");

    }
}
