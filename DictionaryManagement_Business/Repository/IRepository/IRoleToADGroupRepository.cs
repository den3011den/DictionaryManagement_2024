using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IRoleToADGroupRepository
    {
        public Task<RoleToADGroupDTO> Get(Guid adGroupId, Guid roleId);
        public Task<RoleToADGroupDTO> GetById(int id);
        public Task<IEnumerable<RoleToADGroupDTO>> GetAll();
        public Task<RoleToADGroupDTO> Update(RoleToADGroupDTO objDTO);
        public Task<RoleToADGroupDTO> Create(RoleToADGroupDTO objectToAddDTO);
        public Task<IEnumerable<RoleToADGroupDTO>> GetByRoleId(Guid roleId);
        public Task<int> Delete(int Id);
    }
}
