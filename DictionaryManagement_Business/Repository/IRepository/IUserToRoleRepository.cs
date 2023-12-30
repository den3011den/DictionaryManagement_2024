using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IUserToRoleRepository
    {
        public Task<UserToRoleDTO> Get(Guid userId, Guid roleId);
        public Task<int> DeleteByRoleIdAndUserId(Guid roleId, Guid userId);
        public Task<UserToRoleDTO> GetById(int id);
        public Task<IEnumerable<UserToRoleDTO>> GetAll();
        public Task<UserToRoleDTO> Update(UserToRoleDTO objDTO);
        public Task<UserToRoleDTO> Create(UserToRoleDTO objectToAddDTO);
        public Task<int> Delete(int Id);
        public Task<bool> IsUserInRoleByUserLoginAndRoleName(string userLogin, string roleName);
    }
}
