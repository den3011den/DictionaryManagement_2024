using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IUserRepository
    {
        public Task<UserDTO> Get(Guid Id);
        public Task<IEnumerable<UserDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<UserDTO> Update(UserDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<UserDTO> Create(UserDTO objectToAddDTO);

        public Task<UserDTO> GetByLogin(string login = "");
        public Task<UserDTO> GetByLoginNotInArchive(string login = "");
        public Task<UserDTO> GetByUserName(string userName = "");

    }
}
