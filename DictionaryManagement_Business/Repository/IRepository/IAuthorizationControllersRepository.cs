using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public interface IAuthorizationControllersRepository
    {
        public Task<AdminMode> CurrentUserIsInAdminRoleByLogin(string userLogin, MessageBoxMode messageBoxModePar = MessageBoxMode.Off);
    }
}
