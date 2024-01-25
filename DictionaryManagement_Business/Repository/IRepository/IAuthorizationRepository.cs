using DictionaryManagement_Common;
using DictionaryManagement_Models.IntDBModels;
using Microsoft.AspNetCore.Components.Authorization;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public interface IAuthorizationRepository
    {
        public Task<AdminMode> CurrentUserIsInAdminRole(MessageBoxMode messageBoxModePar = SD.MessageBoxMode.Off);
        public Task<UserDTO>? GetCurrentUserDTO(MessageBoxMode messageBoxModePar = MessageBoxMode.Off);
        public Task<string> GetCurrentUser(MessageBoxMode messageBoxModePar = MessageBoxMode.Off, LoginReturnMode loginReturnMode = LoginReturnMode.LoginOnly);
        public Task<AdminMode> CurrentUserIsInAdminRoleByLogin(string userLogin, MessageBoxMode messageBoxModePar = MessageBoxMode.Off);
        public Task SyncRolesByLoginWithADGroup(AuthenticationState authStatePar);
    }
}
