using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using Microsoft.IdentityModel.Tokens;
using Microsoft.JSInterop;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class AuthorizationControllersRepository : IAuthorizationControllersRepository
    {
        private readonly IUserToRoleRepository _userToRoleRepository;
        private readonly IUserRepository _userRepository;
        private readonly IJSRuntime _jsRuntime;

        public AuthorizationControllersRepository(IUserToRoleRepository userToRoleRepository, IUserRepository userRepository, IJSRuntime jsRuntime)
        {
            _userToRoleRepository = userToRoleRepository;
            _userRepository = userRepository;
            _jsRuntime = jsRuntime;
        }


        public async Task<AdminMode> CurrentUserIsInAdminRoleByLogin(string userLogin, MessageBoxMode messageBoxModePar = SD.MessageBoxMode.Off)
        {
            AdminMode retVar = AdminMode.None;
            bool messShownFlag = false;

            if (!userLogin.IsNullOrEmpty())
            {
                retVar = await _userToRoleRepository.IsUserInAdminRoleByUserLogin(userLogin);
            }

            if (retVar == AdminMode.None)
            {
                if (messageBoxModePar == MessageBoxMode.On)
                {
                    messShownFlag = true;
                    await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Пользователь " + userLogin +
                        " не найден, находится в архиве или не имеет роли с правами администрирования СИР. Обратитесь в техподдержку.");
                }
            }

            if (retVar == AdminMode.None && messShownFlag == false)
            {
                if (messageBoxModePar == MessageBoxMode.On)
                {
                    messShownFlag = true;
                    await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Не удалось получить логин текущего пользователя.");
                }
            }
            return retVar;
        }
    }
}
