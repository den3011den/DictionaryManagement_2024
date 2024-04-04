using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_Models.IntDBModels;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.IdentityModel.Tokens;
using Microsoft.JSInterop;
using System.DirectoryServices.AccountManagement;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class AuthorizationRepository : IAuthorizationRepository
    {

        private readonly AuthenticationStateProvider _authenticationStateProvider;

        private readonly IUserToRoleRepository _userToRoleRepository;
        private readonly IUserRepository _userRepository;
        private readonly IRoleToADGroupRepository _roleToADGroupRepository;
        private readonly ISettingsRepository _settingsRepository;
        private readonly IJSRuntime _jsRuntime;
        private readonly ILogEventRepository _logEventRepository;

        public AuthorizationRepository(AuthenticationStateProvider authenticationStateProvider, IUserToRoleRepository userToRoleRepository,
            IUserRepository userRepository, IJSRuntime jsRuntime,
            IRoleToADGroupRepository roleToADGroupRepository,
            ISettingsRepository settingsRepository, ILogEventRepository logEventRepository)
        {
            _authenticationStateProvider = authenticationStateProvider;
            _userToRoleRepository = userToRoleRepository;
            _userRepository = userRepository;
            _jsRuntime = jsRuntime;
            _roleToADGroupRepository = roleToADGroupRepository;
            _settingsRepository = settingsRepository;
            _logEventRepository = logEventRepository;
        }

        public async Task<AdminMode> CurrentUserIsInAdminRole(MessageBoxMode messageBoxModePar = SD.MessageBoxMode.Off)
        {
            AdminMode retVar = AdminMode.None;
            bool messShownFlag = false;

            if (_authenticationStateProvider is not null)
            {
                var authState = await _authenticationStateProvider.GetAuthenticationStateAsync();
                if (authState.User != null)
                {
                    if (authState.User.Identity is not null && authState.User.Identity.IsAuthenticated)
                    {
                        await SyncRolesByLoginWithADGroup(authState);
                        retVar = await _userToRoleRepository.IsUserInAdminRoleByUserLogin(authState.User.Identity.Name);
                    }
                }

                if (retVar == AdminMode.None)
                {
                    if (messageBoxModePar == MessageBoxMode.On)
                    {
                        messShownFlag = true;
                        await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Пользователь " + authState.User.Identity.Name +
                            " не найден, находится в архиве или не имеет роли с правами администрирования СИР. Обратитесь в техподдержку.");
                    }
                }
            }


            if (retVar == AdminMode.None && messShownFlag == false)
            {
                if (messageBoxModePar == MessageBoxMode.On)
                {
                    await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Не удалось получить логин текущего пользователя.");
                }
            }

            return retVar;
        }

        public async Task<UserDTO>? GetCurrentUserDTO(MessageBoxMode messageBoxModePar = MessageBoxMode.Off)
        {

            UserDTO retVar = null;

            bool messShownFlag = false;

            if (_authenticationStateProvider is not null)
            {
                var authState = await _authenticationStateProvider.GetAuthenticationStateAsync();
                if (authState.User != null)
                {
                    if (authState.User.Identity is not null && authState.User.Identity.IsAuthenticated)
                    {
                        retVar = await _userRepository.GetByLoginNotInArchive(authState.User.Identity.Name);
                    }
                }

                if (retVar == null)
                {
                    if (messageBoxModePar == MessageBoxMode.On)
                    {
                        messShownFlag = true;
                        await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Пользователь " + authState.User.Identity.Name +
                            " не найден в справочнике пользователей СИР");
                    }
                }
            }

            if (messShownFlag == false && retVar == null)
            {
                if (messageBoxModePar == MessageBoxMode.On)
                {
                    await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Не удалось получить логин текущего пользователя.");
                }

            }
            return retVar;
        }

        public async Task<string> GetCurrentUser(MessageBoxMode messageBoxModePar = MessageBoxMode.Off, LoginReturnMode loginReturnMode = LoginReturnMode.LoginOnly)
        {

            string? returnString = "";

            if (_authenticationStateProvider is not null)
            {
                var authState = await _authenticationStateProvider.GetAuthenticationStateAsync();
                if (authState.User != null)
                {
                    if (authState.User.Identity is not null && authState.User.Identity.IsAuthenticated)
                    {
                        string? loginStr = authState.User.Identity.Name;

                        if (loginReturnMode == LoginReturnMode.LoginOnly)
                            returnString = loginStr;
                        else
                        {
                            if (OperatingSystem.IsWindows())
                            {
                                try
                                {
                                    //#pragma warning disable CA1416 // Validate platform compatibility
                                    var context = new PrincipalContext(ContextType.Domain);
                                    var principal = UserPrincipal.FindByIdentity(context, authState.User.Identity.Name);
                                    //#pragma warning restore CA1416 // Validate platform compatibility

                                    var varString = principal.Name;

                                    if (varString.IsNullOrEmpty())
                                        varString = principal.DisplayName;
                                    if (varString.IsNullOrEmpty())
                                        varString = principal.Surname + " " + principal.GivenName + " ";

                                    if (loginReturnMode == LoginReturnMode.NameOnly || loginReturnMode == LoginReturnMode.LoginAndNameAndAccessMode)
                                    {
                                        returnString = varString;
                                    }
                                    if (loginReturnMode == LoginReturnMode.LoginAndName || loginReturnMode == LoginReturnMode.LoginAndNameAndAccessMode)
                                    {
                                        returnString = varString + " ( " + loginStr + " )";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    returnString = loginStr;
                                }
                            }
                            else
                            {
                                returnString = loginStr;
                            }
                        }

                        if (loginReturnMode == LoginReturnMode.LoginAndNameAndAccessMode)
                        {
                            AdminMode IsAdminMode = await CurrentUserIsInAdminRole(SD.MessageBoxMode.Off);
                            switch (IsAdminMode)
                            {
                                case AdminMode.IsAdmin: returnString = returnString + " - полный доступ"; break;
                                case AdminMode.IsAdminReadOnly: returnString = returnString + " - только чтение"; break;
                                default: returnString = returnString + " - НЕТ ДОСТУПА"; break;
                            }
                        }
                    }
                }
            }

            if (returnString.IsNullOrEmpty() && messageBoxModePar == MessageBoxMode.On)
            {
                await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Не удалось получить логин текущего пользователя.");
            }
            return returnString;
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
                    await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Пользователь " + userLogin +
                        " не найден, находится в архиве или не имеет роли с правами на администрирование СИР. Обратитесь в техподдержку.");
                }
            }

            if (retVar == AdminMode.None && messShownFlag == false)
            {
                if (messageBoxModePar == MessageBoxMode.On)
                {
                    await _jsRuntime.InvokeVoidAsync("ShowSwal", "error", "Не удалось получить логин текущего пользователя.");
                }
            }
            return retVar;
        }

        public async Task SyncRolesByLoginWithADGroup(AuthenticationState authStatePar)
        {
            Guid dictionaryManagementUserGuid = (await _userRepository.GetByUserName(SD.DictionaryManagementUserName)).Id;
            string userLogin = await GetCurrentUser(SD.MessageBoxMode.Off, SD.LoginReturnMode.LoginOnly);
            string userName = await GetCurrentUser(SD.MessageBoxMode.Off, SD.LoginReturnMode.NameOnly);

            if (!userLogin.IsNullOrEmpty())
            {

                UserDTO userFromDBDTO = await _userRepository.GetByLogin(userLogin);
                UserToRoleDTO userToRoleDTO = null;

                bool needAddUser = false;
                bool needCheckAddGroups = false;
                bool needCheckDeleteGroups = false;
                if (userFromDBDTO != null)
                {
                    needAddUser = false;
                    if (userFromDBDTO.IsArchive != true)
                    {
                        if (userFromDBDTO.IsSyncWithAD == true)
                            if (userFromDBDTO.SyncWithADGroupsLastTime != null)
                            {
                                {
                                    int syncMinutes = 0;
                                    try
                                    {
                                        syncMinutes = int.Parse((await _settingsRepository.GetByName(SD.SyncUserADGroupsIntervalInMinutesSettingName)).Value);
                                    }
                                    catch
                                    {
                                        syncMinutes = 0;
                                    }

                                    TimeSpan diff = DateTime.Now - ((DateTime)userFromDBDTO.SyncWithADGroupsLastTime);

                                    if (diff.TotalMinutes >= syncMinutes)
                                    {
                                        needCheckAddGroups = true;
                                        needCheckDeleteGroups = true;
                                    }
                                    else
                                    {
                                        needCheckAddGroups = false;
                                        needCheckDeleteGroups = false;
                                    }
                                }
                            }
                            else
                            {
                                needCheckAddGroups = true;
                                needCheckDeleteGroups = true;
                            }
                        else
                        {
                            needCheckAddGroups = false;
                            needCheckDeleteGroups = false;
                        }

                    }
                    else
                    {
                        needCheckAddGroups = false;
                        needCheckDeleteGroups = false;
                    }
                }
                else
                {
                    needAddUser = true;
                    needCheckAddGroups = true;
                    needCheckDeleteGroups = false;
                }

                if (needAddUser)
                {
                    userFromDBDTO = new UserDTO
                    {
                        UserName = userName,
                        Login = userLogin,
                        Description = "Добавлено автоматически " + DateTime.Now.ToString(),
                        IsArchive = false,
                        IsSyncWithAD = true,
                        IsServiceUser = false
                    };
                    userFromDBDTO = await _userRepository.Create(userFromDBDTO);
                    await _logEventRepository.AddRecord("Добавление пользователя", dictionaryManagementUserGuid, "", "", false, "Пользователь: " + userFromDBDTO.ToString());
                }

                IEnumerable<RoleToADGroupDTO> RoleToADGroupDTOList = null;
                if ((needCheckAddGroups == true) || (needCheckDeleteGroups == true))
                {
                    _jsRuntime.InvokeVoidAsync("ShowSwal", "loading", "Подождите. Выполняется синхронизация с AD ...");
                    RoleToADGroupDTOList = (await _roleToADGroupRepository.GetAll()).OrderBy(u => u.RoleId);
                }

                if (needCheckAddGroups == true)
                {
                    foreach (var item in RoleToADGroupDTOList)
                    {
                        if (item.ADGroupDTOFK.IsArchive != true)
                        {
                            userToRoleDTO = await _userToRoleRepository.Get(userFromDBDTO.Id, item.RoleId);
                            if (userToRoleDTO == null)
                            {
                                if (authStatePar.User.IsInRole(item.ADGroupDTOFK.Name.Trim()))
                                {
                                    {
                                        UserToRoleDTO userToRoleAddDTO = new UserToRoleDTO
                                        {
                                            UserId = userFromDBDTO.Id,
                                            RoleId = item.RoleId,
                                            UserDTOFK = userFromDBDTO,
                                            RoleDTOFK = item.RoleDTOFK
                                        };
                                        _userToRoleRepository.Create(userToRoleAddDTO);
                                        await _logEventRepository.AddRecord("Добавление связки пользователя с ролью", dictionaryManagementUserGuid,
                                            "<Пусто>", userToRoleAddDTO.RoleDTOFK.ToString() + " --> " + userToRoleAddDTO.UserDTOFK.ToString(), false,
                                            "Пользователь: " + userToRoleAddDTO.UserDTOFK.ToString() + " Роль: " + userToRoleAddDTO.RoleDTOFK.ToString());
                                    }
                                }
                            }

                        }
                    }
                }

                if (needCheckDeleteGroups == true)
                {
                    IEnumerable<UserToRoleDTO> rolesFormDBList = await _userToRoleRepository.GetAll();
                    List<RoleDTO> roledList = rolesFormDBList.Select(u => u.RoleDTOFK).Distinct().ToList();
                    IEnumerable<RoleToADGroupDTO> roleToADDelDTOList = null;
                    UserToRoleDTO foundUserToRoleDTO = null;
                    bool needDeleteUserInRoleFlag = false;
                    foreach (var itemRoleDTO in roledList)
                    {

                        needDeleteUserInRoleFlag = true;

                        if (itemRoleDTO.IsArchive != true)
                        {
                            foundUserToRoleDTO = await _userToRoleRepository.Get(userFromDBDTO.Id, itemRoleDTO.Id);

                            if (foundUserToRoleDTO != null)
                            {
                                roleToADDelDTOList = RoleToADGroupDTOList.Where(u => u.RoleId == itemRoleDTO.Id && u.ADGroupDTOFK.IsArchive != true);
                                if (roleToADDelDTOList.Count() > 0)
                                {
                                    foreach (var varitem in roleToADDelDTOList)
                                    {
                                        if (authStatePar.User.IsInRole(varitem.ADGroupDTOFK.Name.Trim()))
                                        {
                                            if (varitem.ADGroupDTOFK.IsArchive != true)
                                            {
                                                needDeleteUserInRoleFlag = false;
                                                break;
                                            }
                                        }
                                    }
                                }
                                else
                                    needDeleteUserInRoleFlag = true;
                            }
                            else
                            {
                                needDeleteUserInRoleFlag = false;
                            }
                        }
                        else
                        {
                            needDeleteUserInRoleFlag = true;
                        }

                        if (needDeleteUserInRoleFlag)
                        {
                            await _userToRoleRepository.DeleteByRoleIdAndUserId(itemRoleDTO.Id, userFromDBDTO.Id);
                            await _logEventRepository.AddRecord("Удаление связки пользователя с ролью", dictionaryManagementUserGuid,
                                userFromDBDTO.ToString() + " --> " + itemRoleDTO.ToString(),
                                "<Пусто>", false,
                                "Пользователь: " + userFromDBDTO.ToString() + " Роль: " + itemRoleDTO.ToString());
                        }
                    }
                }

                if ((needCheckAddGroups) || (needAddUser) || (needCheckDeleteGroups))
                {
                    if (userFromDBDTO != null)
                    {
                        string oldSyncTime = userFromDBDTO.SyncWithADGroupsLastTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                        userFromDBDTO.SyncWithADGroupsLastTime = DateTime.Now;
                        await _userRepository.Update(userFromDBDTO, UpdateMode.Update);
                        await _logEventRepository.AddRecord("Изменение пользователя", dictionaryManagementUserGuid, oldSyncTime,
                            userFromDBDTO.SyncWithADGroupsLastTime.ToString("dd.MM.yyyy HH:mm:ss.fff"), false,
                            "Пользователь: " + userFromDBDTO.ToString() + " Поле: Время посл. синх. AD");
                        await _jsRuntime.InvokeVoidAsync("CloseSwal");
                    }
                }
            }
        }
    }
}
