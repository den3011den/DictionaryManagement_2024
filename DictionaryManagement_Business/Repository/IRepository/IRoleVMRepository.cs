using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IRoleVMRepository
    {
        public Task<RoleVMDTO?> Get(Guid Id);
        public Task<IEnumerable<RoleVMDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<RoleVMDTO> Update(RoleVMDTO? objVMDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<RoleVMDTO?> Create(RoleVMDTO objectToAddDTO);
        public Task<RoleVMDTO?> GetByName(string name = "");

        public Task<UserToRoleDTO?> AddUserToRole(RoleVMDTO roleVMDTO, UserDTO addUserDTO);
        public Task<ReportTemplateTypeTоRoleDTO?> AddReportTemplateTypeToRole(RoleVMDTO roleVMDTO, ReportTemplateTypeDTO addreportTemplateTypeDTO);
        public Task<RoleToADGroupDTO?> AddRoleToADGroup(RoleVMDTO roleVMDTO, ADGroupDTO addADGroupDTO);

        public Task<int> DeleteUserToRole(int userToRoleId);
        public Task<int> DeleteReportTemplateTypeToRole(int reportTemplateTypeToRoleId);
        public Task<int> DeleteRoleToADGroup(int roleToAdGroupId);
        public Task<IEnumerable<UserDTO>> GetAllNotArchiveAndNotAutomaticAndNotServiceUsersExceptAlreadyInRole(Guid roleId);
        public Task<IEnumerable<ReportTemplateTypeDTO>> GetAllNotArchiveReportTemplateTypesExceptAlreadyInRole(Guid roleId);
        public Task<IEnumerable<ADGroupDTO>> GetAllNotArchiveADGroupsExceptAlreadyInRole(Guid roleId);

        public Task<IEnumerable<UserToRoleDTO>?> GetUsersLinkedToRoleByRoleId(Guid roleId);
        public Task<IEnumerable<ReportTemplateTypeTоRoleDTO>?> GetReportTemplateTypesLinkedToRoleByRoleId(Guid roleId);
        public Task<IEnumerable<RoleToADGroupDTO>?> GetADGroupsLinkedToRoleByRoleId(Guid roleId);
        public Task<IEnumerable<MesDepartmentVMDTO>> GetAllDepartmentWithChildrenCheckedWithLinkRole(Guid roleId, int? mesDepartmentRootId, MesDepartmentVMDTO? parentDepartmentVMDTO);
        public Task<IEnumerable<object>> GetAllDepartmentCheckedObjects(IEnumerable<MesDepartmentVMDTO> topLevelList);

        public Task<IEnumerable<RoleToDepartmentDTO>?> GetDepartmentsLinkedToRoleByRoleId(Guid roleId);

        public Task<int> DeleteRoleToDepartment(int roleToDepartmentId);

        public Task DeleteAllLikedDepartmentsToRoleByRoleId(Guid roleId);
        public Task<int> AddDepartmentsToRole(IEnumerable<object> objectList, RoleVMDTO roleVMDTO);
        public Task<RoleToDepartmentDTO?> AddRoleToDepartment(RoleVMDTO roleVMDTO, MesDepartmentDTO addDepartmentDTO);
    }
}
