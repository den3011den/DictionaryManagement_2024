using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.Linq;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class RoleVMRepository : IRoleVMRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;
        private readonly IUserToRoleRepository _userToRoleRepository;
        private readonly IReportTemplateTypeToRoleRepository _reportTemplateTypeToRoleRepository;
        private readonly IRoleToADGroupRepository _roleToADGroupRepository;
        private readonly IRoleToDepartmentRepository _roleToDepartmentRepository;
        private readonly IMesDepartmentRepository _mesDepartmentRepository;

        public RoleVMRepository(IntDBApplicationDbContext db, IMapper mapper
            , IUserToRoleRepository userToRoleRepository
            , IReportTemplateTypeToRoleRepository reportTemplateTypeToRoleRepository
            , IRoleToADGroupRepository roleToADGroupRepository
            , IRoleToDepartmentRepository roleToDepartmentRepository
            , IMesDepartmentRepository mesDepartmentRepository)
        {
            _db = db;
            _mapper = mapper;
            _userToRoleRepository = userToRoleRepository;
            _reportTemplateTypeToRoleRepository = reportTemplateTypeToRoleRepository;
            _roleToADGroupRepository = roleToADGroupRepository;
            _roleToDepartmentRepository = roleToDepartmentRepository;
            _mesDepartmentRepository = mesDepartmentRepository;
        }

        public async Task<RoleVMDTO?> Create(RoleVMDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<RoleVMDTO, Role>(objectToAddDTO);
            var addedRole = _db.Role.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<Role, RoleVMDTO>(addedRole.Entity);
        }

        public async Task<RoleVMDTO?> Get(Guid Id)
        {

            RoleVMDTO RoleVMDTOToReturn = null;
            if (Id != null && Id != Guid.Empty)
            {
                var objToGet = _db.Role.FirstOrDefaultWithNoLock(u => (u.Id == Id));
                if (objToGet != null)
                {
                    var objUserToRole = _db.UserToRole.Where(u => u.RoleId == Id).Include("UserFK").Include("RoleFK").
                            OrderBy(u => u.UserFK.UserName).ToListWithNoLock();
                    var objReportTemplateTypeToRole = _db.ReportTemplateTypeTоRole.Where(u => u.RoleId == Id).Include("ReportTemplateTypeFK").Include("RoleFK")
                        .OrderBy(u => u.ReportTemplateTypeFK.Name).ToListWithNoLock();
                    var objRoleToADGroup = _db.RoleToADGroup.Where(u => u.RoleId == Id).Include("ADGroupFK").Include("RoleFK")
                        .OrderBy(u => u.ADGroupFK.Name).ToListWithNoLock();
                    var objRoleToDepartment = _db.RoleToDepartment.Where(u => u.RoleId == Id).Include("DepartmentFK").Include("RoleFK")
                        .OrderBy(u => u.DepartmentFK.ShortName).ToListWithNoLock();



                    RoleVMDTOToReturn = _mapper.Map<Role, RoleVMDTO>(objToGet);
                    if (objUserToRole != null)
                    {
                        RoleVMDTOToReturn.UserToRoleDTOs = _mapper.Map<IEnumerable<UserToRole>, IEnumerable<UserToRoleDTO>>(objUserToRole);
                    }

                    if (objReportTemplateTypeToRole != null)
                    {
                        RoleVMDTOToReturn.ReportTemplateTypeTоRoleDTOs = _mapper.Map<IEnumerable<ReportTemplateTypeTоRole>, IEnumerable<ReportTemplateTypeTоRoleDTO>>(objReportTemplateTypeToRole);
                    }

                    if (objRoleToADGroup != null)
                    {
                        RoleVMDTOToReturn.RoleToADGroupDTOs = _mapper.Map<IEnumerable<RoleToADGroup>, IEnumerable<RoleToADGroupDTO>>(objRoleToADGroup);
                    }

                    if (objRoleToDepartment != null)
                    {
                        RoleVMDTOToReturn.RoleToDepartmentDTOs = _mapper.Map<IEnumerable<RoleToDepartment>, IEnumerable<RoleToDepartmentDTO>>(objRoleToDepartment);
                    }

                }
            }
            return RoleVMDTOToReturn;
        }

        public async Task<IEnumerable<RoleVMDTO>> GetAll(SD.SelectDictionaryScope selectDictionaryScope = SD.SelectDictionaryScope.All)
        {
            IEnumerable<RoleDTO> roleListDTOs = null;
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
                roleListDTOs = _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                roleListDTOs = _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                roleListDTOs = _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.Where(u => u.IsArchive != true).ToListWithNoLock());

            List<RoleVMDTO> roleVMDTOs = new();
            foreach (var roleDTO in roleListDTOs)
            {

                var roleId = roleDTO.Id;
                var objUserToRole = _db.UserToRole.Where(u => u.RoleId == roleId).Include("UserFK").Include("RoleFK").
                    OrderBy(u => u.UserFK.UserName).ToListWithNoLock();
                var objReportTemplateTypeToRole = _db.ReportTemplateTypeTоRole.Where(u => u.RoleId == roleId).Include("ReportTemplateTypeFK").Include("RoleFK")
                    .OrderBy(u => u.ReportTemplateTypeFK.Name).ToListWithNoLock();
                var objRoleToADGroup = _db.RoleToADGroup.Where(u => u.RoleId == roleId).Include("ADGroupFK").Include("RoleFK").
                    OrderBy(u => u.ADGroupFK.Name).ToListWithNoLock();

                var objRoleToDepartment = _db.RoleToDepartment.Where(u => u.RoleId == roleId).Include("DepartmentFK").Include("RoleFK").
                        OrderBy(u => u.DepartmentFK.ShortName).ToListWithNoLock();

                var addRoleVMDTO = _mapper.Map<RoleDTO, RoleVMDTO>(roleDTO);
                if (objUserToRole != null)
                {
                    addRoleVMDTO.UserToRoleDTOs = _mapper.Map<IEnumerable<UserToRole>, IEnumerable<UserToRoleDTO>>(objUserToRole);
                }

                if (objReportTemplateTypeToRole != null)
                {
                    addRoleVMDTO.ReportTemplateTypeTоRoleDTOs = _mapper.Map<IEnumerable<ReportTemplateTypeTоRole>, IEnumerable<ReportTemplateTypeTоRoleDTO>>(objReportTemplateTypeToRole);
                }

                if (objRoleToADGroup != null)
                {
                    addRoleVMDTO.RoleToADGroupDTOs = _mapper.Map<IEnumerable<RoleToADGroup>, IEnumerable<RoleToADGroupDTO>>(objRoleToADGroup);
                }

                if (objRoleToDepartment != null)
                {
                    addRoleVMDTO.RoleToDepartmentDTOs = _mapper.Map<IEnumerable<RoleToDepartment>, IEnumerable<RoleToDepartmentDTO>>(objRoleToDepartment);
                }


                if (addRoleVMDTO != null)
                {
                    roleVMDTOs.Add(addRoleVMDTO);
                }
            }
            return roleVMDTOs;
        }

        public async Task<RoleVMDTO?> GetByName(string name)
        {

            RoleVMDTO RoleVMDTOToReturn = null;
            if (name.IsNullOrEmpty())
            {
                var objToGet = _db.Role.FirstOrDefaultWithNoLock(u => (u.Name.Trim().ToUpper() == name.Trim().ToUpper()));
                if (objToGet != null)
                {

                    var roleId = objToGet.Id;
                    var objUserToRole = _db.UserToRole.Where(u => u.RoleId == roleId).Include("UserFK").Include("RoleFK").
                            OrderBy(u => u.UserFK.UserName).ToListWithNoLock();
                    var objReportTemplateTypeToRole = _db.ReportTemplateTypeTоRole.Where(u => u.RoleId == roleId).Include("ReportTemplateTypeFK").Include("RoleFK")
                        .OrderBy(u => u.ReportTemplateTypeFK.Name).ToListWithNoLock();
                    var objRoleToADGroup = _db.RoleToADGroup.Where(u => u.RoleId == roleId).Include("ADGroupFK").Include("RoleFK").
                        OrderBy(u => u.ADGroupFK.Name).ToListWithNoLock();
                    var objRoleToDepartment = _db.RoleToDepartment.Where(u => u.RoleId == roleId).Include("DepartmentFK").Include("RoleFK").
                        OrderBy(u => u.DepartmentFK.ShortName).ToListWithNoLock();


                    RoleVMDTOToReturn = _mapper.Map<Role, RoleVMDTO>(objToGet);
                    if (objUserToRole != null)
                    {
                        RoleVMDTOToReturn.UserToRoleDTOs = _mapper.Map<IEnumerable<UserToRole>, IEnumerable<UserToRoleDTO>>(objUserToRole);
                    }

                    if (objReportTemplateTypeToRole != null)
                    {
                        RoleVMDTOToReturn.ReportTemplateTypeTоRoleDTOs = _mapper.Map<IEnumerable<ReportTemplateTypeTоRole>, IEnumerable<ReportTemplateTypeTоRoleDTO>>(objReportTemplateTypeToRole);

                    }

                    if (objRoleToADGroup != null)
                    {
                        RoleVMDTOToReturn.RoleToADGroupDTOs = _mapper.Map<IEnumerable<RoleToADGroup>, IEnumerable<RoleToADGroupDTO>>(objRoleToADGroup);
                    }

                    if (objRoleToDepartment != null)
                    {
                        RoleVMDTOToReturn.RoleToDepartmentDTOs = _mapper.Map<IEnumerable<RoleToDepartment>, IEnumerable<RoleToDepartmentDTO>>(objRoleToDepartment);
                    }

                }
            }
            return RoleVMDTOToReturn;
        }


        public async Task<RoleVMDTO?> Update(RoleVMDTO objectToUpdateVMDTO, SD.UpdateMode updateMode = SD.UpdateMode.Update)
        {
            var objectToUpdate = _db.Role.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateVMDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateVMDTO.Name)
                        objectToUpdate.Name = objectToUpdateVMDTO.Name;
                    if (objectToUpdate.Description != objectToUpdateVMDTO.Description)
                        objectToUpdate.Description = objectToUpdateVMDTO.Description;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.Role.Update(objectToUpdate);
                _db.SaveChanges();

                var retRoleVMDTO = await Get(objectToUpdateVMDTO.Id);
                return retRoleVMDTO;
            }
            return null;
        }

        public async Task<UserToRoleDTO?> AddUserToRole(RoleVMDTO roleVMDTO, UserDTO addUserDTO)
        {
            var checkUserToRole = _db.UserToRole.FirstOrDefaultWithNoLock(u => u.RoleId == roleVMDTO.Id && u.UserId == addUserDTO.Id);
            if (checkUserToRole != null)
            {
                // уже есть связка роли и пользователя
                return null;
            }

            UserToRole objectUserToRoleToAdd = new UserToRole();

            objectUserToRoleToAdd.UserId = addUserDTO.Id;
            objectUserToRoleToAdd.RoleId = roleVMDTO.Id;

            var addedUserToRole = _db.UserToRole.Add(objectUserToRoleToAdd);
            _db.SaveChanges();

            return _mapper.Map<UserToRole, UserToRoleDTO>(addedUserToRole.Entity);

        }

        public async Task<ReportTemplateTypeTоRoleDTO?> AddReportTemplateTypeToRole(RoleVMDTO roleVMDTO, ReportTemplateTypeDTO addreportTemplateTypeDTO)
        {
            var checkReportTemplateTypeTоRole = _db.ReportTemplateTypeTоRole.FirstOrDefaultWithNoLock(u => u.RoleId == roleVMDTO.Id && u.ReportTemplateTypeId == addreportTemplateTypeDTO.Id);
            if (checkReportTemplateTypeTоRole != null)
            {
                // уже есть связка роли и типа шаблона отчёта
                return null;
            }

            ReportTemplateTypeTоRole objectReportTemplateTypeTоRoleToAdd = new ReportTemplateTypeTоRole();

            objectReportTemplateTypeTоRoleToAdd.ReportTemplateTypeId = addreportTemplateTypeDTO.Id;
            objectReportTemplateTypeTоRoleToAdd.RoleId = roleVMDTO.Id;
            objectReportTemplateTypeTоRoleToAdd.CanDownload = addreportTemplateTypeDTO.CanDownload;
            objectReportTemplateTypeTоRoleToAdd.CanUpload = addreportTemplateTypeDTO.CanUpload;

            var addedReportTemplateTypeTоRole = _db.ReportTemplateTypeTоRole.Add(objectReportTemplateTypeTоRoleToAdd);
            _db.SaveChanges();

            return _mapper.Map<ReportTemplateTypeTоRole, ReportTemplateTypeTоRoleDTO>(addedReportTemplateTypeTоRole.Entity);

        }

        public async Task<RoleToDepartmentDTO?> AddRoleToDepartment(RoleVMDTO roleVMDTO, MesDepartmentDTO addDepartmentDTO)
        {
            var checkRoleToDepartment = _db.RoleToDepartment.FirstOrDefaultWithNoLock(u => u.RoleId == roleVMDTO.Id && u.DepartmentId == addDepartmentDTO.Id);
            if (checkRoleToDepartment != null)
            {
                // уже есть связка роли и производства
                return null;
            }

            RoleToDepartment objectRoleToDepartmentToAdd = new RoleToDepartment();

            objectRoleToDepartmentToAdd.DepartmentId = addDepartmentDTO.Id;
            objectRoleToDepartmentToAdd.RoleId = roleVMDTO.Id;

            var addedRoleToDepartment = _db.RoleToDepartment.Add(objectRoleToDepartmentToAdd);
            _db.SaveChanges();

            return _mapper.Map<RoleToDepartment, RoleToDepartmentDTO>(addedRoleToDepartment.Entity);

        }


        public async Task<RoleToADGroupDTO?> AddRoleToADGroup(RoleVMDTO roleVMDTO, ADGroupDTO addADGroupDTO)
        {
            var checkRoleToADGroup = _db.RoleToADGroup.FirstOrDefaultWithNoLock(u => u.RoleId == roleVMDTO.Id && u.ADGroupId == addADGroupDTO.Id);
            if (checkRoleToADGroup != null)
            {
                // уже есть связка роли и группы AD
                return null;
            }

            RoleToADGroup objectRoleToADGroupToAdd = new RoleToADGroup();

            objectRoleToADGroupToAdd.ADGroupId = addADGroupDTO.Id;
            objectRoleToADGroupToAdd.RoleId = roleVMDTO.Id;

            var addedRoleToADGroup = _db.RoleToADGroup.Add(objectRoleToADGroupToAdd);
            _db.SaveChanges();

            return _mapper.Map<RoleToADGroup, RoleToADGroupDTO>(addedRoleToADGroup.Entity);

        }



        public async Task<int> DeleteUserToRole(int userToRoleId)
        {
            if (userToRoleId > 0)
            {
                return await _userToRoleRepository.Delete(userToRoleId);
            }
            return 0;
        }

        public async Task<int> DeleteReportTemplateTypeToRole(int reportTemplateTypeToRoleId)
        {
            if (reportTemplateTypeToRoleId > 0)
            {
                return await _reportTemplateTypeToRoleRepository.Delete(reportTemplateTypeToRoleId);
            }
            return 0;
        }

        public async Task<int> DeleteRoleToDepartment(int roleToDepartmentId)
        {
            if (roleToDepartmentId > 0)
            {
                return await _roleToDepartmentRepository.Delete(roleToDepartmentId);
            }
            return 0;
        }

        public async Task<int> DeleteRoleToADGroup(int roleToADGroupId)
        {
            if (roleToADGroupId > 0)
            {
                return await _roleToADGroupRepository.Delete(roleToADGroupId);
            }
            return 0;
        }


        public async Task<IEnumerable<UserDTO>> GetAllNotArchiveAndNotAutomaticUsersExceptAlreadyInRole(Guid roleId)
        {
            IEnumerable<UserDTO> userListDTOs = null;
            IEnumerable<User> userInRoleDTOs = null;
            userInRoleDTOs = _db.UserToRole.Include("UserFK").Where(u => u.RoleId == roleId).Select(u => u.UserFK)
                .OrderBy(u => u.UserName).AsNoTracking().ToListWithNoLock();

            userListDTOs = _mapper.Map<IEnumerable<User>, IEnumerable<UserDTO>>
                (
                _db.User.Where(u => u.IsArchive != true && u.IsSyncWithAD != true)
                    .Where(u => !userInRoleDTOs.Contains(u))
                    .OrderBy(u => u.UserName).AsNoTracking().ToListWithNoLock()
                );
            return userListDTOs;
        }


        public async Task<IEnumerable<ReportTemplateTypeDTO>> GetAllNotArchiveReportTemplateTypesExceptAlreadyInRole(Guid roleId)
        {
            IEnumerable<ReportTemplateTypeDTO> reportTemplateTypeListDTOs = null;
            IEnumerable<ReportTemplateType> reportTemplateTypeInRoleDTOs = null;
            reportTemplateTypeInRoleDTOs = _db.ReportTemplateTypeTоRole.Include("ReportTemplateTypeFK").Where(u => u.RoleId == roleId).Select(u => u.ReportTemplateTypeFK)
                .OrderBy(u => u.Name).AsNoTracking().ToListWithNoLock();

            reportTemplateTypeListDTOs = _mapper.Map<IEnumerable<ReportTemplateType>, IEnumerable<ReportTemplateTypeDTO>>
                (
                _db.ReportTemplateType.Where(u => u.IsArchive != true)
                    .Where(u => !reportTemplateTypeInRoleDTOs.Contains(u))
                    .OrderBy(u => u.Name).AsNoTracking().ToListWithNoLock()
                );
            return reportTemplateTypeListDTOs;
        }
        public async Task<IEnumerable<ADGroupDTO>> GetAllNotArchiveADGroupsExceptAlreadyInRole(Guid roleId)
        {
            IEnumerable<ADGroupDTO> adGroupListDTOs = null;
            IEnumerable<ADGroup> adGroupsInRole = null;
            adGroupsInRole = _db.RoleToADGroup.Include("ADGroupFK").Where(u => u.RoleId == roleId).Select(u => u.ADGroupFK)
                .OrderBy(u => u.Name).AsNoTracking().ToListWithNoLock();

            adGroupListDTOs = _mapper.Map<IEnumerable<ADGroup>, IEnumerable<ADGroupDTO>>
                (
                _db.ADGroup.Where(u => u.IsArchive != true)
                    .Where(u => !adGroupsInRole.Contains(u)).OrderBy(u => u.Name).AsNoTracking().ToListWithNoLock()
                );
            return adGroupListDTOs;

        }

        public async Task<IEnumerable<UserToRoleDTO?>> GetUsersLinkedToRoleByRoleId(Guid roleId)
        {

            IEnumerable<UserToRoleDTO> userToRoleDTOs =
                _mapper.Map<IEnumerable<UserToRole>, IEnumerable<UserToRoleDTO>>
                (
                    _db.UserToRole.Include("UserFK").Include("RoleFK").Where(u => u.RoleId == roleId)
                        .OrderBy(u => u.UserFK.UserName).AsNoTracking().ToListWithNoLock()
                );

            return userToRoleDTOs;

        }

        public async Task<IEnumerable<ReportTemplateTypeTоRoleDTO>?> GetReportTemplateTypesLinkedToRoleByRoleId(Guid roleId)
        {
            IEnumerable<ReportTemplateTypeTоRoleDTO> reportTemplateTypeTоRoleDTOs =
                _mapper.Map<IEnumerable<ReportTemplateTypeTоRole>, IEnumerable<ReportTemplateTypeTоRoleDTO>>
                (
                    _db.ReportTemplateTypeTоRole.Include("ReportTemplateTypeFK").Include("RoleFK").Where(u => u.RoleId == roleId)
                        .OrderBy(u => u.ReportTemplateTypeFK.Name).AsNoTracking().ToListWithNoLock()
                );

            return reportTemplateTypeTоRoleDTOs;

        }

        public async Task<IEnumerable<RoleToADGroupDTO>?> GetADGroupsLinkedToRoleByRoleId(Guid roleId)
        {
            IEnumerable<RoleToADGroupDTO> roleToADGroupDTOs =
                _mapper.Map<IEnumerable<RoleToADGroup>, IEnumerable<RoleToADGroupDTO>>
                (
                    _db.RoleToADGroup.Include("ADGroupFK").Include("RoleFK").Where(u => u.RoleId == roleId)
                        .OrderBy(u => u.ADGroupFK.Name).AsNoTracking().ToListWithNoLock()
                );

            return roleToADGroupDTOs;

        }

        public async Task<IEnumerable<RoleToDepartmentDTO>?> GetDepartmentsLinkedToRoleByRoleId(Guid roleId)
        {
            IEnumerable<RoleToDepartmentDTO> roleToDepartmentDTOs =
                _mapper.Map<IEnumerable<RoleToDepartment>, IEnumerable<RoleToDepartmentDTO>>
                (
                    _db.RoleToDepartment.Include("DepartmentFK").Include("RoleFK").Where(u => u.RoleId == roleId)
                        .OrderBy(u => u.DepartmentFK.ShortName).AsNoTracking().ToListWithNoLock()
                );

            return roleToDepartmentDTOs;

        }

        public async Task<IEnumerable<MesDepartmentVMDTO>> GetAllDepartmentWithChildrenCheckedWithLinkRole(Guid roleId, int? mesDepartmentRootId, MesDepartmentVMDTO? parentDepartmentVMDTO)
        {

            List<MesDepartmentVMDTO>? resutlList = new List<MesDepartmentVMDTO>();

            IEnumerable<MesDepartmentVMDTO> topLevelList = _mapper.Map<IEnumerable<MesDepartmentDTO>, IEnumerable<MesDepartmentVMDTO>>(await _mesDepartmentRepository.GetChildList(mesDepartmentRootId));
            //IEnumerable<MesDepartmentVMDTO>? childList = null;

            if (topLevelList != null)
            {
                foreach (var topLevelItem in topLevelList)
                {

                    MesDepartmentVMDTO mesDepartmentVMDTO = new MesDepartmentVMDTO();
                    mesDepartmentVMDTO.Id = topLevelItem.Id;
                    mesDepartmentVMDTO.MesCode = topLevelItem.MesCode;
                    mesDepartmentVMDTO.Name = topLevelItem.Name;
                    mesDepartmentVMDTO.ShortName = topLevelItem.ShortName;
                    mesDepartmentVMDTO.ParentDepartmentId = parentDepartmentVMDTO == null ? null : parentDepartmentVMDTO.ParentDepartmentId;
                    mesDepartmentVMDTO.DepartmentParentVMDTO = parentDepartmentVMDTO;
                    mesDepartmentVMDTO.IsArchive = topLevelItem.IsArchive;
                    mesDepartmentVMDTO.ToStringValue = topLevelItem.ToStringValue;

                    var foundRoleToDepartmentDTO = await _roleToDepartmentRepository.Get(roleId, topLevelItem.Id);
                    if (foundRoleToDepartmentDTO != null)
                    {
                        mesDepartmentVMDTO.Checked = true;
                    }
                    else
                    {
                        mesDepartmentVMDTO.Checked = false;
                    }
                    mesDepartmentVMDTO.ChildrenDTO = (await GetAllDepartmentWithChildrenCheckedWithLinkRole(roleId, topLevelItem.Id, mesDepartmentVMDTO));
                    resutlList.Add(mesDepartmentVMDTO);

                }
                return resutlList;
            }
            else
            {
                return resutlList;
            }

        }




        public async Task<IEnumerable<object>> GetAllDepartmentCheckedObjects(IEnumerable<MesDepartmentVMDTO> topLevelList)
        {

            List<object>? resutlList = new List<object>();

            if (topLevelList != null)
            {
                foreach (var topLevelItem in topLevelList)
                {
                    resutlList.AddRange(await GetAllDepartmentCheckedObjects(topLevelItem.ChildrenDTO));
                    if (topLevelItem.Checked == true)
                    {
                        resutlList.Add(topLevelItem);
                    }
                }
                return resutlList;
            }
            else
            {
                return resutlList;
            }

        }


        //public async Task<int> AddDepartmentsToRole(IEnumerable<MesDepartmentVMDTO> topLevelList, RoleVMDTO roleVMDTO)
        //{
        //    int retVar = 0;
        //    List<object>? resutlList = new List<object>();

        //    RoleToDepartmentDTO foundRoleToDepartmentDTO;

        //    if (topLevelList != null)
        //    {
        //        foreach (var topLevelItem in topLevelList)
        //        {
        //            retVar = retVar + (await AddDepartmentsToRole(topLevelItem.ChildrenDTO, roleVMDTO));

        //            foundRoleToDepartmentDTO = await _roleToDepartmentRepository.Get(roleVMDTO.Id, topLevelItem.Id);
        //            if (foundRoleToDepartmentDTO == null)
        //            {
        //                    RoleToDepartmentDTO newRoleToDepartmentDTO = new RoleToDepartmentDTO
        //                    {
        //                        RoleId = roleVMDTO.Id,
        //                        DepartmentId = topLevelItem.Id,
        //                        RoleDTOFK = _mapper.Map<RoleVMDTO,RoleDTO>(roleVMDTO),
        //                        DepartmentDTOFK = _mapper.Map<MesDepartmentVMDTO, MesDepartmentDTO>(topLevelItem)

        //                    };

        //                if ((await _roleToDepartmentRepository.Create(newRoleToDepartmentDTO)) != null)
        //                    retVar = retVar + 1;                            
        //            }
        //        }
        //        return retVar;
        //    }
        //    else
        //    {
        //        return retVar;
        //    }
        //}


        public async Task DeleteAllLikedDepartmentsToRoleByRoleId(Guid roleId)
        {

            IEnumerable<RoleToDepartment> roleToDepartmentList = _db.RoleToDepartment.Where(u => u.RoleId == roleId).ToListWithNoLock();
            _db.RoleToDepartment.RemoveRange(roleToDepartmentList);
            _db.SaveChanges();
        }




        public async Task<int> AddDepartmentsToRole(IEnumerable<object> objectList, RoleVMDTO roleVMDTO)
        {
            int retVar = 0;

            await DeleteAllLikedDepartmentsToRoleByRoleId(roleVMDTO.Id);
            retVar = 0;

            RoleToDepartmentDTO foundRoleToDepartmentDTO;

            if (objectList != null)
            {
                foreach (MesDepartmentVMDTO objectItem in objectList)
                {
                    {
                        RoleToDepartmentDTO newRoleToDepartmentDTO = new RoleToDepartmentDTO
                        {
                            RoleId = roleVMDTO.Id,
                            DepartmentId = objectItem.Id,
                            RoleDTOFK = _mapper.Map<RoleVMDTO, RoleDTO>(roleVMDTO),
                            DepartmentDTOFK = _mapper.Map<MesDepartmentVMDTO, MesDepartmentDTO>(objectItem)

                        };

                        if ((await _roleToDepartmentRepository.Create(newRoleToDepartmentDTO)) != null)
                            retVar = retVar + 1;
                    }
                }

            }
            return retVar;
        }
    }
}
