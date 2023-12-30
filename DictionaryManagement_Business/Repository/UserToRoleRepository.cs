using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;


namespace DictionaryManagement_Business.Repository
{
    public class UserToRoleRepository : IUserToRoleRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public UserToRoleRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<UserToRoleDTO> Create(UserToRoleDTO objectToAddDTO)
        {

            UserToRole objectToAdd = new UserToRole();

            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.UserId = objectToAddDTO.UserId;
            objectToAdd.RoleId = objectToAddDTO.RoleId;


            var addedUserToRole = _db.UserToRole.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<UserToRole, UserToRoleDTO>(addedUserToRole.Entity);
        }

        public async Task<UserToRoleDTO> Get(Guid userId, Guid roleId)
        {
            var objToGet = _db.UserToRole.Include("UserFK").Include("RoleFK").
                            FirstOrDefault(u => u.UserId == userId && u.RoleId == roleId);
            if (objToGet != null)
            {
                return _mapper.Map<UserToRole, UserToRoleDTO>(objToGet);
            }
            return null;
        }

        public async Task<UserToRoleDTO> GetById(int id)
        {
            var objToGet = _db.UserToRole.Include("UserFK").Include("RoleFK").
                            FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<UserToRole, UserToRoleDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<UserToRoleDTO>> GetAll()
        {
            IEnumerable<UserToRole> hhh = _db.UserToRole.Include("UserFK").Include("RoleFK").AsNoTracking().ToListWithNoLock();
            return _mapper.Map<IEnumerable<UserToRole>, IEnumerable<UserToRoleDTO>>(hhh);


        }

        public async Task<UserToRoleDTO> Update(UserToRoleDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.UserToRole.Include("UserFK").Include("RoleFK").
                    FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {

                if (objectToUpdate.UserId != objectToUpdateDTO.UserDTOFK.Id)
                {
                    objectToUpdate.UserId = objectToUpdateDTO.UserDTOFK.Id;
                    objectToUpdate.UserFK = _mapper.Map<UserDTO, User>(objectToUpdateDTO.UserDTOFK);
                }
                if (objectToUpdate.RoleId != objectToUpdateDTO.RoleDTOFK.Id)
                {
                    objectToUpdate.RoleId = objectToUpdateDTO.RoleDTOFK.Id;
                    objectToUpdate.RoleFK = _mapper.Map<RoleDTO, Role>(objectToUpdateDTO.RoleDTOFK);
                }
                _db.UserToRole.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<UserToRole, UserToRoleDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id)
        {
            if (id > 0)
            {
                var objectToDelete = _db.UserToRole.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.UserToRole.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }


        public async Task<int> DeleteByRoleIdAndUserId(Guid roleId, Guid userId)
        {
            if (roleId != Guid.Empty && userId != Guid.Empty)
            {
                var listToDelete = _db.UserToRole.Where(u => u.RoleId == roleId && u.UserId == userId).ToListWithNoLock();
                if (listToDelete != null)
                {
                    foreach (var item in listToDelete)
                        _db.UserToRole.Remove(item);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }

        public async Task<bool> IsUserInRoleByUserLoginAndRoleName(string userLogin, string roleName)
        {
            //var objToGet = await _db.UserToRole.Include("UserFK").Include("RoleFK").
            //    Where(u => u.UserFK != null && u.RoleFK != null
            //        && u.UserFK.Login.Trim().ToUpper() == userLogin.Trim().ToUpper()
            //        && u.UserFK.IsArchive != true
            //        && u.RoleFK.Name.Trim().ToUpper() == roleName.Trim().ToUpper()
            //        && u.RoleFK.IsArchive != true).AsNoTracking().FirstOrDefaultAsync();                    

            var objToGet = _db.UserToRole.Include("UserFK").Include("RoleFK").
                Where(u => u.UserFK != null && u.RoleFK != null
                    && u.UserFK.Login.Trim().ToUpper() == userLogin.Trim().ToUpper()
                    && u.UserFK.IsArchive != true
                    && u.RoleFK.Name.Trim().ToUpper() == roleName.Trim().ToUpper()
                    && u.RoleFK.IsArchive != true).AsNoTracking().FirstOrDefaultWithNoLock();

            if (objToGet == null)
                return false;
            else
                return true;
        }
    }
}
