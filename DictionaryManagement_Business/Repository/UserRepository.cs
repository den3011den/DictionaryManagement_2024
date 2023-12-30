using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class UserRepository : IUserRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public UserRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<UserDTO> Create(UserDTO objectToAddDTO)
        {

            objectToAddDTO.UserName = objectToAddDTO.UserName.Replace("\\", "_").Replace("/", "_"); ;
            var objectToAdd = _mapper.Map<UserDTO, User>(objectToAddDTO);
            var addedUser = _db.User.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<User, UserDTO>(addedUser.Entity);
        }

        public async Task<UserDTO> Get(Guid Id)
        {
            if (Id != null && Id != Guid.Empty)
            {
                var objToGet = _db.User.FirstOrDefaultWithNoLock(u => u.Id == Id);
                if (objToGet != null)
                {
                    if (objToGet.SyncWithADGroupsLastTime == null)
                        objToGet.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;
                    return _mapper.Map<User, UserDTO>(objToGet);
                }
            }
            return null;
        }

        public async Task<IEnumerable<UserDTO>> GetAll(SD.SelectDictionaryScope selectDictionaryScope = SD.SelectDictionaryScope.All)
        {

            IEnumerable<UserDTO> hhh = null;
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
                hhh = _mapper.Map<IEnumerable<User>, IEnumerable<UserDTO>>(_db.User.ToListWithNoLock());

            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                hhh = _mapper.Map<IEnumerable<User>, IEnumerable<UserDTO>>(_db.User.Where(u => u.IsArchive == true).ToListWithNoLock());

            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                hhh = _mapper.Map<IEnumerable<User>, IEnumerable<UserDTO>>(_db.User.Where(u => u.IsArchive != true).ToListWithNoLock());

            //           hhh = _mapper.Map<IEnumerable<User>, IEnumerable<UserDTO>>(_db.User.ToListWithNoLock());

            hhh.ToList().FindAll(s => s.SyncWithADGroupsLastTime == DateTime.MinValue)
                .ForEach(x => x.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue);

            return hhh;

        }

        public async Task<UserDTO> GetByLogin(string login = "")
        {
            var objToGet = _db.User.FirstOrDefaultWithNoLock(u => ((u.Login.Trim().ToUpper() == login.Trim().ToUpper())));
            if (objToGet != null)
            {
                if (objToGet.SyncWithADGroupsLastTime == null)
                    objToGet.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

                return _mapper.Map<User, UserDTO>(objToGet);
            }
            return null;
        }

        public async Task<UserDTO> GetByLoginNotInArchive(string login = "")
        {
            var objToGet = _db.User.FirstOrDefaultWithNoLock(u => ((u.Login.Trim().ToUpper() == login.Trim().ToUpper()))
                && u.IsArchive != true);
            if (objToGet != null)
            {
                if (objToGet.SyncWithADGroupsLastTime == null)
                    objToGet.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

                return _mapper.Map<User, UserDTO>(objToGet);
            }
            return null;
        }

        public async Task<UserDTO> GetByUserName(string userName = "")
        {
            var objToGet = _db.User.FirstOrDefaultWithNoLock(u => ((u.UserName.Trim().ToUpper()) == (userName.Trim().ToUpper())));
            if (objToGet != null)
            {
                if (objToGet.SyncWithADGroupsLastTime == null)
                    objToGet.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

                return _mapper.Map<User, UserDTO>(objToGet);
            }
            return null;
        }


        public async Task<UserDTO> Update(UserDTO objectToUpdateDTO, SD.UpdateMode updateMode = SD.UpdateMode.Update)
        {
            var objectToUpdate = _db.User.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    objectToUpdateDTO.UserName = objectToUpdateDTO.UserName.Replace("\\", "_").Replace("/", "_");
                    if (objectToUpdate.Login != objectToUpdateDTO.Login)
                        objectToUpdate.Login = objectToUpdateDTO.Login;
                    if (objectToUpdate.UserName != objectToUpdateDTO.UserName)
                        objectToUpdate.UserName = objectToUpdateDTO.UserName;
                    if (objectToUpdate.Description != objectToUpdateDTO.Description)
                        objectToUpdate.Description = objectToUpdateDTO.Description;
                    if (objectToUpdate.IsSyncWithAD != objectToUpdateDTO.IsSyncWithAD)
                        objectToUpdate.IsSyncWithAD = objectToUpdateDTO.IsSyncWithAD;
                    if (objectToUpdate.IsServiceUser != objectToUpdateDTO.IsServiceUser)
                        objectToUpdate.IsServiceUser = objectToUpdateDTO.IsServiceUser;
                    if (objectToUpdate.SyncWithADGroupsLastTime != objectToUpdateDTO.SyncWithADGroupsLastTime)
                        objectToUpdate.SyncWithADGroupsLastTime = objectToUpdateDTO.SyncWithADGroupsLastTime;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.User.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<User, UserDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }
    }
}