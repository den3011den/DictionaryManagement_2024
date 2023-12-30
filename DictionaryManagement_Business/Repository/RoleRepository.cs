using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class RoleRepository : IRoleRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public RoleRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<RoleDTO> Create(RoleDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<RoleDTO, Role>(objectToAddDTO);
            var addedRole = _db.Role.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<Role, RoleDTO>(addedRole.Entity);
        }

        public async Task<RoleDTO> GetById(Guid Id)
        {
            if (Id != null && Id != Guid.Empty)
            {
                var objToGet = _db.Role.FirstOrDefaultWithNoLock(u => (u.Id == Id));
                if (objToGet != null)
                {
                    return _mapper.Map<Role, RoleDTO>(objToGet);
                }
            }
            return null;
        }

        public async Task<IEnumerable<RoleDTO>> GetAll(SD.SelectDictionaryScope selectDictionaryScope = SD.SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<Role>, IEnumerable<RoleDTO>>(_db.Role.ToListWithNoLock());
        }

        public async Task<RoleDTO> GetByName(string name = "")
        {
            var objToGet = _db.Role.FirstOrDefaultWithNoLock(u => ((u.Name.Trim().ToUpper()) == (name.Trim().ToUpper())));
            if (objToGet != null)
            {
                return _mapper.Map<Role, RoleDTO>(objToGet);
            }
            return null;
        }


        public async Task<RoleDTO> Update(RoleDTO objectToUpdateDTO, SD.UpdateMode updateMode = SD.UpdateMode.Update)
        {
            var objectToUpdate = _db.Role.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.Description != objectToUpdateDTO.Description)
                        objectToUpdate.Description = objectToUpdateDTO.Description;
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
                return _mapper.Map<Role, RoleDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }
    }
}
