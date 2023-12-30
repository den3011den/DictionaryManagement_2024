using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class MesMaterialRepository : IMesMaterialRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public MesMaterialRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<MesMaterialDTO> Create(MesMaterialDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<MesMaterialDTO, MesMaterial>(objectToAddDTO);
            var addedMesMaterial = _db.MesMaterial.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<MesMaterial, MesMaterialDTO>(addedMesMaterial.Entity);
        }

        public async Task<MesMaterialDTO> Get(int id)
        {
            if (id > 0)
            {
                var objToGet = _db.MesMaterial.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objToGet != null)
                {
                    return _mapper.Map<MesMaterial, MesMaterialDTO>(objToGet);
                }
            }
            return null;
        }

        public async Task<MesMaterialDTO> GetByCode(string code = "")
        {
            var objToGet = _db.MesMaterial.FirstOrDefaultWithNoLock(u => u.Code.Trim().ToUpper() == code.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesMaterial, MesMaterialDTO>(objToGet);
            }
            return null;
        }

        public async Task<MesMaterialDTO> GetByName(string name = "")
        {
            var objToGet = _db.MesMaterial.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesMaterial, MesMaterialDTO>(objToGet);
            }
            return null;
        }

        public async Task<MesMaterialDTO> GetByShortName(string shortName = "")
        {
            var objToGet = _db.MesMaterial.FirstOrDefaultWithNoLock(u => u.ShortName.Trim().ToUpper() == shortName.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesMaterial, MesMaterialDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<MesMaterialDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<MesMaterial>, IEnumerable<MesMaterialDTO>>(_db.MesMaterial.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<MesMaterial>, IEnumerable<MesMaterialDTO>>(_db.MesMaterial.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<MesMaterial>, IEnumerable<MesMaterialDTO>>(_db.MesMaterial.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<MesMaterial>, IEnumerable<MesMaterialDTO>>(_db.MesMaterial.ToListWithNoLock());
        }

        public async Task<MesMaterialDTO> Update(MesMaterialDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.MesMaterial.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Code != objectToUpdateDTO.Code)
                        objectToUpdate.Code = objectToUpdateDTO.Code;
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.ShortName != objectToUpdateDTO.ShortName)
                        objectToUpdate.ShortName = objectToUpdateDTO.ShortName;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.MesMaterial.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<MesMaterial, MesMaterialDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }
    }
}
