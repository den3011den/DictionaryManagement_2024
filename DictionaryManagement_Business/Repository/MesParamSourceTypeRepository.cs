using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class MesParamSourceTypeRepository : IMesParamSourceTypeRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public MesParamSourceTypeRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<MesParamSourceTypeDTO> Create(MesParamSourceTypeDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<MesParamSourceTypeDTO, MesParamSourceType>(objectToAddDTO);
            var addedMesParamSourceType = _db.MesParamSourceType.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<MesParamSourceType, MesParamSourceTypeDTO>(addedMesParamSourceType.Entity);
        }

        public async Task<MesParamSourceTypeDTO> Get(int Id)
        {
            var objToGet = _db.MesParamSourceType.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<MesParamSourceType, MesParamSourceTypeDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<MesParamSourceTypeDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<MesParamSourceType>, IEnumerable<MesParamSourceTypeDTO>>(_db.MesParamSourceType.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<MesParamSourceType>, IEnumerable<MesParamSourceTypeDTO>>(_db.MesParamSourceType.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<MesParamSourceType>, IEnumerable<MesParamSourceTypeDTO>>(_db.MesParamSourceType.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<MesParamSourceType>, IEnumerable<MesParamSourceTypeDTO>>(_db.MesParamSourceType.ToListWithNoLock());
        }

        public async Task<MesParamSourceTypeDTO> Update(MesParamSourceTypeDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.MesParamSourceType.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.Immutable != objectToUpdateDTO.Immutable)
                    {
                        objectToUpdate.Immutable = objectToUpdateDTO.Immutable;
                    }
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }

                _db.MesParamSourceType.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<MesParamSourceType, MesParamSourceTypeDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }

        public async Task<MesParamSourceTypeDTO> GetByName(string name)
        {
            var objToGet = _db.MesParamSourceType.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesParamSourceType, MesParamSourceTypeDTO>(objToGet);
            }
            return null;
        }
    }
}
