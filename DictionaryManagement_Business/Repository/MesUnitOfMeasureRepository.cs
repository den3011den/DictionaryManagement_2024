using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class MesUnitOfMeasureRepository : IMesUnitOfMeasureRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public MesUnitOfMeasureRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<MesUnitOfMeasureDTO> Create(MesUnitOfMeasureDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<MesUnitOfMeasureDTO, MesUnitOfMeasure>(objectToAddDTO);
            var addedMesUnitOfMeasure = _db.MesUnitOfMeasure.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<MesUnitOfMeasure, MesUnitOfMeasureDTO>(addedMesUnitOfMeasure.Entity);
        }

        public async Task<MesUnitOfMeasureDTO> Get(int Id)
        {
            var objToGet = _db.MesUnitOfMeasure.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<MesUnitOfMeasure, MesUnitOfMeasureDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<MesUnitOfMeasureDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<MesUnitOfMeasure>, IEnumerable<MesUnitOfMeasureDTO>>(_db.MesUnitOfMeasure.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<MesUnitOfMeasure>, IEnumerable<MesUnitOfMeasureDTO>>(_db.MesUnitOfMeasure.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<MesUnitOfMeasure>, IEnumerable<MesUnitOfMeasureDTO>>(_db.MesUnitOfMeasure.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<MesUnitOfMeasure>, IEnumerable<MesUnitOfMeasureDTO>>(_db.MesUnitOfMeasure.ToListWithNoLock());
        }

        public async Task<MesUnitOfMeasureDTO> Update(MesUnitOfMeasureDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.MesUnitOfMeasure.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
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
                _db.MesUnitOfMeasure.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<MesUnitOfMeasure, MesUnitOfMeasureDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<MesUnitOfMeasureDTO> GetByName(string name)
        {
            var objToGet = _db.MesUnitOfMeasure.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesUnitOfMeasure, MesUnitOfMeasureDTO>(objToGet);
            }
            return null;
        }

        public async Task<MesUnitOfMeasureDTO> GetByShortName(string shortName)
        {
            var objToGet = _db.MesUnitOfMeasure.FirstOrDefaultWithNoLock(u => u.ShortName.Trim().ToUpper() == shortName.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesUnitOfMeasure, MesUnitOfMeasureDTO>(objToGet);
            }
            return null;
        }
    }
}
