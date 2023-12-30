using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class LogEventTypeRepository : ILogEventTypeRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public LogEventTypeRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<LogEventTypeDTO> Create(LogEventTypeDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<LogEventTypeDTO, LogEventType>(objectToAddDTO);
            var addedLogEventType = _db.LogEventType.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<LogEventType, LogEventTypeDTO>(addedLogEventType.Entity);
        }

        public async Task<LogEventTypeDTO> Get(int Id)
        {
            var objToGet = _db.LogEventType.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<LogEventType, LogEventTypeDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<LogEventTypeDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<LogEventType>, IEnumerable<LogEventTypeDTO>>(_db.LogEventType.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<LogEventType>, IEnumerable<LogEventTypeDTO>>(_db.LogEventType.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<LogEventType>, IEnumerable<LogEventTypeDTO>>(_db.LogEventType.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<LogEventType>, IEnumerable<LogEventTypeDTO>>(_db.LogEventType.ToListWithNoLock());
        }

        public async Task<LogEventTypeDTO> Update(LogEventTypeDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.LogEventType.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.LogEventType.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<LogEventType, LogEventTypeDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }
        public async Task<LogEventTypeDTO> GetByName(string name)
        {
            var objToGet = _db.LogEventType.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<LogEventType, LogEventTypeDTO>(objToGet);
            }
            return null;
        }
    }
}
