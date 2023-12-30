using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class DataTypeRepository : IDataTypeRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public DataTypeRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<DataTypeDTO> Create(DataTypeDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<DataTypeDTO, DataType>(objectToAddDTO);
            var addedDataType = _db.DataType.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<DataType, DataTypeDTO>(addedDataType.Entity);
        }

        public async Task<DataTypeDTO> Get(int Id)
        {
            var objToGet = _db.DataType.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<DataType, DataTypeDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<DataTypeDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<DataType>, IEnumerable<DataTypeDTO>>(_db.DataType.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<DataType>, IEnumerable<DataTypeDTO>>(_db.DataType.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<DataType>, IEnumerable<DataTypeDTO>>(_db.DataType.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<DataType>, IEnumerable<DataTypeDTO>>(_db.DataType.ToListWithNoLock());
        }

        public async Task<DataTypeDTO> Update(DataTypeDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.DataType.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.Priority != objectToUpdateDTO.Priority)
                        objectToUpdate.Priority = objectToUpdateDTO.Priority;
                    if (objectToUpdate.IsAutoCalcDestDataType != objectToUpdateDTO.IsAutoCalcDestDataType)
                        objectToUpdate.IsAutoCalcDestDataType = objectToUpdateDTO.IsAutoCalcDestDataType;

                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.DataType.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<DataType, DataTypeDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<DataTypeDTO> GetByName(string name)
        {
            var objToGet = _db.DataType.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<DataType, DataTypeDTO>(objToGet);
            }
            return null;
        }

        public async Task<DataTypeDTO> GetByPriority(int? priority)
        {
            if (priority != null)
                if (priority != 0)
                {
                    var objToGet = _db.DataType.FirstOrDefaultWithNoLock(u => u.Priority == priority);
                    if (objToGet != null)
                    {
                        return _mapper.Map<DataType, DataTypeDTO>(objToGet);
                    }
                }
            return null;
        }
    }
}
