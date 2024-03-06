using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class DataSourceRepository : IDataSourceRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public DataSourceRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<DataSourceDTO> Create(DataSourceDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<DataSourceDTO, DataSource>(objectToAddDTO);
            var addedDataSource = _db.DataSource.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<DataSource, DataSourceDTO>(addedDataSource.Entity);
        }

        public async Task<DataSourceDTO> Get(int Id)
        {
            var objToGet = _db.DataSource.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<DataSource, DataSourceDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<DataSourceDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<DataSource>, IEnumerable<DataSourceDTO>>(_db.DataSource.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<DataSource>, IEnumerable<DataSourceDTO>>(_db.DataSource.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<DataSource>, IEnumerable<DataSourceDTO>>(_db.DataSource.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<DataSource>, IEnumerable<DataSourceDTO>>(_db.DataSource.ToListWithNoLock());
        }

        public async Task<DataSourceDTO> Update(DataSourceDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.DataSource.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.Immutable != objectToUpdateDTO.Immutable)
                        objectToUpdate.Immutable = objectToUpdateDTO.Immutable;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.DataSource.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<DataSource, DataSourceDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<DataSourceDTO> GetByName(string name)
        {
            var objToGet = _db.DataSource.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<DataSource, DataSourceDTO>(objToGet);
            }
            return null;
        }
    }
}
