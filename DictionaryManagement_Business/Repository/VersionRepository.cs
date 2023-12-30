using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class VersionRepository : IVersionRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public VersionRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<VersionDTO> Get()
        {
            var objToGet = _db.Version.FirstOrDefaultWithNoLock();
            if (objToGet == null)
            {
                var objToGetDTO = new VersionDTO
                {
                    version = ""
                };
                return objToGetDTO;
            }
            else
                return _mapper.Map<DictionaryManagement_DataAccess.Data.IntDB.Version, VersionDTO>(objToGet);
        }


        public async Task<VersionDTO> Set(VersionDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.Version.FirstOrDefaultWithNoLock();
            if (objectToUpdate != null)
            {
                if (objectToUpdate.version != objectToUpdateDTO.version)
                    objectToUpdate.version = objectToUpdateDTO.version;
                _db.Version.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<DictionaryManagement_DataAccess.Data.IntDB.Version, VersionDTO>(objectToUpdate);
            }
            else
            {
                var objectToAdd = _mapper.Map<VersionDTO, DictionaryManagement_DataAccess.Data.IntDB.Version>(objectToUpdateDTO);
                var addedVersion = _db.Version.Add(objectToAdd);
                _db.SaveChanges();
                return _mapper.Map<DictionaryManagement_DataAccess.Data.IntDB.Version, VersionDTO>(addedVersion.Entity);
            }
        }
    }
}