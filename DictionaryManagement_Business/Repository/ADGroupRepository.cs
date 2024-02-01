using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class ADGroupRepository : IADGroupRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ADGroupRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<ADGroupDTO> Create(ADGroupDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<ADGroupDTO, ADGroup>(objectToAddDTO);
            var addedADGroup = _db.ADGroup.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<ADGroup, ADGroupDTO>(addedADGroup.Entity);
        }

        public async Task<ADGroupDTO> Get(Guid Id)
        {
            if (Id != null && Id != Guid.Empty)
            {
                var objToGet = _db.ADGroup.FirstOrDefaultWithNoLock(u => (u.Id == Id));
                if (objToGet != null)
                {
                    return _mapper.Map<ADGroup, ADGroupDTO>(objToGet);
                }
            }
            return null;
        }

        public async Task<IEnumerable<ADGroupDTO>> GetAll(SD.SelectDictionaryScope selectDictionaryScope = SD.SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<ADGroup>, IEnumerable<ADGroupDTO>>(_db.ADGroup.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<ADGroup>, IEnumerable<ADGroupDTO>>(_db.ADGroup.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<ADGroup>, IEnumerable<ADGroupDTO>>(_db.ADGroup.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<ADGroup>, IEnumerable<ADGroupDTO>>(_db.ADGroup.ToListWithNoLock());
        }

        public async Task<ADGroupDTO> GetByName(string name = "")
        {
            var objToGet = _db.ADGroup.FirstOrDefaultWithNoLock(u => ((u.Name.Trim().ToUpper()) == (name.Trim().ToUpper())));
            if (objToGet != null)
            {
                return _mapper.Map<ADGroup, ADGroupDTO>(objToGet);
            }
            return null;
        }


        public async Task<ADGroupDTO> Update(ADGroupDTO objectToUpdateDTO, SD.UpdateMode updateMode = SD.UpdateMode.Update)
        {
            var objectToUpdate = _db.ADGroup.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
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
                _db.ADGroup.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<ADGroup, ADGroupDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }
    }
}
