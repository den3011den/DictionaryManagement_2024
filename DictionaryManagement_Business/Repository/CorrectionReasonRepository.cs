using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class CorrectionReasonRepository : ICorrectionReasonRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public CorrectionReasonRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<CorrectionReasonDTO> Create(CorrectionReasonDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<CorrectionReasonDTO, CorrectionReason>(objectToAddDTO);
            var addedCorrectionReason = _db.CorrectionReason.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<CorrectionReason, CorrectionReasonDTO>(addedCorrectionReason.Entity);
        }

        public async Task<CorrectionReasonDTO> Get(int Id)
        {
            var objToGet = _db.CorrectionReason.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<CorrectionReason, CorrectionReasonDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<CorrectionReasonDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                var ttt = _db.CorrectionReason;
                return _mapper.Map<IEnumerable<CorrectionReason>, IEnumerable<CorrectionReasonDTO>>(_db.CorrectionReason.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<CorrectionReason>, IEnumerable<CorrectionReasonDTO>>(_db.CorrectionReason.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<CorrectionReason>, IEnumerable<CorrectionReasonDTO>>(_db.CorrectionReason.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<CorrectionReason>, IEnumerable<CorrectionReasonDTO>>(_db.CorrectionReason.ToListWithNoLock());
        }

        public async Task<CorrectionReasonDTO> Update(CorrectionReasonDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.CorrectionReason.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
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
                _db.CorrectionReason.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<CorrectionReason, CorrectionReasonDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }
        public async Task<CorrectionReasonDTO> GetByName(string name)
        {
            var objToGet = _db.CorrectionReason.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<CorrectionReason, CorrectionReasonDTO>(objToGet);
            }
            return null;
        }


    }
}
