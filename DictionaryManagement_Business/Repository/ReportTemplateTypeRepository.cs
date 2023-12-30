using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class ReportTemplateTypeRepository : IReportTemplateTypeRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ReportTemplateTypeRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<ReportTemplateTypeDTO> Create(ReportTemplateTypeDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<ReportTemplateTypeDTO, ReportTemplateType>(objectToAddDTO);
            var addedReportTemplateType = _db.ReportTemplateType.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<ReportTemplateType, ReportTemplateTypeDTO>(addedReportTemplateType.Entity);
        }

        public async Task<ReportTemplateTypeDTO> Get(int Id)
        {
            var objToGet = _db.ReportTemplateType.FirstOrDefaultWithNoLock(u => u.Id == Id);
            if (objToGet != null)
            {
                return _mapper.Map<ReportTemplateType, ReportTemplateTypeDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<ReportTemplateTypeDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                return _mapper.Map<IEnumerable<ReportTemplateType>, IEnumerable<ReportTemplateTypeDTO>>(_db.ReportTemplateType.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<ReportTemplateType>, IEnumerable<ReportTemplateTypeDTO>>(_db.ReportTemplateType.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<ReportTemplateType>, IEnumerable<ReportTemplateTypeDTO>>(_db.ReportTemplateType.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<ReportTemplateType>, IEnumerable<ReportTemplateTypeDTO>>(_db.ReportTemplateType.ToListWithNoLock());
        }

        public async Task<ReportTemplateTypeDTO> Update(ReportTemplateTypeDTO objectToUpdateDTO, UpdateMode updateMode = UpdateMode.Update)
        {
            var objectToUpdate = _db.ReportTemplateType.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.NeedAutoCalc != objectToUpdateDTO.NeedAutoCalc)
                        objectToUpdate.NeedAutoCalc = objectToUpdateDTO.NeedAutoCalc;
                    if (objectToUpdate.CanAutoCalc != objectToUpdateDTO.CanAutoCalc)
                        objectToUpdate.CanAutoCalc = objectToUpdateDTO.CanAutoCalc;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.ReportTemplateType.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<ReportTemplateType, ReportTemplateTypeDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }
        public async Task<ReportTemplateTypeDTO> GetByName(string name)
        {
            var objToGet = _db.ReportTemplateType.FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<ReportTemplateType, ReportTemplateTypeDTO>(objToGet);
            }
            return null;
        }
    }
}
