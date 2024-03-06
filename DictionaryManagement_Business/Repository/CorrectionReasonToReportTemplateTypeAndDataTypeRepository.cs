using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class CorrectionReasonToReportTemplateTypeAndDataTypeRepository : ICorrectionReasonToReportTemplateTypeAndDataTypeRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public CorrectionReasonToReportTemplateTypeAndDataTypeRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> Create(CorrectionReasonToReportTemplateTypeAndDataTypeDTO objectToAddDTO)
        {
            CorrectionReasonToReportTemplateTypeAndDataType objectToAdd = new CorrectionReasonToReportTemplateTypeAndDataType();
            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.CorrectionReasonId = objectToAddDTO.CorrectionReasonId;
            objectToAdd.ReportTemplateTypeId = objectToAddDTO.ReportTemplateTypeId;
            objectToAdd.DataTypeId = objectToAddDTO.DataTypeId;
            try
            {
                var addedCorrectionReasonToReportTemplateTypeAndDataType = _db.CorrectionReasonToReportTemplateTypeAndDataType.Add(objectToAdd);
                _db.SaveChanges();
                return (await GetById(addedCorrectionReasonToReportTemplateTypeAndDataType.Entity.Id));
            }
            catch
            {
                return null;
            }
        }

        public async Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> GetByCorrectionReasonIdAndReportTemplateTypeIdAndDataType(int correctionReasonId, int reportTemplateTypeId, int dataTypeId)
        {
            try
            {
                var objToGet = _db.CorrectionReasonToReportTemplateTypeAndDataType
                    .Include("CorrectionReasonFK")
                    .Include("ReportTemplateTypeFK")
                    .Include("DataTypeFK")
                        .FirstOrDefaultWithNoLock(u => u.CorrectionReasonId == correctionReasonId && u.ReportTemplateTypeId == reportTemplateTypeId
                            && u.DataTypeId == dataTypeId);
                if (objToGet != null)
                {
                    return _mapper.Map<CorrectionReasonToReportTemplateTypeAndDataType, CorrectionReasonToReportTemplateTypeAndDataTypeDTO>(objToGet);
                }
            }
            catch { };
            return null;
        }

        public async Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> GetById(int id)
        {
            try
            {
                var objToGet = _db.CorrectionReasonToReportTemplateTypeAndDataType
                    .Include("CorrectionReasonFK")
                    .Include("ReportTemplateTypeFK")
                    .Include("DataTypeFK")
                                .FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objToGet != null)
                {
                    return _mapper.Map<CorrectionReasonToReportTemplateTypeAndDataType, CorrectionReasonToReportTemplateTypeAndDataTypeDTO>(objToGet);
                }
            }
            catch
            { }
            return null;
        }

        public async Task<IEnumerable<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>> GetAll()
        {
            try
            {
                var hhh = _db.CorrectionReasonToReportTemplateTypeAndDataType
                    .Include("CorrectionReasonFK")
                    .Include("ReportTemplateTypeFK")
                    .Include("DataTypeFK")
                    .ToListWithNoLock();

                if (hhh != null)
                    return _mapper.Map<IEnumerable<CorrectionReasonToReportTemplateTypeAndDataType>, IEnumerable<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>>(hhh);
            }
            catch { };
            return new List<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>();
        }

        public async Task<IEnumerable<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>?> GetByCorrectionReasonId(int correctionReasonId)
        {
            try
            {
                var hhh = _db.CorrectionReasonToReportTemplateTypeAndDataType
                    .Include("CorrectionReasonFK")
                    .Include("ReportTemplateTypeFK")
                    .Include("DataTypeFK")
                    .Where(u => u.CorrectionReasonId == correctionReasonId)
                    .ToListWithNoLock();
                return _mapper.Map<IEnumerable<CorrectionReasonToReportTemplateTypeAndDataType>, IEnumerable<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>>(hhh);
            }
            catch { };
            return null;
        }

        public async Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> Update(CorrectionReasonToReportTemplateTypeAndDataTypeDTO objectToUpdateDTO)
        {
            try
            {
                var objectToUpdate = _db.CorrectionReasonToReportTemplateTypeAndDataType
                        .Include("CorrectionReasonFK")
                        .Include("ReportTemplateTypeFK")
                        .Include("DataTypeFK")
                        .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
                if (objectToUpdate != null)
                {
                    if (objectToUpdate.CorrectionReasonId != objectToUpdateDTO.CorrectionReasonDTOFK.Id)
                    {
                        objectToUpdate.CorrectionReasonId = objectToUpdateDTO.CorrectionReasonDTOFK.Id;
                        objectToUpdate.CorrectionReasonFK = _mapper.Map<CorrectionReasonDTO, CorrectionReason>(objectToUpdateDTO.CorrectionReasonDTOFK);
                    }
                    if (objectToUpdate.ReportTemplateTypeId != objectToUpdateDTO.ReportTemplateTypeDTOFK.Id)
                    {
                        objectToUpdate.ReportTemplateTypeId = objectToUpdateDTO.ReportTemplateTypeDTOFK.Id;
                        objectToUpdate.ReportTemplateTypeFK = _mapper.Map<ReportTemplateTypeDTO, ReportTemplateType>(objectToUpdateDTO.ReportTemplateTypeDTOFK);
                    }
                    if (objectToUpdate.DataTypeId != objectToUpdateDTO.DataTypeDTOFK.Id)
                    {
                        objectToUpdate.DataTypeId = objectToUpdateDTO.DataTypeDTOFK.Id;
                        objectToUpdate.DataTypeFK = _mapper.Map<DataTypeDTO, DataType>(objectToUpdateDTO.DataTypeDTOFK);
                    }
                    _db.CorrectionReasonToReportTemplateTypeAndDataType.Update(objectToUpdate);
                    _db.SaveChanges();
                    return _mapper.Map<CorrectionReasonToReportTemplateTypeAndDataType, CorrectionReasonToReportTemplateTypeAndDataTypeDTO>(objectToUpdate);
                }
            }
            catch { };
            return null;

        }

        public async Task<int> Delete(int id)
        {
            if (id > 0)
            {
                try
                {
                    var objectToDelete = _db.CorrectionReasonToReportTemplateTypeAndDataType.FirstOrDefaultWithNoLock(u => u.Id == id);
                    if (objectToDelete != null)
                    {
                        _db.CorrectionReasonToReportTemplateTypeAndDataType.Remove(objectToDelete);
                        return _db.SaveChanges();
                    }
                }
                catch { };
            }
            return 0;
        }
    }
}
