using AutoMapper;
using ClosedXML.Excel;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class ReportTemplateToMesParamRepository : IReportTemplateToMesParamRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;
        private readonly IReportTemplateRepository _reportTemplateRepository;
        private readonly IReportTemplateTypeRepository _reportTemplateTypeRepository;

        public ReportTemplateToMesParamRepository(IntDBApplicationDbContext db, IMapper mapper, IReportTemplateRepository reportTemplateRepository, IReportTemplateTypeRepository reportTemplateTypeRepository)
        {
            _db = db;
            _mapper = mapper;
            _reportTemplateRepository = reportTemplateRepository;
            _reportTemplateTypeRepository = reportTemplateTypeRepository;
        }

        public async Task<ReportTemplateToMesParamDTO?> Create(ReportTemplateToMesParamDTO objectToAddDTO)
        {
            ReportTemplateToMesParam objectToAdd = new ReportTemplateToMesParam();
            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.ReportTemplateId = objectToAddDTO.ReportTemplateId;
            objectToAdd.MesParamId = objectToAddDTO.MesParamId;
            objectToAdd.MesParamCode = objectToAddDTO.MesParamCode == null ? objectToAddDTO.MesParamCode : objectToAddDTO.MesParamCode.Substring(0, Math.Min(100, objectToAddDTO.MesParamCode.Length));
            objectToAdd.SheetName = objectToAddDTO.SheetName;
            try
            {
                var addedReportTemplateToMesParam = _db.ReportTemplateToMesParam.Add(objectToAdd);
                _db.SaveChanges();
                return await GetById(addedReportTemplateToMesParam.Entity.Id);
            }
            catch
            {
                return null;
            }
        }

        public async Task<ReportTemplateToMesParamDTO?> GetById(int id)
        {
            try
            {
                var objToGet = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objToGet != null)
                {
                    return _mapper.Map<ReportTemplateToMesParam, ReportTemplateToMesParamDTO>(objToGet);
                }
            }
            catch
            { }
            return null;
        }

        public async Task<ReportTemplateToMesParamDTO?> Update(ReportTemplateToMesParamDTO objectToUpdateDTO)
        {
            try
            {
                var objectToUpdate = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
                if (objectToUpdate != null)
                {
                    if (objectToUpdate.ReportTemplateId != objectToUpdateDTO.ReportTemplateDTOFK.Id)
                    {
                        objectToUpdate.ReportTemplateId = objectToUpdateDTO.ReportTemplateDTOFK.Id;
                        objectToUpdate.ReportTemplateFK = _mapper.Map<ReportTemplateDTO, ReportTemplate>(objectToUpdateDTO.ReportTemplateDTOFK);
                    }
                    if (objectToUpdateDTO.MesParamId != null)
                    {
                        if (objectToUpdate.MesParamId != objectToUpdateDTO.MesParamId)
                        {
                            objectToUpdate.MesParamId = objectToUpdateDTO.MesParamId;
                            //objectToUpdate.MesParamFK = _mapper.Map<MesParamDTO, MesParam>(objectToUpdateDTO.MesParamDTOFK);
                        }
                    }
                    else
                    {
                        objectToUpdate.MesParamId = null;
                        objectToUpdate.MesParamFK = null;
                    }
                    if (objectToUpdate.MesParamCode != objectToUpdateDTO.MesParamCode)
                        objectToUpdate.MesParamCode = objectToUpdateDTO.MesParamCode == null ? objectToUpdateDTO.MesParamCode : objectToUpdateDTO.MesParamCode.Substring(0, Math.Min(100, objectToUpdateDTO.MesParamCode.Length));
                    if (objectToUpdate.SheetName != objectToUpdateDTO.SheetName)
                        objectToUpdate.SheetName = objectToUpdateDTO.SheetName;

                    _db.ReportTemplateToMesParam.Update(objectToUpdate);
                    _db.SaveChanges();
                    return _mapper.Map<ReportTemplateToMesParam, ReportTemplateToMesParamDTO>(objectToUpdate);
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
                    var objectToDelete = _db.ReportTemplateToMesParam.FirstOrDefaultWithNoLock(u => u.Id == id);
                    if (objectToDelete != null)
                    {
                        _db.ReportTemplateToMesParam.Remove(objectToDelete);
                        return _db.SaveChanges();
                    }
                }
                catch { };
            }
            return 0;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>> GetAll()
        {
            try
            {
                var hhh = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .ToListWithNoLock();

                if (hhh != null)
                    return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(hhh);
            }
            catch { };
            return new List<ReportTemplateToMesParamDTO>();
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByReportTemplateId(Guid reportTemplateId)
        {
            try
            {
                var objToGet = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .Where(u => u.ReportTemplateId == reportTemplateId)
                    .ToListWithNoLock();
                if (objToGet != null)
                {
                    return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                }
            }
            catch { };
            return null;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByMesParamId(int mesParamId)
        {
            try
            {
                var objToGet = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .Where(u => u.MesParamId == mesParamId)
                    .ToListWithNoLock();
                if (objToGet != null)
                {
                    return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                }
            }
            catch { };
            return null;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByMesParamCode(string mesParamCode, bool? reportTemplateIsInArchive = null)
        {
            try
            {
                if (reportTemplateIsInArchive == null)
                {
                    var objToGet = _db.ReportTemplateToMesParam
                        .Include("ReportTemplateFK")
                        .Include("ReportTemplateFK.MesDepartmentFK")
                        .Include("ReportTemplateFK.ReportTemplateTypeFK")
                        .Include("MesParamFK")
                        .Where(u => u.MesParamCode == mesParamCode)
                        .ToListWithNoLock();
                    if (objToGet != null)
                    {
                        return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                    }
                }
                else
                {
                    var objToGet = _db.ReportTemplateToMesParam
                        .Include("ReportTemplateFK")
                        .Include("ReportTemplateFK.MesDepartmentFK")
                        .Include("ReportTemplateFK.ReportTemplateTypeFK")
                        .Include("MesParamFK")
                        .Where(u => u.MesParamCode == mesParamCode && u.ReportTemplateFK.IsArchive == reportTemplateIsInArchive)
                        .ToListWithNoLock();
                    if (objToGet != null)
                    {
                        return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                    }
                }
            }
            catch { };
            return null;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetBySheetName(string sheetName)
        {
            try
            {
                var objToGet = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .Where(u => u.SheetName.ToUpper() == sheetName.ToUpper())
                    .ToListWithNoLock();
                if (objToGet != null)
                {
                    return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                }
            }
            catch { };
            return null;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByMesParamIdOrAndMesParamCode(int? mesParamId, string? mesParamCode)
        {
            if (mesParamId == null && mesParamCode == null)
                return null;
            if (mesParamId != null && mesParamCode == null)
            {
                try
                {
                    var objToGet = await GetByMesParamId((int)mesParamId);
                    if (objToGet != null)
                    {
                        return objToGet;
                    }
                }
                catch { };
            }
            if (mesParamId == null && mesParamCode != null)
            {
                try
                {
                    var objToGet = await GetByMesParamCode((string)mesParamCode);
                    if (objToGet != null)
                    {
                        return objToGet;
                    }
                }
                catch { };
            }

            if (mesParamId != null && mesParamCode != null)
            {
                try
                {
                    var objToGet = _db.ReportTemplateToMesParam
                        .Include("ReportTemplateFK")
                        .Include("ReportTemplateFK.MesDepartmentFK")
                        .Include("ReportTemplateFK.ReportTemplateTypeFK")
                        .Include("MesParamFK")
                        .Where(u => u.MesParamId == mesParamId && u.MesParamCode == mesParamCode.ToUpper())
                        .ToListWithNoLock();
                    if (objToGet != null)
                    {
                        return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                    }
                }
                catch { };
            }

            return null;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByReportTemplateIdAndMesParamIdAndSheetName(Guid reportTemplateTypeId, int mesParamId, string sheetName)
        {
            try
            {
                var objToGet = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .Where(u => u.ReportTemplateId == reportTemplateTypeId && u.MesParamId == mesParamId && u.SheetName.ToUpper() == sheetName.ToUpper())
                    .ToListWithNoLock();
                if (objToGet != null)
                {
                    return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                }
            }
            catch { };
            return null;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetReportTemplateIdAndMesParamCodeAndSheetName(Guid reportTemplateTypeId, string mesParamCode, string sheetName)
        {
            try
            {
                var objToGet = _db.ReportTemplateToMesParam
                    .Include("ReportTemplateFK")
                    .Include("ReportTemplateFK.MesDepartmentFK")
                    .Include("ReportTemplateFK.ReportTemplateTypeFK")
                    .Include("MesParamFK")
                    .Where(u => u.ReportTemplateId == reportTemplateTypeId && u.MesParamCode == mesParamCode && u.SheetName.ToUpper() == sheetName.ToUpper())
                    .ToListWithNoLock();
                if (objToGet != null)
                {
                    return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(objToGet);
                }
            }
            catch { };
            return null;
        }

        public async Task<int> DeleteAllByReportTemplateId(Guid reportTemplateid)
        {
            if (reportTemplateid != Guid.Empty)
            {
                try
                {
                    var listForDelete = _db.ReportTemplateToMesParam.Where(u => u.ReportTemplateId == reportTemplateid).ToList();
                    if (listForDelete != null)
                    {
                        foreach (var item in listForDelete)
                        {
                            _db.ReportTemplateToMesParam.Remove(item);
                        }
                        return _db.SaveChanges();
                    }
                }
                catch { };
            }
            return 0;
        }

        public async Task<int> CreateByList(IEnumerable<ReportTemplateToMesParamDTO> listToAddDTO)
        {
            if (listToAddDTO.Any())
            {

                foreach (var item in listToAddDTO)
                {
                    ReportTemplateToMesParam objectToAdd = new ReportTemplateToMesParam();
                    objectToAdd.Id = item.Id;
                    objectToAdd.ReportTemplateId = item.ReportTemplateId;
                    objectToAdd.MesParamId = item.MesParamId;
                    objectToAdd.MesParamCode = item.MesParamCode == null ? item.MesParamCode : item.MesParamCode.Substring(0, Math.Min(100, item.MesParamCode.Length));
                    objectToAdd.SheetName = item.SheetName;
                    var addedReportTemplateToMesParam = _db.ReportTemplateToMesParam.Add(objectToAdd);
                }
                return _db.SaveChanges();
            }
            return 0;
        }

        public async Task AddFromWorksheet(Guid reportTemplateId, IXLWorkbook workbook, string sheetName, string columnName, int skipRowsCount, IEnumerable<MesParamDTO> mesParamListDTO)
        {
            workbook.TryGetWorksheet(sheetName, out IXLWorksheet worksheet);
            if (worksheet != null)
            {
                var mesParamCodeList = worksheet.Range(columnName + (1 + skipRowsCount).ToString() + ":" + columnName + "1000000").CellsUsed().Select(c => c.CachedValue.ToString()/*.Trim()*/).ToList();
                mesParamCodeList = mesParamCodeList.Where(u => !String.IsNullOrEmpty(u.Trim())).Distinct().ToList();


                var returnResult = (from mesParamCodeListAlias in mesParamCodeList
                                    join mesParamListDTOAlias in mesParamListDTO on
                                     mesParamCodeListAlias equals mesParamListDTOAlias.Code
                                into MP_prom2
                                    from MP in MP_prom2.DefaultIfEmpty()
                                    select
                                            new ReportTemplateToMesParamDTO
                                            {
                                                ReportTemplateId = reportTemplateId,
                                                MesParamId = (MP == null ? null : MP.Id),
                                                MesParamCode = mesParamCodeListAlias,
                                                SheetName = sheetName,
                                            }).ToList();
                if (returnResult != null)
                    await CreateByList(returnResult);

                //var result = sheetList.Select(u => u.Trim().ToUpper()).ToList().Except(sheets.Select(u => u.Name.Trim().ToUpper()).ToList()).ToList();
                //// выбор тех тэгов, которые в архиве и при этом есть в базе
                //List<string> isInArchiveList = (mesParamCodeList.Intersect(mesParamDTOList.Select(u => u.Code).ToList(), StringComparer.OrdinalIgnoreCase).ToList())
                //    .Except(mesParamDTOList.Where(u => u.IsArchive != true).Select(u => u.Code).ToList(), StringComparer.OrdinalIgnoreCase).ToList();
            }
        }

        public async Task<int> UpdateEmptyMesParamIdByMesParamCode(string mesParamCode, int mesParamId)
        {
            int retCount = 0;
            var tmpList = await GetByMesParamCode(mesParamCode);
            if (tmpList != null && tmpList.Any())
            {
                var tmpList2 = tmpList.Where(u => u.MesParamId == null).ToList();
                if (tmpList2 != null && tmpList2.Any())
                {
                    foreach (var item in tmpList2)
                    {
                        item.MesParamId = mesParamId;
                        await Update(item);
                        retCount++;
                    }
                }
            }
            return retCount;
        }

        public async Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetTagListInOtherNotArchiveReportTemplatesBySheetName(Guid sourceReportTemplateId, string sourceSheetName, string destSheetName
                , List<string> sourceCodeList = null, string reportTemplateTypeName = null)
        {
            try
            {
                List<string> sourceList = null;
                if (sourceCodeList != null)
                {
                    sourceList = sourceCodeList.Distinct().ToList();
                }
                else
                {
                    var sourceListTmp = _db.ReportTemplateToMesParam
                        .Where(u => u.ReportTemplateId == sourceReportTemplateId && u.SheetName.ToUpper() == sourceSheetName.ToUpper())
                        .ToListWithNoLock();
                    if (sourceListTmp != null && sourceListTmp.Any())
                    {
                        sourceList = sourceListTmp.Select(u => u.MesParamCode).Distinct().ToList();
                    }
                }

                if ((sourceList != null && sourceList.Any()) || (sourceCodeList != null && sourceCodeList.Any()))
                {
                    int findReportTemplateTypeId1 = -999999;
                    int findReportTemplateTypeId2 = -999999;


                    string reportTemplateTypeNameToCompare = "";
                    if (String.IsNullOrEmpty(reportTemplateTypeName))
                    {
                        var reportTemplateDTO = await _reportTemplateRepository.GetById(sourceReportTemplateId);
                        if (reportTemplateDTO != null)
                            reportTemplateTypeNameToCompare = reportTemplateDTO.ReportTemplateTypeDTOFK.Name.Trim().ToUpper();
                    }
                    else
                    {
                        reportTemplateTypeNameToCompare = reportTemplateTypeName.Trim().ToUpper();
                    }

                    if (reportTemplateTypeNameToCompare == SD.CorrectionReportTemplateTypeName.Trim().ToUpper())
                    {
                        var dto1 = await _reportTemplateTypeRepository.GetByName(reportTemplateTypeNameToCompare);
                        if (dto1 != null) { findReportTemplateTypeId1 = dto1.Id; }
                    }
                    else
                    {
                        if (reportTemplateTypeNameToCompare == SD.EmbReportTemplateTypeName.Trim().ToUpper()
                            || reportTemplateTypeNameToCompare == SD.TebReportTemplateTypeName.Trim().ToUpper())
                        {
                            var dto1 = await _reportTemplateTypeRepository.GetByName(SD.EmbReportTemplateTypeName);
                            if (dto1 != null) { findReportTemplateTypeId1 = dto1.Id; }
                            var dto2 = await _reportTemplateTypeRepository.GetByName(SD.TebReportTemplateTypeName);
                            if (dto2 != null) { findReportTemplateTypeId2 = dto2.Id; }
                        }
                        else
                        {
                            return null;
                        }
                    }

                    var destList = _db.ReportTemplateToMesParam
                        .Include("ReportTemplateFK")
                        .Include("ReportTemplateFK.MesDepartmentFK")
                        .Include("ReportTemplateFK.ReportTemplateTypeFK")
                        .Include("MesParamFK")
                        .Where(u => u.ReportTemplateFK.IsArchive != true && u.SheetName.ToUpper() == destSheetName.ToUpper()
                            && (u.ReportTemplateId != sourceReportTemplateId && (u.ReportTemplateFK.ReportTemplateTypeId == findReportTemplateTypeId1 || u.ReportTemplateFK.ReportTemplateTypeId == findReportTemplateTypeId2)))
                        .ToListWithNoLock();

                    if (destList != null && destList.Any())
                    {
                        IEnumerable<ReportTemplateToMesParam> resultList = null;
                        resultList = (from destListAlias in destList
                                      join sourceListAlias in sourceList on destListAlias.MesParamCode equals sourceListAlias
                                      select
                                            new ReportTemplateToMesParam
                                            {
                                                Id = destListAlias.Id,
                                                ReportTemplateId = destListAlias.ReportTemplateId,
                                                ReportTemplateFK = destListAlias.ReportTemplateFK,
                                                MesParamCode = destListAlias.MesParamCode,
                                                MesParamId = destListAlias.MesParamId,
                                                MesParamFK = destListAlias.MesParamFK,
                                                SheetName = destListAlias.SheetName,
                                            }).ToList();
                        if (resultList != null && resultList.Any())
                        {
                            return _mapper.Map<IEnumerable<ReportTemplateToMesParam>, IEnumerable<ReportTemplateToMesParamDTO>>(resultList);
                        }
                    }
                }
            }
            catch { };
            return null;
        }
    }
}
