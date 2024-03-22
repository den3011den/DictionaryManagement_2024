using AutoMapper;
using ClosedXML.Excel;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class CheckReportTemplateRepository : ICheckReportTemplateRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;
        private readonly IReportTemplateToMesParamRepository _reportTemplateToMesParamRepository;

        public CheckReportTemplateRepository(IntDBApplicationDbContext db, IMapper mapper, IReportTemplateToMesParamRepository reportTemplateToMesParamRepository)
        {
            _db = db;
            _mapper = mapper;
            _reportTemplateToMesParamRepository = reportTemplateToMesParamRepository;

        }

        public record OutputDataRecord
        {
            public string MesParamCode { get; set; } = "";
            public string ValueTime { get; set; } = "";
        }

        public async Task<List<string>?> IsNotExistSheets(IXLWorkbook workbook, List<string> sheetList)
        {
            var sheets = workbook.Worksheets.ToList();
            if (sheets == null || sheets.Count <= 0)
            {
                return sheetList;
            }
            var result = sheetList.Select(u => u.Trim().ToUpper()).ToList().Except(sheets.Select(u => u.Name.Trim().ToUpper()).ToList()).ToList();
            if (result == null || result.Count <= 0)
                return null;
            return result;
        }

        public async Task<bool> IsEmptySheet(IXLWorksheet worksheet)
        {
            var rowsUsed = worksheet.RowsUsed();
            if (rowsUsed == null || rowsUsed.Count() <= 0)
                return true;
            return false;
        }

        // возвращает список не найденых заголовков в первой строке листа
        public async Task<List<SheetHeader>?> CheckSheetHeader(IXLWorksheet worksheet, List<SheetHeader>? sheetHeaderList)
        {
            if (sheetHeaderList != null)
            {
                if (sheetHeaderList.Count() > 0)
                {
                    List<SheetHeader> resultList = new List<SheetHeader>();
                    var rowVar = worksheet.Row(1);
                    foreach (var sheetHeader in sheetHeaderList)
                    {
                        if (!rowVar.Cell(sheetHeader.SheetHeaderColumnNumber).CachedValue.ToString().Trim().Replace(" ", "").Replace("_", "").ToUpper().Equals(sheetHeader.SheetHeaderColumnName.Trim().ToUpper()))
                        {
                            resultList.Add(sheetHeader);
                        }
                    }
                    if (resultList.Count() > 0)
                    { return resultList; }
                    else
                    { return null; }
                }
            }
            return null;
        }

        public async Task<List<string>?> CheckSheetTags(IXLWorksheet worksheet, IEnumerable<MesParamDTO> mesParamDTOList, CheckReportTemplateTagsType checkReportTemplateTagsType
            , string reportTemplateTypeName, Guid reportTemplateId)
        {
            List<string> mesParamCodeList;
            List<OutputDataRecord> outputDataList = new List<OutputDataRecord>();

            mesParamCodeList = worksheet.Range("A:A").CellsUsed().Select(c => c.CachedValue.ToString()/*.Trim()*/).Skip(1).ToList();
            if (worksheet.Name.Trim().ToUpper() == "OUTPUTDATA")
            {
                outputDataList = worksheet.Range("A:B").RowsUsed()
                    .Select(c => new OutputDataRecord { MesParamCode = c.Cell(1).CachedValue.ToString(), ValueTime = c.Cell(2).CachedValue.ToString() }).Skip(1).ToList();
            }
            else
            {
                mesParamCodeList = worksheet.Range("A:A").CellsUsed().Select(c => c.CachedValue.ToString()/*.Trim()*/).Skip(1).ToList();
            }


            switch (checkReportTemplateTagsType)
            {
                case CheckReportTemplateTagsType.IsDuplicate:
                    {
                        List<string> duplicateList = new List<string>();
                        if (worksheet.Name.Trim().ToUpper() == "OUTPUTDATA")
                        {
                            duplicateList = outputDataList.GroupBy(u => new { u.MesParamCode, u.ValueTime }).Where(u => u.Count() > 1 && !String.IsNullOrEmpty(u.Key.ValueTime))
                                .Select(u => u.Key.MesParamCode + " на дату " + u.Key.ValueTime).ToList();
                        }
                        else
                        {
                            duplicateList = mesParamCodeList.GroupBy(u => u).Where(u => u.Count() > 1).Select(u => u.Key).ToList();
                        }
                        if (duplicateList != null && duplicateList.Any())
                        { return duplicateList; }
                        else
                        { return null; }
                    }
                case CheckReportTemplateTagsType.IsNotInBase:
                    {
                        mesParamDTOList = mesParamDTOList.Distinct().ToList();
                        List<string> notInBaseList = mesParamCodeList.Except(mesParamDTOList.Select(u => u.Code).ToList(), StringComparer.OrdinalIgnoreCase).ToList();
                        if (notInBaseList.Any())
                        { return notInBaseList; }
                        else
                        { return null; }
                    }
                case CheckReportTemplateTagsType.IsInArchive:
                    {
                        mesParamDTOList = mesParamDTOList.Distinct().ToList();
                        // выбор тех тэгов, которые в архиве и при этом есть в базе
                        List<string> isInArchiveList = (mesParamCodeList.Intersect(mesParamDTOList.Select(u => u.Code).ToList(), StringComparer.OrdinalIgnoreCase).ToList())
                            .Except(mesParamDTOList.Where(u => u.IsArchive != true).Select(u => u.Code).ToList(), StringComparer.OrdinalIgnoreCase).ToList();

                        if (isInArchiveList.Any())
                        { return isInArchiveList; }
                        else
                        { return null; }
                    }
                case CheckReportTemplateTagsType.IsInOtherNotArchiveReportTemplatesBySheetName:
                    {
                        mesParamCodeList = outputDataList.Select(u => u.MesParamCode).ToList();

                        var reportTemplateToMesParamList = await _reportTemplateToMesParamRepository.GetTagListInOtherNotArchiveReportTemplatesBySheetName(
                                reportTemplateId, worksheet.Name, worksheet.Name, mesParamCodeList, reportTemplateTypeName);

                        if (reportTemplateToMesParamList != null && reportTemplateToMesParamList.Any())
                        {
                            string templateStrId = "";
                            string templateStrType = "";
                            string templateStrDepartment = "";
                            string reportsStr = "";
                            List<string> listToReturn = new List<string>();
                            foreach (var item in reportTemplateToMesParamList.OrderBy(u => u.ReportTemplateId))
                            {
                                templateStrId = item.ReportTemplateId.ToString();
                                templateStrType = item.ReportTemplateDTOFK != null ? item.ReportTemplateDTOFK.ReportTemplateTypeDTOFK.Name : "";
                                templateStrDepartment = item.ReportTemplateDTOFK != null ? (item.ReportTemplateDTOFK.MesDepartmentDTOFK != null ? item.ReportTemplateDTOFK.MesDepartmentDTOFK.ToStringHierarchyShortName : "") : "";
                                reportsStr = "Тэг: " + item.MesParamCode + " присутствует на листе \"" + worksheet.Name + "\""
                                    + " Ид шаблона: " + templateStrId + " Тип: "
                                    + templateStrType + " Производство: "
                                    + templateStrDepartment + " ";
                                listToReturn.Add(reportsStr);

                            }
                            if (listToReturn.Any())
                            { return listToReturn; }
                            else
                            { return null; }
                        }
                        return null;
                    }
            }

            return null;
        }
    }
}
