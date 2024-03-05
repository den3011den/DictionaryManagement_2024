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

        public CheckReportTemplateRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
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

        public async Task<List<string>?> CheckSheetTags(IXLWorksheet worksheet, IEnumerable<MesParamDTO> mesParamDTOList, CheckReportTemplateTagsType checkReportTemplateTagsType)
        {
            var mesParamCodeList = worksheet.Range("A:A").CellsUsed().Select(c => c.CachedValue.ToString()/*.Trim()*/).Skip(1).ToList();
            switch (checkReportTemplateTagsType)
            {
                case CheckReportTemplateTagsType.IsDuplicate:
                    {
                        List<string> duplicateList = mesParamCodeList.GroupBy(u => u).Where(u => u.Count() > 1).Select(u => u.Key).ToList();
                        if (duplicateList.Any())
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
            }

            return null;
        }
    }
}
