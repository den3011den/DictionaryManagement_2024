using ClosedXML.Excel;
using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ICheckReportTemplateRepository
    {
        public Task<IEnumerable<string>?> IsNotExistSheets(IXLWorkbook workbook, List<string> sheetList);
        public Task<bool> IsEmptySheet(IXLWorksheet worksheet);
        public Task<List<SheetHeader>?> CheckSheetHeader(IXLWorksheet worksheet, List<SheetHeader>? sheetHeaderList);
        public Task<IEnumerable<string>?> CheckSheetTags(IXLWorksheet worksheet, IEnumerable<MesParamDTO> mesParamDTOList, CheckReportTemplateTagsType checkReportTemplateTagsType);
    }
}
