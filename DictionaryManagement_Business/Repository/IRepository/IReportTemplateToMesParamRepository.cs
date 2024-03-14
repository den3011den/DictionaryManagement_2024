using ClosedXML.Excel;
using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IReportTemplateToMesParamRepository
    {
        public Task<ReportTemplateToMesParamDTO?> GetById(int id);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>> GetAll();
        public Task<ReportTemplateToMesParamDTO?> Update(ReportTemplateToMesParamDTO objDTO);
        public Task<ReportTemplateToMesParamDTO?> Create(ReportTemplateToMesParamDTO objectToAddDTO);
        public Task<int> Delete(int Id);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByReportTemplateId(Guid reportTemplateTypeId);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByMesParamId(int mesParamId);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByMesParamCode(string mesParamCode);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetBySheetName(string sheetName);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByMesParamIdOrAndMesParamCode(int? mesParamId, string? mesParamCode);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetByReportTemplateIdAndMesParamIdAndSheetName(Guid reportTemplateTypeId, int mesParamId, string sheetName);
        public Task<IEnumerable<ReportTemplateToMesParamDTO>?> GetReportTemplateIdAndMesParamCodeAndSheetName(Guid reportTemplateTypeId, string mesParamCode, string sheetName);
        public Task<int> DeleteAllByReportTemplateId(Guid reportTemplateId);
        public Task<int> CreateByList(IEnumerable<ReportTemplateToMesParamDTO> listToAddDTO);
        public Task AddFromWorksheet(Guid reportTemplateId, IXLWorkbook workbook, string sheetName, string columnName, IEnumerable<MesParamDTO> mesParamListDTO);
    }
}
