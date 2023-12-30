using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IReportTemplateRepository
    {
        public Task<ReportTemplateDTO> GetById(Guid id);
        public Task<IEnumerable<ReportTemplateDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<ReportTemplateDTO> Update(ReportTemplateDTO objDTO);
        public Task<ReportTemplateDTO> Create(ReportTemplateDTO objectToAddDTO);
        public Task<int> Delete(Guid id, UpdateMode updateMode = UpdateMode.Update);
        public Task<ReportTemplateDTO> GetByTemplateFileName(string templateFileName = "");
        public Task<ReportTemplateDTO> GetByReportTemplateTypeIdAndDestDataTypeIdAndDepartmentId(int reportTemplateTypeId, int destDataTypeId, int departmentId);
    }
}
