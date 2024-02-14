using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IReportTemplateFileHistoryRepository
    {
        public Task<IEnumerable<ReportTemplateFileHistoryDTO>?> GetListByReportEntityId(Guid? reportEntityId);
        public Task<ReportTemplateFileHistoryDTO> Create(ReportTemplateFileHistoryDTO objectToAddDTO);
    }
}
