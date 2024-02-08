using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IReportEntityResendDatesRepository
    {
        public Task<IEnumerable<ReportEntityResendDatesDTO>?> GetListByReportEntityId(Guid? reportEntityId);
    }
}
