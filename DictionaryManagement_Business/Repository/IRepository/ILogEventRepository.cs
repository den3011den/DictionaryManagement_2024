using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ILogEventRepository
    {
        public Task<IEnumerable<LogEventDTO>> GetAllByTimeInterval(DateTime? startEventTime, DateTime? endEventTime);
        public Task<LogEventDTO> AddRecord(string logEventTypeName, Guid userId, string oldValue, string newValue, bool isError, string description);
        public Task ToLog<T>(T? oldObject, T newObject, string logEventTypeName, string prefixString, IAuthorizationRepository _authorizationRepository);
    }
}
