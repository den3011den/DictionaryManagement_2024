using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ILogEventTypeRepository
    {
        public Task<LogEventTypeDTO> Get(int Id);
        public Task<IEnumerable<LogEventTypeDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<LogEventTypeDTO> Update(LogEventTypeDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<LogEventTypeDTO> Create(LogEventTypeDTO objectToAddDTO);
        public Task<LogEventTypeDTO> GetByName(string Name);
    }
}
