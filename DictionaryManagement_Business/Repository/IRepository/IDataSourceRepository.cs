using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IDataSourceRepository
    {
        public Task<DataSourceDTO> Get(int Id);
        public Task<IEnumerable<DataSourceDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<DataSourceDTO> Update(DataSourceDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<DataSourceDTO> Create(DataSourceDTO objectToAddDTO);
        public Task<DataSourceDTO> GetByName(string Name);
    }
}
