using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISettingsRepository
    {
        public Task<SettingsDTO> Get(int Id);
        public Task<SettingsDTO> GetByName(string name);
        public Task<IEnumerable<SettingsDTO>> GetAll();
        public Task<SettingsDTO> Update(SettingsDTO objDTO);
        public Task<SettingsDTO> Create(SettingsDTO objectToAddDTO);
        public Task<int> Delete(int Id);
    }
}
