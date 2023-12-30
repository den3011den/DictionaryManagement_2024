using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IVersionRepository
    {
        public Task<VersionDTO> Get();
        public Task<VersionDTO> Set(VersionDTO objectToAddDTO);
    }
}
