using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISapNdoOUTRepository
    {
        public Task<SapNdoOUTDTO> GetById(Int64 id);
        public Task<IEnumerable<SapNdoOUTDTO>> GetAllByTimeInterval(DateTime? startDownloadTime, DateTime? endDownloadTime, string intervalMode);
        public Task<SapNdoOUTDTO> Update(SapNdoOUTDTO objDTO);
        public Task<SapNdoOUTDTO> Create(SapNdoOUTDTO objectToAddDTO);
        public Task<Int64> Delete(Int64 id);
    }
}
