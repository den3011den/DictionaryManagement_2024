using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IMesNdoStocksRepository
    {
        public Task<MesNdoStocksDTO> GetById(Int64 id);
        public Task<IEnumerable<MesNdoStocksDTO>> GetAllByTimeInterval(DateTime? startDownloadTime, DateTime? endDownloadTime, string intervalMode);
        public Task<IEnumerable<MesNdoStocksDTO>> GetAllByReportEntityId(Guid? reportEntityId);
        public Task<MesNdoStocksDTO> Update(MesNdoStocksDTO objDTO);
        public Task<MesNdoStocksDTO> Create(MesNdoStocksDTO objectToAddDTO);
        public Task<Int64> Delete(Int64 id);
        public Task<IEnumerable<MesNdoStocksDTO>?> GetBySapNdoOutIdList(Int64 id);
        public Task<MesNdoStocksDTO?> CleanSapNdoOutId(MesNdoStocksDTO objectToUpdateDTO);
    }
}
