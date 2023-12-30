using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISapMovementsINRepository
    {
        public Task<SapMovementsINDTO> GetById(string erpId);
        public Task<IEnumerable<SapMovementsINDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode);
        public Task<SapMovementsINDTO> Update(SapMovementsINDTO objDTO);
        public Task<SapMovementsINDTO> Create(SapMovementsINDTO objectToAddDTO);
        public Task<int> Delete(string erpId);
    }
}
