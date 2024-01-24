using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISapMovementsOUTRepository
    {
        public Task<SapMovementsOUTDTO> GetById(Guid id);
        public Task<IEnumerable<SapMovementsOUTDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode);
        public Task<SapMovementsOUTDTO> Update(SapMovementsOUTDTO objDTO);
        public Task<SapMovementsOUTDTO> Create(SapMovementsOUTDTO objectToAddDTO);
        public Task<int> Delete(Guid id);
        public Task<IEnumerable<SapMovementsOUTDTO>?> GetListByMesMovementId(Guid? idPar);
        public Task<SapMovementsOUTDTO?> CleanMesMovementId(SapMovementsOUTDTO objectToUpdateDTO);
        public Task<IEnumerable<SapMovementsOUTDTO>?> GetListByPreviousRecordId(Guid? idPar);
        public Task<SapMovementsOUTDTO?> CleanPreviousRecordId(SapMovementsOUTDTO objectToUpdateDTO);
    }
}
