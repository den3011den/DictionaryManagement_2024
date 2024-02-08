using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IMesMovementsRepository
    {
        public Task<MesMovementsDTO> GetById(Guid id);
        public Task<IEnumerable<MesMovementsDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode);
        public Task<MesMovementsDTO> Update(MesMovementsDTO objDTO);
        public Task<MesMovementsDTO> Create(MesMovementsDTO objectToAddDTO);
        public Task<int> Delete(Guid id);
        public Task<IEnumerable<MesMovementsDTO>> GetAllByReportEntityId(Guid? reportEntityId);
        public Task<IEnumerable<MesMovementsDTO>?> GetListByPreviousRecordId(Guid? idPar);
        public Task<MesMovementsDTO?> CleanPreviousRecordId(MesMovementsDTO objectToUpdateDTO);
        public Task<int> DeleteMesMovementsCommentByMesMovementsId(Guid objectId);
        public Task<IEnumerable<MesMovementsDTO>?> GetListBySapMovementOutId(Guid? idPar);
        public Task<IEnumerable<MesMovementsDTO>?> GetListBySapMovementInId(string? idPar);
        public Task<MesMovementsDTO?> CleanSapMovementInId(MesMovementsDTO objectToUpdateDTO);
        public Task<MesMovementsDTO?> CleanSapMovementOutId(MesMovementsDTO objectToUpdateDTO);

    }
}
