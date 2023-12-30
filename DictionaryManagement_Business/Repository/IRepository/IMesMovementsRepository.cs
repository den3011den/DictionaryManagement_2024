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
    }
}
