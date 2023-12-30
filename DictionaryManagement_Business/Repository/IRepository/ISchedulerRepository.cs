using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISchedulerRepository
    {

        public Task<SchedulerDTO> GetById(int id);
        public Task<IEnumerable<SchedulerDTO>> GetAll();
        public Task<SchedulerDTO> Update(SchedulerDTO objDTO);
        public Task<SchedulerDTO> Create(SchedulerDTO objectToAddDTO);
        public Task<int> Delete(int id);
    }
}
