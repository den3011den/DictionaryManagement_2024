using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class SchedulerRepository : ISchedulerRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SchedulerRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SchedulerDTO> Create(SchedulerDTO objectToAddDTO)
        {

            Scheduler objectToAdd = new Scheduler();

            objectToAdd.ModuleName = objectToAddDTO.ModuleName;
            objectToAdd.StartTime = objectToAddDTO.StartTime;
            objectToAdd.LastExecuted = objectToAddDTO.LastExecuted;

            var addedScheduler = _db.Scheduler.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<Scheduler, SchedulerDTO>(addedScheduler.Entity);
        }


        public async Task<SchedulerDTO> GetById(int id)
        {
            var objToGet = _db.Scheduler
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<Scheduler, SchedulerDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<SchedulerDTO>> GetAll()
        {
            var hhh1 = _db.Scheduler.ToListWithNoLock();
            return _mapper.Map<IEnumerable<Scheduler>, IEnumerable<SchedulerDTO>>(hhh1);
        }


        public async Task<SchedulerDTO> Update(SchedulerDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.Scheduler.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.ModuleName != objectToUpdate.ModuleName)
                    objectToUpdate.ModuleName = objectToUpdateDTO.ModuleName;

                if (objectToUpdateDTO.StartTime != objectToUpdate.StartTime)
                    objectToUpdate.StartTime = objectToUpdateDTO.StartTime;

                if (objectToUpdateDTO.LastExecuted != objectToUpdate.LastExecuted)
                    objectToUpdate.LastExecuted = objectToUpdateDTO.LastExecuted;

                _db.Scheduler.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<Scheduler, SchedulerDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }
        public async Task<int> Delete(int id)
        {
            if (id != null && id != 0)
            {
                var objectToDelete = _db.Scheduler.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.Scheduler.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
