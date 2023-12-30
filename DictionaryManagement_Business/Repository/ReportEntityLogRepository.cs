using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class ReportEntityLogRepository : IReportEntityLogRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ReportEntityLogRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<ReportEntityLogDTO> Create(ReportEntityLogDTO objectToAddDTO)
        {
            ReportEntityLog objectToAdd = new ReportEntityLog();

            objectToAdd.Id = objectToAddDTO.Id;

            if (objectToAddDTO.LogTime == null)
                objectToAdd.LogTime = DateTime.Now;
            else
                objectToAdd.LogTime = objectToAddDTO.LogTime;

            objectToAdd.ReportEntityId = objectToAddDTO.ReportEntityId;
            objectToAdd.LogMessage = objectToAddDTO.LogMessage;
            objectToAdd.LogType = objectToAddDTO.LogType;
            objectToAdd.IsError = objectToAddDTO.IsError;

            var addedReportEntityLog = _db.ReportEntityLog.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<ReportEntityLog, ReportEntityLogDTO>(addedReportEntityLog.Entity);
        }


        public async Task<ReportEntityLogDTO> GetById(Int64 id)
        {
            var objToGet = _db.ReportEntityLog
                            .Include("ReportEntityFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<ReportEntityLog, ReportEntityLogDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<ReportEntityLogDTO>> GetAll()
        {
            var hhh1 = _db.ReportEntityLog
                            .Include("ReportEntityFK").ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntityLog>, IEnumerable<ReportEntityLogDTO>>(hhh1);
        }

        public async Task<IEnumerable<ReportEntityLogDTO>> GetByReportEntityId(Guid reportEntityId)
        {
            var hhh1 = _db.ReportEntityLog.Include("ReportEntityFK").Where(u => u.ReportEntityId == reportEntityId)
                .OrderBy(u => u.LogTime)
                .OrderBy(u => u.LogType)
                .ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntityLog>, IEnumerable<ReportEntityLogDTO>>(hhh1);
        }

        public async Task<IEnumerable<ReportEntityLogDTO>> GetAllByLogTimeInterval(DateTime? startLogTime, DateTime? endLogTime)
        {

            if (startLogTime == null)
                startLogTime = DateTime.MinValue;
            if (endLogTime == null)
                endLogTime = DateTime.MaxValue;

            var hhh1 = _db.ReportEntityLog
                            .Include("ReportEntityFK")
                            .Where(u => u.LogTime >= startLogTime && u.LogTime <= endLogTime).ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntityLog>, IEnumerable<ReportEntityLogDTO>>(hhh1);

        }

        public async Task<IEnumerable<ReportEntityLogDTO>> GetAllByReportEntityId(Guid reportEntityId)
        {

            var hhh1 = _db.ReportEntityLog
                            .Include("ReportEntityFK")
                            .Where(u => u.ReportEntityId == reportEntityId).ToListWithNoLock().OrderBy(u => u.LogTime);
            return _mapper.Map<IEnumerable<ReportEntityLog>, IEnumerable<ReportEntityLogDTO>>(hhh1);

        }

        public async Task<ReportEntityLogDTO> Update(ReportEntityLogDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.ReportEntityLog
                .Include("ReportEntityFK")
               .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.ReportEntityId == null || objectToUpdateDTO.ReportEntityId == Guid.Empty)
                {
                    objectToUpdate.ReportEntityId = Guid.Empty;
                    objectToUpdate.ReportEntityFK = null;
                }
                else
                {
                    if (objectToUpdate.ReportEntityId != objectToUpdateDTO.ReportEntityId)
                    {
                        objectToUpdate.ReportEntityId = objectToUpdateDTO.ReportEntityId;
                        var objectReportEntityToUpdate = _db.ReportEntity.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.ReportEntityId);
                        objectToUpdate.ReportEntityFK = objectReportEntityToUpdate;
                    }
                }

                if (objectToUpdateDTO.LogTime != objectToUpdate.LogTime)
                    objectToUpdate.LogTime = objectToUpdateDTO.LogTime;

                if (objectToUpdateDTO.LogMessage != objectToUpdate.LogMessage)
                    objectToUpdate.LogMessage = objectToUpdateDTO.LogMessage;

                if (objectToUpdateDTO.LogType != objectToUpdate.LogType)
                    objectToUpdate.LogType = objectToUpdateDTO.LogType;

                if (objectToUpdateDTO.IsError != objectToUpdate.IsError)
                    objectToUpdate.IsError = objectToUpdateDTO.IsError;

                _db.ReportEntityLog.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<ReportEntityLog, ReportEntityLogDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(Int64 id)
        {
            if (id > 0)
            {
                var objectToDelete = _db.ReportEntityLog.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.ReportEntityLog.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
