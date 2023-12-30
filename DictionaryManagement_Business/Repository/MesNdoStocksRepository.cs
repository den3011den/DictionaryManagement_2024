using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class MesNdoStocksRepository : IMesNdoStocksRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public MesNdoStocksRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<MesNdoStocksDTO> Create(MesNdoStocksDTO objectToAddDTO)
        {

            MesNdoStocks objectToAdd = new MesNdoStocks();

            objectToAdd.Id = objectToAddDTO.Id;

            if (objectToAddDTO.AddTime == null)
                objectToAdd.AddTime = DateTime.Now;
            else
                objectToAdd.AddTime = objectToAddDTO.AddTime;

            objectToAdd.ValueTime = objectToAddDTO.ValueTime;
            objectToAdd.Value = objectToAddDTO.Value;
            objectToAdd.AddUserId = objectToAddDTO.AddUserId;
            objectToAdd.MesParamId = objectToAddDTO.MesParamId;
            objectToAdd.ValueDifference = objectToAddDTO.ValueDifference;
            objectToAdd.ReportGuid = objectToAddDTO.ReportGuid;
            objectToAdd.SapNdoOutId = objectToAddDTO.SapNdoOutId;

            var addedMesNdoStocks = _db.MesNdoStocks.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<MesNdoStocks, MesNdoStocksDTO>(addedMesNdoStocks.Entity);
        }


        public async Task<MesNdoStocksDTO> GetById(Int64 id)
        {
            var objToGet = _db.MesNdoStocks
                            .Include("MesParamFK")
                            .Include("AddUserFK")
                            .Include("ReportEntityFK")
                            .Include("SapNdoOUTFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<MesNdoStocks, MesNdoStocksDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<MesNdoStocksDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode)
        {

            if (startTime == null)
                startTime = DateTime.MinValue;
            if (startTime == null)
                endTime = DateTime.MaxValue;

            switch (intervalMode.Trim().ToUpper())
            {
                case "ADDTIME":
                    var hhh1 = _db.MesNdoStocks
                        .Include("MesParamFK")
                        .Include("AddUserFK")
                        .Include("ReportEntityFK")
                        .Include("SapNdoOUTFK")
                        .Where(u => u.AddTime >= startTime && u.AddTime <= endTime).ToListWithNoLock();
                    return _mapper.Map<IEnumerable<MesNdoStocks>, IEnumerable<MesNdoStocksDTO>>(hhh1);

                case "VALUETIME":
                    var hhh2 = _db.MesNdoStocks
                        .Include("MesParamFK")
                        .Include("AddUserFK")
                        .Include("ReportEntityFK")
                        .Include("SapNdoOUTFK")
                        .Where(u => u.ValueTime >= startTime && u.ValueTime <= endTime).ToListWithNoLock();
                    return _mapper.Map<IEnumerable<MesNdoStocks>, IEnumerable<MesNdoStocksDTO>>(hhh2);
                default:
                    return null;
            }

        }

        public async Task<IEnumerable<MesNdoStocksDTO>> GetAllByReportEntityId(Guid? reportEntityId)
        {

            if (reportEntityId != null && reportEntityId != Guid.Empty)
            {
                var hhh1 = _db.MesNdoStocks
                    .Include("MesParamFK")
                    .Include("AddUserFK")
                    .Include("ReportEntityFK")
                    .Include("SapNdoOUTFK")
                    .Where(u => u.ReportGuid == reportEntityId).ToListWithNoLock();
                return _mapper.Map<IEnumerable<MesNdoStocks>, IEnumerable<MesNdoStocksDTO>>(hhh1);
            }
            else
            {
                return new List<MesNdoStocksDTO>();
            }

        }


        public async Task<MesNdoStocksDTO> Update(MesNdoStocksDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.MesNdoStocks
                        .Include("MesParamFK")
                        .Include("AddUserFK")
                        .Include("ReportEntityFK")
                        .Include("SapNdoOUTFK")
               .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.MesParamId == null)
                {
                    objectToUpdate.MesParamId = 0;
                    objectToUpdate.MesParamFK = null;
                }
                else
                {
                    if (objectToUpdate.MesParamId != objectToUpdateDTO.MesParamId)
                    {
                        objectToUpdate.MesParamId = objectToUpdateDTO.MesParamId;
                        var objectMesParamToUpdate = _db.MesParam.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesParamId);
                        objectToUpdate.MesParamFK = objectMesParamToUpdate;
                    }
                }

                if (objectToUpdateDTO.AddUserId == null)
                {
                    objectToUpdate.AddUserId = Guid.Empty;
                    objectToUpdate.AddUserFK = null;
                }
                else
                {
                    if (objectToUpdate.AddUserId != objectToUpdateDTO.AddUserId)
                    {
                        objectToUpdate.AddUserId = objectToUpdateDTO.AddUserId;
                        var objectAddUserToUpdate = _db.User.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.AddUserId);
                        objectToUpdate.AddUserFK = objectAddUserToUpdate;
                    }
                }

                if (objectToUpdateDTO.ReportGuid == null)
                {
                    objectToUpdate.ReportGuid = Guid.Empty;
                    objectToUpdate.ReportEntityFK = null;
                }
                else
                {
                    if (objectToUpdate.ReportGuid != objectToUpdateDTO.ReportGuid)
                    {
                        objectToUpdate.ReportGuid = objectToUpdateDTO.ReportGuid;
                        var objectReportEntityToUpdate = _db.ReportEntity.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.ReportGuid);
                        objectToUpdate.ReportEntityFK = objectReportEntityToUpdate;
                    }
                }

                if (objectToUpdateDTO.SapNdoOutId == null)
                {
                    objectToUpdate.SapNdoOutId = null;
                    objectToUpdate.SapNdoOUTFK = null;
                }
                else
                {
                    if (objectToUpdate.SapNdoOutId != objectToUpdateDTO.SapNdoOutId)
                    {
                        objectToUpdate.SapNdoOutId = objectToUpdateDTO.SapNdoOutId;
                        var objectSapNdoOutToUpdate = _db.SapNdoOUT.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.SapNdoOutId);
                        objectToUpdate.SapNdoOUTFK = objectSapNdoOutToUpdate;
                    }
                }

                if (objectToUpdateDTO.AddTime != objectToUpdate.AddTime)
                    objectToUpdate.AddTime = objectToUpdateDTO.AddTime;

                if (objectToUpdateDTO.ValueTime != objectToUpdate.ValueTime)
                    objectToUpdate.ValueTime = objectToUpdateDTO.ValueTime;

                if (objectToUpdateDTO.Value != objectToUpdate.Value)
                    objectToUpdate.Value = objectToUpdateDTO.Value;

                if (objectToUpdateDTO.ValueDifference != objectToUpdate.ValueDifference)
                    objectToUpdate.ValueDifference = objectToUpdateDTO.ValueDifference;

                _db.MesNdoStocks.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<MesNdoStocks, MesNdoStocksDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<Int64> Delete(Int64 id)
        {
            if (id != null && id != 0)
            {
                var objectToDelete = _db.MesNdoStocks.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.MesNdoStocks.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;
        }
    }
}
