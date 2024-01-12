using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class MesMovementsRepository : IMesMovementsRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public MesMovementsRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<MesMovementsDTO> Create(MesMovementsDTO objectToAddDTO)
        {

            MesMovements objectToAdd = new MesMovements();

            objectToAdd.Id = objectToAddDTO.Id;

            if (objectToAddDTO.AddTime == null)
                objectToAdd.AddTime = DateTime.Now;
            else
                objectToAdd.AddTime = objectToAddDTO.AddTime;

            objectToAdd.AddUserId = objectToAddDTO.AddUserId;
            objectToAdd.MesParamId = objectToAddDTO.MesParamId;
            objectToAdd.ValueTime = objectToAddDTO.ValueTime;
            objectToAdd.SapMovementOutId = objectToAddDTO.SapMovementOutId;
            objectToAdd.SapMovementInId = objectToAddDTO.SapMovementInId;
            objectToAdd.DataSourceId = objectToAddDTO.DataSourceId;
            objectToAdd.DataTypeId = objectToAddDTO.DataTypeId;
            objectToAdd.ReportGuid = objectToAddDTO.ReportGuid;
            objectToAdd.PreviousRecordId = objectToAddDTO.PreviousRecordId;
            objectToAdd.MesGone = objectToAddDTO.MesGone;
            objectToAdd.NeedWriteToSap = objectToAddDTO.NeedWriteToSap;
            objectToAdd.MesGoneTime = objectToAddDTO.MesGoneTime;

            var addedMesMovements = _db.MesMovements.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<MesMovements, MesMovementsDTO>(addedMesMovements.Entity);
        }


        public async Task<MesMovementsDTO> GetById(Guid id)
        {
            var objToGet = _db.MesMovements
                            .Include("AddUserFK")
                            .Include("MesParamFK")
                            .Include("SapMovementsOUTFK")
                            .Include("SapMovementsINFK")
                            .Include("DataSourceFK")
                            .Include("DataTypeFK")
                            .Include("ReportEntityFK")
                            .Include("MesMovementsFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<MesMovements, MesMovementsDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<MesMovementsDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode)
        {

            if (startTime == null)
                startTime = DateTime.MinValue;
            if (startTime == null)
                endTime = DateTime.MaxValue;

            switch (intervalMode.Trim().ToUpper())
            {
                case "ADDTIME":
                    var hhh1 = _db.MesMovements
                            .Include("AddUserFK")
                            .Include("MesParamFK")
                            .Include("SapMovementsOUTFK")
                            .Include("SapMovementsINFK")
                            .Include("DataSourceFK")
                            .Include("DataTypeFK")
                            .Include("ReportEntityFK")
                            .Include("MesMovementsFK")
                            .Include("MesMovementsCommentList").Include("MesMovementsCommentList.CorrectionReasonFK")
                        .Where(u => u.AddTime >= startTime && u.AddTime <= endTime).ToListWithNoLock();
                    return _mapper.Map<IEnumerable<MesMovements>, IEnumerable<MesMovementsDTO>>(hhh1);

                case "VALUETIME":
                    var hhh2 = _db.MesMovements
                            .Include("AddUserFK")
                            .Include("MesParamFK")
                            .Include("SapMovementsOUTFK")
                            .Include("SapMovementsINFK")
                            .Include("DataSourceFK")
                            .Include("DataTypeFK")
                            .Include("ReportEntityFK")
                            .Include("MesMovementsFK")
                            .Include("MesMovementsCommentList").Include("MesMovementsCommentList.CorrectionReasonFK")
                        .Where(u => u.ValueTime >= startTime && u.ValueTime <= endTime).ToListWithNoLock();
                    return _mapper.Map<IEnumerable<MesMovements>, IEnumerable<MesMovementsDTO>>(hhh2);
                default:
                    return null;
            }

        }


        public async Task<MesMovementsDTO> Update(MesMovementsDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.MesMovements
                            .Include("AddUserFK")
                            .Include("MesParamFK")
                            .Include("SapMovementsOUTFK")
                            .Include("SapMovementsINFK")
                            .Include("DataSourceFK")
                            .Include("DataTypeFK")
                            .Include("ReportEntityFK")
                            .Include("MesMovementsFK")
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
                        var objectMesParamToUpdate = _db.MesParam
                            .Include("MesParamSourceTypeFK")
                            .Include("MesDepartmentFK")
                            .Include("SapEquipmentSourceFK")
                            .Include("SapEquipmentDestFK")
                            .Include("MesMaterialFK")
                            .Include("SapMaterialFK")
                            .Include("MesUnitOfMeasureFK")
                            .Include("SapUnitOfMeasureFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesParamId);
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
                        var objectReportEntityToUpdate = _db.ReportEntity
                            .Include("ReportTemplateFK")
                            .Include("ReportDepartmentFK")
                            .Include("DownloadUserFK")
                            .Include("UploadUserFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.ReportGuid);
                        objectToUpdate.ReportEntityFK = objectReportEntityToUpdate;
                    }
                }

                if (objectToUpdateDTO.SapMovementOutId == null)
                {
                    objectToUpdate.SapMovementOutId = null;
                    objectToUpdate.SapMovementsOUTFK = null;
                }
                else
                {
                    if (objectToUpdate.SapMovementOutId != objectToUpdateDTO.SapMovementOutId)
                    {
                        objectToUpdate.SapMovementOutId = objectToUpdateDTO.SapMovementOutId;
                        var objectSapMovementsOUTToUpdate = _db.SapMovementsOUT
                                .Include("MesMovementsFK")
                                .Include("PreviousRecordFK")
                                .Include("MesParamFK")
                                .Include("SapEquipmentSourceFK")
                                .Include("SapEquipmentDestFK")
                                .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.SapMovementOutId);
                        objectToUpdate.SapMovementsOUTFK = objectSapMovementsOUTToUpdate;
                    }
                }

                if (objectToUpdateDTO.SapMovementInId == null)
                {
                    objectToUpdate.SapMovementInId = null;
                    objectToUpdate.SapMovementsINFK = null;
                }
                else
                {
                    if (objectToUpdate.SapMovementInId != objectToUpdateDTO.SapMovementInId)
                    {
                        objectToUpdate.SapMovementInId = objectToUpdateDTO.SapMovementInId;
                        var objectSapMovementsINToUpdate = _db.SapMovementsIN
                            .Include("MesMovementFK")
                            .Include("PreviousRecordFK")
                            .FirstOrDefaultWithNoLock(u => u.ErpId == objectToUpdateDTO.SapMovementInId);
                        objectToUpdate.SapMovementsINFK = objectSapMovementsINToUpdate;
                    }
                }

                if (objectToUpdateDTO.DataSourceId == null)
                {
                    objectToUpdate.DataSourceId = null;
                    objectToUpdate.DataSourceFK = null;
                }
                else
                {
                    if (objectToUpdate.DataSourceId != objectToUpdateDTO.DataSourceId)
                    {
                        objectToUpdate.DataSourceId = objectToUpdateDTO.DataSourceId;
                        var objectDataSourceToUpdate = _db.DataSource.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.DataSourceId);
                        objectToUpdate.DataSourceFK = objectDataSourceToUpdate;
                    }
                }

                if (objectToUpdateDTO.DataTypeId == null)
                {
                    objectToUpdate.DataTypeId = null;
                    objectToUpdate.DataTypeFK = null;
                }
                else
                {
                    if (objectToUpdate.DataTypeId != objectToUpdateDTO.DataTypeId)
                    {
                        objectToUpdate.DataTypeId = objectToUpdateDTO.DataTypeId;
                        var objectDataTypeToUpdate = _db.DataType.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.DataTypeId);
                        objectToUpdate.DataTypeFK = objectDataTypeToUpdate;
                    }
                }

                if (objectToUpdateDTO.PreviousRecordId == null)
                {
                    objectToUpdate.PreviousRecordId = null;
                    objectToUpdate.MesMovementsFK = null;
                }
                else
                {
                    if (objectToUpdate.PreviousRecordId != objectToUpdateDTO.PreviousRecordId)
                    {
                        objectToUpdate.PreviousRecordId = objectToUpdateDTO.PreviousRecordId;
                        var objectMesMovementsToUpdate = _db.MesMovements
                            .Include("AddUserFK")
                            .Include("MesParamFK")
                            .Include("SapMovementsOUTFK")
                            .Include("SapMovementsINFK")
                            .Include("DataSourceFK")
                            .Include("DataTypeFK")
                            .Include("ReportEntityFK")
                            .Include("MesMovementsFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.PreviousRecordId);
                        objectToUpdate.MesMovementsFK = objectMesMovementsToUpdate;
                    }
                }


                if (objectToUpdateDTO.AddTime != objectToUpdate.AddTime)
                    objectToUpdate.AddTime = objectToUpdateDTO.AddTime;

                if (objectToUpdateDTO.ValueTime != objectToUpdate.ValueTime)
                    objectToUpdate.ValueTime = objectToUpdateDTO.ValueTime;

                if (objectToUpdateDTO.Value != objectToUpdate.Value)
                    objectToUpdate.Value = objectToUpdateDTO.Value;

                if (objectToUpdateDTO.MesGone != objectToUpdate.MesGone)
                    objectToUpdate.MesGone = objectToUpdateDTO.MesGone;

                if (objectToUpdateDTO.MesGoneTime != objectToUpdate.MesGoneTime)
                    objectToUpdate.MesGoneTime = objectToUpdateDTO.MesGoneTime;

                if (objectToUpdateDTO.NeedWriteToSap != objectToUpdate.NeedWriteToSap)
                    objectToUpdate.NeedWriteToSap = objectToUpdateDTO.NeedWriteToSap;

                _db.MesMovements.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<MesMovements, MesMovementsDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(Guid id)
        {
            if (id != Guid.Empty)
            {
                var objectToDelete = _db.MesMovements.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.MesMovements.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;
        }


        public async Task<IEnumerable<MesMovementsDTO>> GetAllByReportEntityId(Guid? reportEntityId)
        {
            if (reportEntityId != null && reportEntityId != Guid.Empty)
            {
                var hhh1 = _db.MesMovements
                                    .Include("AddUserFK")
                                    .Include("MesParamFK")
                                    .Include("SapMovementsOUTFK")
                                    .Include("SapMovementsINFK")
                                    .Include("DataSourceFK")
                                    .Include("DataTypeFK")
                                    .Include("ReportEntityFK")
                                    .Include("MesMovementsFK")
                                    .Include("MesMovementsCommentList").Include("MesMovementsCommentList.CorrectionReasonFK")
                                .Where(u => u.ReportGuid == reportEntityId).ToListWithNoLock();
                return _mapper.Map<IEnumerable<MesMovements>, IEnumerable<MesMovementsDTO>>(hhh1);
            }
            else
                return new List<MesMovementsDTO>();
        }

    }
}
