using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;

namespace DictionaryManagement_Business.Repository
{
    public class SapMovementsINRepository : ISapMovementsINRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SapMovementsINRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SapMovementsINDTO> Create(SapMovementsINDTO objectToAddDTO)
        {

            SapMovementsIN objectToAdd = new SapMovementsIN();

            objectToAdd.ErpId = objectToAddDTO.ErpId;

            if (objectToAddDTO.AddTime == null)
                objectToAdd.AddTime = DateTime.Now;
            else
                objectToAdd.AddTime = objectToAddDTO.AddTime;

            objectToAdd.SapDocumentEnterTime = objectToAddDTO.SapDocumentEnterTime;
            objectToAdd.SapDocumentPostTime = objectToAddDTO.SapDocumentPostTime;
            objectToAdd.BatchNo = objectToAddDTO.BatchNo;
            objectToAdd.SapMaterialCode = objectToAddDTO.SapMaterialCode;
            objectToAdd.ErpPlantIdSource = objectToAddDTO.ErpPlantIdSource;
            objectToAdd.ErpIdSource = objectToAddDTO.ErpIdSource;

            objectToAdd.IsWarehouseSource = objectToAddDTO.IsWarehouseSource;

            objectToAdd.ErpPlantIdDest = objectToAddDTO.ErpPlantIdDest;
            objectToAdd.ErpIdDest = objectToAddDTO.ErpIdDest;
            objectToAdd.IsWarehouseDest = objectToAddDTO.IsWarehouseDest;
            objectToAdd.Value = objectToAddDTO.Value;
            objectToAdd.SapUnitOfMeasure = objectToAddDTO.SapUnitOfMeasure;
            objectToAdd.IsStorno = objectToAddDTO.IsStorno;
            objectToAdd.MesGone = objectToAddDTO.MesGone;

            objectToAdd.MesGoneTime = objectToAddDTO.MesGoneTime;
            objectToAdd.MesError = objectToAddDTO.MesError;
            objectToAdd.MesErrorMessage = objectToAddDTO.MesErrorMessage;
            objectToAdd.MesMovementId = objectToAddDTO.MesMovementId;
            objectToAdd.PreviousErpId = objectToAddDTO.PreviousErpId;
            objectToAdd.MoveType = objectToAddDTO.MoveType;


            var addedSapMovementsIN = _db.SapMovementsIN.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<SapMovementsIN, SapMovementsINDTO>(addedSapMovementsIN.Entity);
        }


        public async Task<SapMovementsINDTO> GetById(string erpId)
        {
            var objToGet = _db.SapMovementsIN
                            .Include("MesMovementFK")
                            .Include("PreviousRecordFK")
                            .FirstOrDefaultWithNoLock(u => u.ErpId.Trim().ToUpper() == erpId.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<SapMovementsIN, SapMovementsINDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<SapMovementsINDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode)
        {

            if (startTime == null)
                startTime = DateTime.MinValue;
            if (startTime == null)
                endTime = DateTime.MaxValue;

            IEnumerable<SapMovementsIN>? draftDbSapMovementsINSelection;
            IEnumerable<SapMovementsIN>? dbSapMovementsINSelection;

            switch (intervalMode.Trim().ToUpper())
            {
                case "ADDTIME":
                    draftDbSapMovementsINSelection = _db.SapMovementsIN
                            .Include("PreviousRecordFK")
                            .Include("MesMovementFK").Include("MesMovementFK.MesParamFK")
                        .Where(u => u.AddTime >= startTime && u.AddTime <= endTime).AsNoTracking().ToListWithNoLock();
                    break;
                case "VALUETIME":
                    draftDbSapMovementsINSelection = _db.SapMovementsIN
                            .Include("PreviousRecordFK")
                            .Include("MesMovementFK").Include("MesMovementFK.MesParamFK")
                        .Where(u => u.SapDocumentPostTime >= startTime && u.SapDocumentPostTime <= endTime).AsNoTracking().ToListWithNoLock();
                    break;
                default:
                    return null;
            }

            if (draftDbSapMovementsINSelection != null)
            {
                dbSapMovementsINSelection = (from draftDbSapMovementsINSelectionAlias in draftDbSapMovementsINSelection
                                             join SM_prom1 in _db.SapMaterial.AsNoTracking().ToListWithNoLock() on
                                                 draftDbSapMovementsINSelectionAlias.SapMaterialCode equals SM_prom1.Code
                                             into SM_prom2
                                             from SM in SM_prom2.DefaultIfEmpty()
                                             join SESource_prom1 in _db.SapEquipment.AsNoTracking().ToListWithNoLock() on new
                                             {
                                                 ErpPlantIdSource = draftDbSapMovementsINSelectionAlias.ErpPlantIdSource.ToString(),
                                                 ErpIdSource = draftDbSapMovementsINSelectionAlias.ErpIdSource.ToString()
                                             }
                                                 equals new
                                                 {
                                                     ErpPlantIdSource = SESource_prom1.ErpPlantId.ToString(),
                                                     ErpIdSource = SESource_prom1.ErpId.ToString()
                                                 }
                                             into SESource_prom2
                                             from SESource in SESource_prom2.DefaultIfEmpty()
                                             join SEDest_prom1 in _db.SapEquipment.AsNoTracking().ToListWithNoLock() on new
                                             {
                                                 ErpPlantIdSource = draftDbSapMovementsINSelectionAlias.ErpPlantIdDest.ToString(),
                                                 ErpIdSource = draftDbSapMovementsINSelectionAlias.ErpIdDest.ToString()
                                             }
                                              equals new
                                              {
                                                  ErpPlantIdSource = SEDest_prom1.ErpPlantId.ToString(),
                                                  ErpIdSource = SEDest_prom1.ErpId.ToString()
                                              }
                                             into SEDest_prom2
                                             from SEDest in SEDest_prom2.DefaultIfEmpty()
                                             join SUOM_prom1 in _db.SapUnitOfMeasure.AsNoTracking().ToListWithNoLock() on
                                                    draftDbSapMovementsINSelectionAlias.SapUnitOfMeasure.Trim().ToUpper() equals SUOM_prom1.ShortName.Trim().ToUpper()
                                             into SUOM_prom2
                                             from SUOM in SUOM_prom2.DefaultIfEmpty()
                                             select
                                                          new SapMovementsIN
                                                          {
                                                              ErpId = draftDbSapMovementsINSelectionAlias.ErpId,
                                                              AddTime = draftDbSapMovementsINSelectionAlias.AddTime,
                                                              SapDocumentEnterTime = draftDbSapMovementsINSelectionAlias.SapDocumentEnterTime,
                                                              SapDocumentPostTime = draftDbSapMovementsINSelectionAlias.SapDocumentPostTime,
                                                              BatchNo = draftDbSapMovementsINSelectionAlias.BatchNo,
                                                              SapMaterialCode = draftDbSapMovementsINSelectionAlias.SapMaterialCode,
                                                              SapMaterialFK = SM,
                                                              ErpPlantIdSource = draftDbSapMovementsINSelectionAlias.ErpPlantIdSource,
                                                              ErpIdSource = draftDbSapMovementsINSelectionAlias.ErpIdSource,
                                                              SapEquipmentSourceFK = SESource,
                                                              IsWarehouseSource = draftDbSapMovementsINSelectionAlias.IsWarehouseSource,
                                                              ErpPlantIdDest = draftDbSapMovementsINSelectionAlias.ErpPlantIdDest,
                                                              ErpIdDest = draftDbSapMovementsINSelectionAlias.ErpIdDest,
                                                              SapEquipmentDestFK = SEDest,
                                                              IsWarehouseDest = draftDbSapMovementsINSelectionAlias.IsWarehouseDest,
                                                              Value = draftDbSapMovementsINSelectionAlias.Value,
                                                              SapUnitOfMeasure = draftDbSapMovementsINSelectionAlias.SapUnitOfMeasure,
                                                              IsStorno = draftDbSapMovementsINSelectionAlias.IsStorno,
                                                              MesError = draftDbSapMovementsINSelectionAlias.MesError,
                                                              MesErrorMessage = draftDbSapMovementsINSelectionAlias.MesErrorMessage,
                                                              MesMovementId = draftDbSapMovementsINSelectionAlias.MesMovementId,
                                                              MesMovementFK = draftDbSapMovementsINSelectionAlias.MesMovementFK,
                                                              PreviousErpId = draftDbSapMovementsINSelectionAlias.PreviousErpId,
                                                              PreviousRecordFK = draftDbSapMovementsINSelectionAlias.PreviousRecordFK,
                                                              MoveType = draftDbSapMovementsINSelectionAlias.MoveType,
                                                              MesGone = draftDbSapMovementsINSelectionAlias.MesGone,
                                                              MesGoneTime = draftDbSapMovementsINSelectionAlias.MesGoneTime
                                                          }).ToList();
            }
            else
            {
                dbSapMovementsINSelection = new List<SapMovementsIN>();
            }
            return _mapper.Map<IEnumerable<SapMovementsIN>, IEnumerable<SapMovementsINDTO>>(dbSapMovementsINSelection);

        }


        public async Task<SapMovementsINDTO> Update(SapMovementsINDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.SapMovementsIN
                        .Include("MesMovementFK")
                        .Include("PreviousRecordFK")
               .FirstOrDefaultWithNoLock(u => u.ErpId.Trim().ToUpper() == objectToUpdateDTO.ErpId.Trim().ToUpper());

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.MesMovementId == null)
                {
                    objectToUpdate.MesMovementId = null;
                    objectToUpdate.MesMovementFK = null;
                }
                else
                {
                    if (objectToUpdate.MesMovementId != objectToUpdateDTO.MesMovementId)
                    {
                        objectToUpdate.MesMovementId = objectToUpdateDTO.MesMovementId;
                        var objectMesMovementsToUpdate = _db.MesMovements
                            .Include("AddUserFK")
                            .Include("MesParamFK")
                            .Include("SapMovementsOUTFK")
                            .Include("SapMovementsINFK")
                            .Include("DataSourceFK")
                            .Include("DataTypeFK")
                            .Include("ReportEntityFK")
                            .Include("MesMovementsFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesMovementId);
                        objectToUpdate.MesMovementFK = objectMesMovementsToUpdate;
                    }
                }

                if (objectToUpdateDTO.PreviousErpId == null)
                {
                    objectToUpdate.PreviousErpId = null;
                    objectToUpdate.PreviousRecordFK = null;
                }
                else
                {
                    if (objectToUpdate.PreviousErpId != objectToUpdateDTO.PreviousErpId)
                    {
                        objectToUpdate.PreviousErpId = objectToUpdateDTO.PreviousErpId;
                        var objectSapMovementsINToUpdate = _db.SapMovementsIN
                                .Include("MesMovementFK")
                                .Include("PreviousRecordFK")
                                .FirstOrDefaultWithNoLock(u => u.ErpId.Trim().ToUpper() == objectToUpdateDTO.PreviousErpId.Trim().ToUpper());
                        objectToUpdate.PreviousRecordFK = objectSapMovementsINToUpdate;
                    }
                }


                if (objectToUpdateDTO.AddTime != objectToUpdate.AddTime)
                    objectToUpdate.AddTime = objectToUpdateDTO.AddTime;

                if (objectToUpdateDTO.SapDocumentEnterTime != objectToUpdate.SapDocumentEnterTime)
                    objectToUpdate.SapDocumentEnterTime = objectToUpdateDTO.SapDocumentEnterTime;

                if (objectToUpdateDTO.SapDocumentPostTime != objectToUpdate.SapDocumentPostTime)
                    objectToUpdate.SapDocumentPostTime = objectToUpdateDTO.SapDocumentPostTime;

                if (objectToUpdateDTO.BatchNo != objectToUpdate.BatchNo)
                    objectToUpdate.BatchNo = objectToUpdateDTO.BatchNo;

                if (objectToUpdateDTO.SapMaterialCode != objectToUpdate.SapMaterialCode)
                    objectToUpdate.SapMaterialCode = objectToUpdateDTO.SapMaterialCode;

                if (objectToUpdateDTO.ErpPlantIdSource != objectToUpdate.ErpPlantIdSource)
                    objectToUpdate.ErpPlantIdSource = objectToUpdateDTO.ErpPlantIdSource;

                if (objectToUpdateDTO.ErpIdSource != objectToUpdate.ErpIdSource)
                    objectToUpdate.ErpIdSource = objectToUpdateDTO.ErpIdSource;

                if (objectToUpdateDTO.IsWarehouseSource != objectToUpdate.IsWarehouseSource)
                    objectToUpdate.IsWarehouseSource = objectToUpdateDTO.IsWarehouseSource;

                if (objectToUpdateDTO.ErpPlantIdDest != objectToUpdate.ErpPlantIdDest)
                    objectToUpdate.ErpPlantIdDest = objectToUpdateDTO.ErpPlantIdDest;

                if (objectToUpdateDTO.ErpIdDest != objectToUpdate.ErpIdDest)
                    objectToUpdate.ErpIdDest = objectToUpdateDTO.ErpIdDest;

                if (objectToUpdateDTO.IsWarehouseDest != objectToUpdate.IsWarehouseDest)
                    objectToUpdate.IsWarehouseDest = objectToUpdateDTO.IsWarehouseDest;

                if (objectToUpdateDTO.Value != objectToUpdate.Value)
                    objectToUpdate.Value = objectToUpdateDTO.Value;

                if (objectToUpdateDTO.SapUnitOfMeasure != objectToUpdate.SapUnitOfMeasure)
                    objectToUpdate.SapUnitOfMeasure = objectToUpdateDTO.SapUnitOfMeasure;

                if (objectToUpdateDTO.IsStorno != objectToUpdate.IsStorno)
                    objectToUpdate.IsStorno = objectToUpdateDTO.IsStorno;

                if (objectToUpdateDTO.MesGone != objectToUpdate.MesGone)
                    objectToUpdate.MesGone = objectToUpdateDTO.MesGone;

                if (objectToUpdateDTO.MesGoneTime != objectToUpdate.MesGoneTime)
                    objectToUpdate.MesGoneTime = objectToUpdateDTO.MesGoneTime;

                if (objectToUpdateDTO.MesError != objectToUpdate.MesError)
                    objectToUpdate.MesError = objectToUpdateDTO.MesError;

                if (objectToUpdateDTO.MesErrorMessage != objectToUpdate.MesErrorMessage)
                    objectToUpdate.MesErrorMessage = objectToUpdateDTO.MesErrorMessage;

                if (objectToUpdateDTO.MoveType != objectToUpdate.MoveType)
                    objectToUpdate.MoveType = objectToUpdateDTO.MoveType;

                _db.SapMovementsIN.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapMovementsIN, SapMovementsINDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(string id)
        {
            if (id.Trim().IsNullOrEmpty())
            {
                var objectToDelete = _db.SapMovementsIN.FirstOrDefaultWithNoLock(u => u.ErpId.Trim().ToUpper() == id.Trim().ToUpper());
                if (objectToDelete != null)
                {
                    _db.SapMovementsIN.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;
        }
    }
}
