using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;


namespace DictionaryManagement_Business.Repository
{
    public class SapMovementsOUTRepository : ISapMovementsOUTRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SapMovementsOUTRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SapMovementsOUTDTO> Create(SapMovementsOUTDTO objectToAddDTO)
        {

            SapMovementsOUT objectToAdd = new SapMovementsOUT();

            objectToAdd.Id = objectToAddDTO.Id;

            if (objectToAddDTO.AddTime == null)
                objectToAdd.AddTime = DateTime.Now;
            else
                objectToAdd.AddTime = objectToAddDTO.AddTime;

            objectToAdd.BatchNo = objectToAddDTO.BatchNo;
            objectToAdd.SapMaterialCode = objectToAddDTO.SapMaterialCode;
            objectToAdd.ErpPlantIdSource = objectToAddDTO.ErpPlantIdSource;
            objectToAdd.ErpIdSource = objectToAddDTO.ErpIdSource;
            objectToAdd.IsWarehouseSource = objectToAddDTO.IsWarehouseSource;
            objectToAdd.ErpPlantIdDest = objectToAddDTO.ErpPlantIdDest;
            objectToAdd.ErpIdDest = objectToAddDTO.ErpIdDest;
            objectToAdd.IsWarehouseDest = objectToAddDTO.IsWarehouseDest;
            objectToAdd.ValueTime = objectToAddDTO.ValueTime;
            objectToAdd.Value = objectToAddDTO.Value;
            objectToAdd.Correction2Previous = objectToAddDTO.Correction2Previous;
            objectToAdd.MesParamId = objectToAddDTO.MesParamId;
            objectToAdd.IsReconciled = objectToAddDTO.IsReconciled;
            objectToAdd.SapUnitOfMeasure = objectToAddDTO.SapUnitOfMeasure;
            objectToAdd.SapGone = objectToAddDTO.SapGone;
            objectToAdd.SapGoneTime = objectToAddDTO.SapGoneTime;
            objectToAdd.SapError = objectToAddDTO.SapError;
            objectToAdd.SapErrorMessage = objectToAddDTO.SapErrorMessage;
            objectToAdd.MesMovementId = objectToAddDTO.MesMovementId;
            objectToAdd.PreviousRecordId = objectToAddDTO.PreviousRecordId;
            objectToAdd.MesParamId = objectToAddDTO.MesParamId;

            var addedSapMovementsOUT = _db.SapMovementsOUT.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<SapMovementsOUT, SapMovementsOUTDTO>(addedSapMovementsOUT.Entity);
        }


        public async Task<SapMovementsOUTDTO> GetById(Guid id)
        {
            var objToGet = _db.SapMovementsOUT
                            .Include("MesMovementsFK")
                            .Include("PreviousRecordFK")
                            .Include("MesParamFK")
                            //.Include("SapEquipmentSourceFK")
                            //.Include("SapEquipmentDestFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<SapMovementsOUT, SapMovementsOUTDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<SapMovementsOUTDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode)
        {

            if (startTime == null)
                startTime = DateTime.MinValue;
            if (startTime == null)
                endTime = DateTime.MaxValue;

            IEnumerable<SapMovementsOUT>? draftDbSapMovementsOUTSelection;
            IEnumerable<SapMovementsOUT>? dbSapMovementsOUTSelection;

            switch (intervalMode.Trim().ToUpper())
            {
                case "ADDTIME":
                    draftDbSapMovementsOUTSelection = _db.SapMovementsOUT
                            .Include("MesMovementsFK")
                            .Include("PreviousRecordFK")
                            .Include("MesParamFK")
                        .Where(u => u.AddTime >= startTime && u.AddTime <= endTime).AsNoTracking().ToListWithNoLock();
                    break;
                case "VALUETIME":
                    draftDbSapMovementsOUTSelection = _db.SapMovementsOUT
                        .Include("MesMovementsFK")
                        .Include("PreviousRecordFK")
                        .Include("MesParamFK")
                        .Where(u => u.ValueTime >= startTime && u.ValueTime <= endTime).AsNoTracking().ToListWithNoLock();
                    break;
                default:
                    return null;
            }

            if (draftDbSapMovementsOUTSelection != null)
            {
                dbSapMovementsOUTSelection = (from draftDbSapMovementsOUTSelectionAlias in draftDbSapMovementsOUTSelection
                                              join SM_prom1 in _db.SapMaterial.AsNoTracking().ToListWithNoLock() on
                                                  draftDbSapMovementsOUTSelectionAlias.SapMaterialCode equals SM_prom1.Code
                                              into SM_prom2
                                              from SM in SM_prom2.DefaultIfEmpty()
                                              join SESource_prom1 in _db.SapEquipment.AsNoTracking().ToListWithNoLock() on new
                                              {
                                                  ErpPlantIdSource = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdSource == null ? "Пусто" : draftDbSapMovementsOUTSelectionAlias.ErpPlantIdSource.ToString(),
                                                  ErpIdSource = draftDbSapMovementsOUTSelectionAlias.ErpIdSource == null ? "Пусто" : draftDbSapMovementsOUTSelectionAlias.ErpIdSource.ToString()
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
                                                  ErpPlantIdSource = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdDest == null ? "Пусто" : draftDbSapMovementsOUTSelectionAlias.ErpPlantIdDest.ToString(),
                                                  ErpIdSource = draftDbSapMovementsOUTSelectionAlias.ErpIdDest == null ? "Пусто" : draftDbSapMovementsOUTSelectionAlias.ErpIdDest.ToString()
                                              }
                                               equals new
                                               {
                                                   ErpPlantIdSource = SEDest_prom1.ErpPlantId.ToString(),
                                                   ErpIdSource = SEDest_prom1.ErpId.ToString()
                                               }
                                              into SEDest_prom2
                                              from SEDest in SEDest_prom2.DefaultIfEmpty()
                                              join SUOM_prom1 in _db.SapUnitOfMeasure.AsNoTracking().ToListWithNoLock() on
                                                     draftDbSapMovementsOUTSelectionAlias.SapUnitOfMeasure.Trim().ToUpper() equals SUOM_prom1.ShortName.Trim().ToUpper()
                                              into SUOM_prom2
                                              from SUOM in SUOM_prom2.DefaultIfEmpty()

                                              select
                                                          new SapMovementsOUT
                                                          {
                                                              Id = draftDbSapMovementsOUTSelectionAlias.Id,
                                                              AddTime = draftDbSapMovementsOUTSelectionAlias.AddTime,
                                                              BatchNo = draftDbSapMovementsOUTSelectionAlias.BatchNo,
                                                              SapMaterialCode = draftDbSapMovementsOUTSelectionAlias.SapMaterialCode,
                                                              SapMaterialFK = SM,
                                                              ErpPlantIdSource = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdSource,
                                                              ErpIdSource = draftDbSapMovementsOUTSelectionAlias.ErpIdSource,
                                                              SapEquipmentSourceFK = SESource,
                                                              IsWarehouseSource = draftDbSapMovementsOUTSelectionAlias.IsWarehouseSource,
                                                              ErpPlantIdDest = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdDest,
                                                              ErpIdDest = draftDbSapMovementsOUTSelectionAlias.ErpIdDest,
                                                              SapEquipmentDestFK = SEDest,
                                                              IsWarehouseDest = draftDbSapMovementsOUTSelectionAlias.IsWarehouseDest,
                                                              ValueTime = draftDbSapMovementsOUTSelectionAlias.ValueTime,
                                                              Value = draftDbSapMovementsOUTSelectionAlias.Value,
                                                              Correction2Previous = draftDbSapMovementsOUTSelectionAlias.Correction2Previous,
                                                              IsReconciled = draftDbSapMovementsOUTSelectionAlias.IsReconciled,
                                                              SapUnitOfMeasure = draftDbSapMovementsOUTSelectionAlias.SapUnitOfMeasure,
                                                              SapGone = draftDbSapMovementsOUTSelectionAlias.SapGone,
                                                              SapGoneTime = draftDbSapMovementsOUTSelectionAlias.SapGoneTime,
                                                              SapError = draftDbSapMovementsOUTSelectionAlias.SapError,
                                                              SapErrorMessage = draftDbSapMovementsOUTSelectionAlias.SapErrorMessage,
                                                              MesMovementId = draftDbSapMovementsOUTSelectionAlias.MesMovementId,
                                                              MesMovementsFK = draftDbSapMovementsOUTSelectionAlias.MesMovementsFK,
                                                              MesParamId = draftDbSapMovementsOUTSelectionAlias.MesParamId,
                                                              MesParamFK = draftDbSapMovementsOUTSelectionAlias.MesParamFK,
                                                              PreviousRecordId = draftDbSapMovementsOUTSelectionAlias.PreviousRecordId,
                                                              PreviousRecordFK = draftDbSapMovementsOUTSelectionAlias.PreviousRecordFK
                                                          }).ToList();

                //dbSapMovementsOUTSelection = (from draftDbSapMovementsOUTSelectionAlias in draftDbSapMovementsOUTSelection
                //                              join SM in _db.SapMaterial.AsNoTracking().ToListWithNoLock() on
                //                                  draftDbSapMovementsOUTSelectionAlias.SapMaterialCode equals SM.Code
                //                              join SESource in _db.SapEquipment.AsNoTracking().ToListWithNoLock() on new
                //                              {
                //                                  ErpPlantIdSource = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdSource.ToString(),
                //                                  ErpIdSource = draftDbSapMovementsOUTSelectionAlias.ErpIdSource.ToString()
                //                              }
                //                                  equals new
                //                                  {
                //                                      ErpPlantIdSource = SESource.ErpPlantId.ToString(),
                //                                      ErpIdSource = SESource.ErpId.ToString()
                //                                  }
                //                              join SEDest in _db.SapEquipment.AsNoTracking().ToListWithNoLock() on new
                //                              {
                //                                  ErpPlantIdSource = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdDest.ToString(),
                //                                  ErpIdSource = draftDbSapMovementsOUTSelectionAlias.ErpIdDest.ToString()
                //                              }
                //                              equals new
                //                              {
                //                                  ErpPlantIdSource = SEDest.ErpPlantId.ToString(),
                //                                  ErpIdSource = SEDest.ErpId.ToString()
                //                              }

                //                              select
                //                                          new SapMovementsOUT
                //                                          {
                //                                              Id = draftDbSapMovementsOUTSelectionAlias.Id,
                //                                              AddTime = draftDbSapMovementsOUTSelectionAlias.AddTime,
                //                                              BatchNo = draftDbSapMovementsOUTSelectionAlias.BatchNo,
                //                                              SapMaterialCode = draftDbSapMovementsOUTSelectionAlias.SapMaterialCode,
                //                                              SapMaterialFK = SM,
                //                                              ErpPlantIdSource = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdSource,
                //                                              ErpIdSource = draftDbSapMovementsOUTSelectionAlias.ErpIdSource,
                //                                              SapEquipmentSourceFK = SESource,
                //                                              IsWarehouseSource = draftDbSapMovementsOUTSelectionAlias.IsWarehouseSource,
                //                                              ErpPlantIdDest = draftDbSapMovementsOUTSelectionAlias.ErpPlantIdDest,
                //                                              ErpIdDest = draftDbSapMovementsOUTSelectionAlias.ErpIdDest,
                //                                              SapEquipmentDestFK = SEDest,
                //                                              IsWarehouseDest = draftDbSapMovementsOUTSelectionAlias.IsWarehouseDest,
                //                                              ValueTime = draftDbSapMovementsOUTSelectionAlias.ValueTime,
                //                                              Value = draftDbSapMovementsOUTSelectionAlias.Value,
                //                                              Correction2Previous = draftDbSapMovementsOUTSelectionAlias.Correction2Previous,
                //                                              IsReconciled = draftDbSapMovementsOUTSelectionAlias.IsReconciled,
                //                                              SapUnitOfMeasure = draftDbSapMovementsOUTSelectionAlias.SapUnitOfMeasure,
                //                                              SapGone = draftDbSapMovementsOUTSelectionAlias.SapGone,
                //                                              SapGoneTime = draftDbSapMovementsOUTSelectionAlias.SapGoneTime,
                //                                              SapError = draftDbSapMovementsOUTSelectionAlias.SapError,
                //                                              SapErrorMessage = draftDbSapMovementsOUTSelectionAlias.SapErrorMessage,
                //                                              MesMovementId = draftDbSapMovementsOUTSelectionAlias.MesMovementId,
                //                                              MesMovementsFK = draftDbSapMovementsOUTSelectionAlias.MesMovementsFK,
                //                                              MesParamId = draftDbSapMovementsOUTSelectionAlias.MesParamId,
                //                                              MesParamFK = draftDbSapMovementsOUTSelectionAlias.MesParamFK,
                //                                              PreviousRecordId = draftDbSapMovementsOUTSelectionAlias.PreviousRecordId,
                //                                              PreviousRecordFK = draftDbSapMovementsOUTSelectionAlias.PreviousRecordFK
                //                                          }).ToList();
            }
            else
            {
                dbSapMovementsOUTSelection = new List<SapMovementsOUT>();
            }
            return _mapper.Map<IEnumerable<SapMovementsOUT>, IEnumerable<SapMovementsOUTDTO>>(dbSapMovementsOUTSelection);
        }


        public async Task<SapMovementsOUTDTO> Update(SapMovementsOUTDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.SapMovementsOUT
                            .Include("MesMovementsFK")
                            .Include("PreviousRecordFK")
                            .Include("MesParamFK")
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

                if (objectToUpdateDTO.PreviousRecordId == null)
                {
                    objectToUpdate.PreviousRecordId = null;
                    objectToUpdate.PreviousRecordFK = null;
                }
                else
                {
                    if (objectToUpdate.PreviousRecordId != objectToUpdateDTO.PreviousRecordId)
                    {
                        objectToUpdate.PreviousRecordId = objectToUpdateDTO.PreviousRecordId;
                        var objectSapMovementsOUTToUpdate = _db.SapMovementsOUT
                            .Include("MesMovementsFK")
                            .Include("PreviousRecordFK")
                            .Include("MesParamFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.PreviousRecordId);
                        objectToUpdate.PreviousRecordFK = objectSapMovementsOUTToUpdate;
                    }
                }

                if (objectToUpdateDTO.MesMovementId == null)
                {
                    objectToUpdate.MesMovementId = null;
                    objectToUpdate.MesMovementsFK = null;
                }
                else
                {
                    if (objectToUpdate.MesMovementId != objectToUpdateDTO.MesMovementId)
                    {
                        objectToUpdate.MesMovementId = objectToUpdateDTO.MesMovementId;
                        var objectMesMovementsToUpdate = _db.MesMovements
                            .Include("MesMovementsFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesMovementId);
                        objectToUpdate.MesMovementsFK = objectMesMovementsToUpdate;
                    }
                }

                if (objectToUpdateDTO.AddTime != objectToUpdate.AddTime)
                    objectToUpdate.AddTime = objectToUpdateDTO.AddTime;

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

                if (objectToUpdateDTO.ValueTime != objectToUpdate.ValueTime)
                    objectToUpdate.ValueTime = objectToUpdateDTO.ValueTime;

                if (objectToUpdateDTO.Value != objectToUpdate.Value)
                    objectToUpdate.Value = objectToUpdateDTO.Value;

                if (objectToUpdateDTO.Correction2Previous != objectToUpdate.Correction2Previous)
                    objectToUpdate.Correction2Previous = objectToUpdateDTO.Correction2Previous;

                if (objectToUpdateDTO.IsReconciled != objectToUpdate.IsReconciled)
                    objectToUpdate.IsReconciled = objectToUpdateDTO.IsReconciled;

                if (objectToUpdateDTO.SapUnitOfMeasure != objectToUpdate.SapUnitOfMeasure)
                    objectToUpdate.SapUnitOfMeasure = objectToUpdateDTO.SapUnitOfMeasure;

                if (objectToUpdateDTO.SapGone != objectToUpdate.SapGone)
                    objectToUpdate.SapGone = objectToUpdateDTO.SapGone;

                if (objectToUpdateDTO.SapGoneTime != objectToUpdate.SapGoneTime)
                    objectToUpdate.SapGoneTime = objectToUpdateDTO.SapGoneTime;

                if (objectToUpdateDTO.SapError != objectToUpdate.SapError)
                    objectToUpdate.SapError = objectToUpdateDTO.SapError;

                if (objectToUpdateDTO.SapErrorMessage != objectToUpdate.SapErrorMessage)
                    objectToUpdate.SapErrorMessage = objectToUpdateDTO.SapErrorMessage;

                if (objectToUpdateDTO.MesMovementId != objectToUpdate.MesMovementId)
                    objectToUpdate.MesMovementId = objectToUpdateDTO.MesMovementId;

                if (objectToUpdateDTO.PreviousRecordId != objectToUpdate.PreviousRecordId)
                    objectToUpdate.PreviousRecordId = objectToUpdateDTO.PreviousRecordId;


                _db.SapMovementsOUT.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapMovementsOUT, SapMovementsOUTDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(Guid id)
        {
            if (id != Guid.Empty)
            {
                var objectToDelete = _db.SapMovementsOUT.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.SapMovementsOUT.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;
        }

        public async Task<IEnumerable<SapMovementsOUTDTO>?> GetListByMesMovementId(Guid? idPar)
        {
            if ((idPar != null) && (idPar != Guid.Empty))
            {
                var objectListToGet = _db.SapMovementsOUT.Where(u => u.MesMovementId == idPar).ToListWithNoLock();
                if (objectListToGet != null)
                {
                    return _mapper.Map<IEnumerable<SapMovementsOUT>, IEnumerable<SapMovementsOUTDTO>>(objectListToGet);
                }
            }
            return null;
        }

        public async Task<SapMovementsOUTDTO?> CleanMesMovementId(SapMovementsOUTDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.SapMovementsOUT
               .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                objectToUpdate.MesMovementId = null;
                objectToUpdate.MesMovementsFK = null;
                _db.SapMovementsOUT.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapMovementsOUT, SapMovementsOUTDTO>(objectToUpdate);
            }
            return null;
        }

        public async Task<IEnumerable<SapMovementsOUTDTO>?> GetListByPreviousRecordId(Guid? idPar)
        {
            if ((idPar != null) && (idPar != Guid.Empty))
            {
                var objectListToGet = _db.SapMovementsOUT.Where(u => u.PreviousRecordId == idPar).ToListWithNoLock();
                if (objectListToGet != null)
                {
                    return _mapper.Map<IEnumerable<SapMovementsOUT>, IEnumerable<SapMovementsOUTDTO>>(objectListToGet);
                }
            }
            return null;
        }

        public async Task<SapMovementsOUTDTO?> CleanPreviousRecordId(SapMovementsOUTDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.SapMovementsOUT
               .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                objectToUpdate.PreviousRecordId = null;
                _db.SapMovementsOUT.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapMovementsOUT, SapMovementsOUTDTO>(objectToUpdate);
            }
            return null;
        }
    }
}
