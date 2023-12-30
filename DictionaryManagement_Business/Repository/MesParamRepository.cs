using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;
//using Microsoft.EntityFrameworkCore.ChangeTracking;
using Microsoft.IdentityModel.Tokens;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository
{
    public class MesParamRepository : IMesParamRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public MesParamRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<MesParamDTO> Create(MesParamDTO objectToAddDTO)
        {
            //var objectToAdd = _mapper.Map<UnitOfMeasureSapToMesMappingDTO, UnitOfMeasureSapToMesMapping>(objectToAddDTO);

            MesParam objectToAdd = new MesParam();

            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.Code = objectToAddDTO.Code;
            objectToAdd.Name = objectToAddDTO.Name;
            objectToAdd.Description = objectToAddDTO.Description;
            objectToAdd.MesParamSourceType = objectToAddDTO.MesParamSourceType;
            objectToAdd.MesParamSourceLink = objectToAddDTO.MesParamSourceLink;
            objectToAdd.DepartmentId = objectToAddDTO.DepartmentId;
            objectToAdd.SapEquipmentIdSource = objectToAddDTO.SapEquipmentIdSource;
            objectToAdd.SapEquipmentIdDest = objectToAddDTO.SapEquipmentIdDest;
            objectToAdd.MesMaterialId = objectToAddDTO.MesMaterialId;
            objectToAdd.SapMaterialId = objectToAddDTO.SapMaterialId;
            objectToAdd.MesUnitOfMeasureId = objectToAddDTO.MesUnitOfMeasureId;
            objectToAdd.SapUnitOfMeasureId = objectToAddDTO.SapUnitOfMeasureId;
            objectToAdd.DaysRequestInPast = objectToAddDTO.DaysRequestInPast;

            objectToAdd.TI = objectToAddDTO.TI;
            objectToAdd.NameTI = objectToAddDTO.NameTI;
            objectToAdd.TM = objectToAddDTO.TM;
            objectToAdd.NameTM = objectToAddDTO.NameTM;
            objectToAdd.MesToSirUnitOfMeasureKoef = objectToAddDTO.MesToSirUnitOfMeasureKoef;

            objectToAdd.NeedWriteToSap = objectToAddDTO.NeedWriteToSap;
            objectToAdd.NeedReadFromSap = objectToAddDTO.NeedReadFromSap;
            objectToAdd.NeedReadFromMes = objectToAddDTO.NeedReadFromMes;
            objectToAdd.NeedWriteToMes = objectToAddDTO.NeedWriteToMes;
            objectToAdd.IsNdo = objectToAddDTO.IsNdo;
            objectToAdd.IsArchive = objectToAddDTO.IsArchive;

            var addedMesParam = _db.MesParam.Add(objectToAdd);
            _db.SaveChanges();

            return (await GetById(addedMesParam.Entity.Id));
        }


        public async Task<MesParamDTO> GetById(int id)
        {
            var objToGet = _db.MesParam
                            .Include("MesParamSourceTypeFK")
                            .Include("SapEquipmentSourceFK")
                            .Include("SapEquipmentDestFK")
                            .Include("MesMaterialFK")
                            .Include("SapMaterialFK")
                            .Include("MesUnitOfMeasureFK")
                            .Include("SapUnitOfMeasureFK")
                            .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<MesParam, MesParamDTO>(objToGet);
            }
            return null;
        }


        public async Task<MesParamDTO> GetByCode(string code = "")
        {
            var objToGet = _db.MesParam
                            .Include("MesParamSourceTypeFK")
                            .Include("SapEquipmentSourceFK")
                            .Include("SapEquipmentDestFK")
                            .Include("MesMaterialFK")
                            .Include("SapMaterialFK")
                            .Include("MesUnitOfMeasureFK")
                            .Include("SapUnitOfMeasureFK")
                            .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .FirstOrDefaultWithNoLock(u => u.Code.Trim().ToUpper() == code.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesParam, MesParamDTO>(objToGet);
            }
            return null;
        }

        public async Task<MesParamDTO> GetByName(string name = "")
        {
            var objToGet = _db.MesParam
                            .Include("MesParamSourceTypeFK")
                            .Include("SapEquipmentSourceFK")
                            .Include("SapEquipmentDestFK")
                            .Include("MesMaterialFK")
                            .Include("SapMaterialFK")
                            .Include("MesUnitOfMeasureFK")
                            .Include("SapUnitOfMeasureFK")
                            .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .FirstOrDefaultWithNoLock(u => u.Name.Trim().ToUpper() == name.Trim().ToUpper());
            if (objToGet != null)
            {
                return _mapper.Map<MesParam, MesParamDTO>(objToGet);
            }
            return null;
        }
        public async Task<MesParamDTO> GetByMesParamSourceLink(string mesParamSourceLink = "")
        {
            if (!mesParamSourceLink.IsNullOrEmpty())
            {
                var objToGet = _db.MesParam
                            .Include("MesParamSourceTypeFK")
                            .Include("SapEquipmentSourceFK")
                            .Include("SapEquipmentDestFK")
                            .Include("MesMaterialFK")
                            .Include("SapMaterialFK")
                            .Include("MesUnitOfMeasureFK")
                            .Include("SapUnitOfMeasureFK")
                            .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .FirstOrDefaultWithNoLock(u => u.MesParamSourceLink.Trim().ToUpper() == mesParamSourceLink.Trim().ToUpper());
                if (objToGet != null)
                {
                    return _mapper.Map<MesParam, MesParamDTO>(objToGet);
                }

            }
            return null;
        }


        public async Task<IEnumerable<MesParamDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
            {
                var hhh2 = _db.MesParam
                            .Include("MesParamSourceTypeFK").AsNoTracking()
                            .Include("SapEquipmentSourceFK").AsNoTracking()
                            .Include("SapEquipmentDestFK").AsNoTracking()
                            .Include("MesMaterialFK").AsNoTracking()
                            .Include("SapMaterialFK").AsNoTracking()
                            .Include("MesUnitOfMeasureFK").AsNoTracking()
                            .Include("SapUnitOfMeasureFK").AsNoTracking()
                            .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Where(u => u.IsArchive == true).AsNoTracking().ToListWithNoLock();
                var retVar2 = _mapper.Map<IEnumerable<MesParam>, IEnumerable<MesParamDTO>>(hhh2);
                //GC.Collect(2, GCCollectionMode.Forced);
                //_db.Dispose();
                //GC.Collect();
                return retVar2;
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
            {
                var hhh3 = _db.MesParam
                            .Include("MesParamSourceTypeFK").AsNoTracking()
                            .Include("SapEquipmentSourceFK").AsNoTracking()
                            .Include("SapEquipmentDestFK").AsNoTracking()
                            .Include("MesMaterialFK").AsNoTracking()
                            .Include("SapMaterialFK").AsNoTracking()
                            .Include("MesUnitOfMeasureFK").AsNoTracking()
                            .Include("SapUnitOfMeasureFK").AsNoTracking()
                            .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Where(u => u.IsArchive != true).AsNoTracking().ToListWithNoLock();
                var retVar3 = _mapper.Map<IEnumerable<MesParam>, IEnumerable<MesParamDTO>>(hhh3);
                // _db.Dispose();
                //GC.Collect(2, GCCollectionMode.Forced);
                //GC.Collect();
                return retVar3;

            }
            var hhh1 = _db.MesParam
                        .Include("MesParamSourceTypeFK")
                        .Include("SapEquipmentSourceFK")
                        .Include("SapEquipmentDestFK")
                        .Include("MesMaterialFK")
                        .Include("SapMaterialFK")
                        .Include("MesUnitOfMeasureFK")
                        .Include("SapUnitOfMeasureFK")
                        .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                        .ToListWithNoLock();
            var retVar1 = _mapper.Map<IEnumerable<MesParam>, IEnumerable<MesParamDTO>>(hhh1);
            //GC.Collect(2, GCCollectionMode.Forced);
            //_db.Dispose();
            //GC.Collect();            
            return retVar1;
        }


        public async Task<MesParamDTO> Update(MesParamDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.MesParam
                        .Include("MesParamSourceTypeFK")
                        .Include("SapEquipmentSourceFK")
                        .Include("SapEquipmentDestFK")
                        .Include("MesMaterialFK")
                        .Include("SapMaterialFK")
                        .Include("MesUnitOfMeasureFK")
                        .Include("SapUnitOfMeasureFK")
                        .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                        .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                        .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.MesParamSourceType == null)
                {
                    objectToUpdate.MesParamSourceType = null;
                    objectToUpdate.MesParamSourceTypeFK = null;
                }
                else
                {
                    if (objectToUpdate.MesParamSourceType != objectToUpdateDTO.MesParamSourceType)
                    {
                        objectToUpdate.MesParamSourceType = objectToUpdateDTO.MesParamSourceType;
                        var objectMesParamSourceTypeToUpdate = _db.MesParamSourceType.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesParamSourceType);
                        objectToUpdate.MesParamSourceTypeFK = objectMesParamSourceTypeToUpdate;
                    }
                }

                if (objectToUpdateDTO.DepartmentId == null)
                {
                    objectToUpdate.DepartmentId = null;
                    objectToUpdate.MesDepartmentFK = null;
                }
                else
                {
                    if (objectToUpdate.DepartmentId != objectToUpdateDTO.DepartmentId)
                    {
                        objectToUpdate.DepartmentId = objectToUpdateDTO.DepartmentId;
                        var objectMesDepartmentToUpdate = _db.MesDepartment.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.DepartmentId);
                        objectToUpdate.MesDepartmentFK = objectMesDepartmentToUpdate;
                    }
                }

                if (objectToUpdateDTO.SapEquipmentIdSource == null)
                {
                    objectToUpdate.SapEquipmentIdSource = null;
                    objectToUpdate.SapEquipmentSourceFK = null;
                }
                else
                {
                    if (objectToUpdate.SapEquipmentIdSource != objectToUpdateDTO.SapEquipmentIdSource)
                    {
                        objectToUpdate.SapEquipmentIdSource = objectToUpdateDTO.SapEquipmentIdSource;
                        var objectSapEquipmentToUpdate = _db.SapEquipment.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.SapEquipmentIdSource);
                        objectToUpdate.SapEquipmentSourceFK = objectSapEquipmentToUpdate;
                    }
                }

                if (objectToUpdateDTO.SapEquipmentIdDest == null)
                {
                    objectToUpdate.SapEquipmentIdDest = null;
                    objectToUpdate.SapEquipmentDestFK = null;
                }
                else
                {
                    if (objectToUpdate.SapEquipmentIdDest != objectToUpdateDTO.SapEquipmentIdDest)
                    {
                        objectToUpdate.SapEquipmentIdDest = objectToUpdateDTO.SapEquipmentIdDest;
                        var objectSapEquipmentToUpdate = _db.SapEquipment.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.SapEquipmentIdDest);
                        objectToUpdate.SapEquipmentDestFK = objectSapEquipmentToUpdate;
                    }
                }


                if (objectToUpdateDTO.MesMaterialId == null)
                {
                    objectToUpdate.MesMaterialId = null;
                    objectToUpdate.MesMaterialFK = null;
                }
                else
                {
                    if (objectToUpdate.MesMaterialId != objectToUpdateDTO.MesMaterialId)
                    {
                        objectToUpdate.MesMaterialId = objectToUpdateDTO.MesMaterialId;
                        var objectMesMaterialToUpdate = _db.MesMaterial.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesMaterialId);
                        objectToUpdate.MesMaterialFK = objectMesMaterialToUpdate;
                    }
                }


                if (objectToUpdateDTO.SapMaterialId == null)
                {
                    objectToUpdate.SapMaterialId = null;
                    objectToUpdate.SapMaterialFK = null;
                }
                else
                {
                    if (objectToUpdate.SapMaterialId != objectToUpdateDTO.SapMaterialId)
                    {
                        objectToUpdate.SapMaterialId = objectToUpdateDTO.SapMaterialId;
                        var objectSapMaterialToUpdate = _db.SapMaterial.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.SapMaterialId);
                        objectToUpdate.SapMaterialFK = objectSapMaterialToUpdate;
                    }
                }

                if (objectToUpdateDTO.MesUnitOfMeasureId == null)
                {
                    objectToUpdate.MesUnitOfMeasureId = null;
                    objectToUpdate.MesUnitOfMeasureFK = null;
                }
                else
                {
                    if (objectToUpdate.MesUnitOfMeasureId != objectToUpdateDTO.MesUnitOfMeasureId)
                    {
                        objectToUpdate.MesUnitOfMeasureId = objectToUpdateDTO.MesUnitOfMeasureId;
                        var objectMesUnitOfMeasureToUpdate = _db.MesUnitOfMeasure.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.MesUnitOfMeasureId);
                        objectToUpdate.MesUnitOfMeasureFK = objectMesUnitOfMeasureToUpdate;
                    }
                }

                if (objectToUpdateDTO.SapUnitOfMeasureId == null)
                {
                    objectToUpdate.SapUnitOfMeasureId = null;
                    objectToUpdate.SapUnitOfMeasureFK = null;
                }
                else
                {
                    if (objectToUpdate.SapUnitOfMeasureId != objectToUpdateDTO.SapUnitOfMeasureId)
                    {
                        objectToUpdate.SapUnitOfMeasureId = objectToUpdateDTO.SapUnitOfMeasureId;
                        var objectSapUnitOfMeasureToUpdate = _db.SapUnitOfMeasure.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.SapUnitOfMeasureId);
                        objectToUpdate.SapUnitOfMeasureFK = objectSapUnitOfMeasureToUpdate;
                    }
                }

                if (objectToUpdateDTO.Code != objectToUpdate.Code)
                    objectToUpdate.Code = objectToUpdateDTO.Code;
                if (objectToUpdateDTO.Name != objectToUpdate.Name)
                    objectToUpdate.Name = objectToUpdateDTO.Name;
                if (objectToUpdateDTO.Description != objectToUpdate.Description)
                    objectToUpdate.Description = objectToUpdateDTO.Description;
                if (objectToUpdateDTO.MesParamSourceLink != objectToUpdate.MesParamSourceLink)
                    objectToUpdate.MesParamSourceLink = objectToUpdateDTO.MesParamSourceLink;
                if (objectToUpdateDTO.DaysRequestInPast != objectToUpdate.DaysRequestInPast)
                    objectToUpdate.DaysRequestInPast = objectToUpdateDTO.DaysRequestInPast;

                if (objectToUpdateDTO.TI != objectToUpdate.TI)
                    objectToUpdate.TI = objectToUpdateDTO.TI;
                if (objectToUpdateDTO.NameTI != objectToUpdate.NameTI)
                    objectToUpdate.NameTI = objectToUpdateDTO.NameTI;
                if (objectToUpdateDTO.TM != objectToUpdate.TM)
                    objectToUpdate.TM = objectToUpdateDTO.TM;
                if (objectToUpdateDTO.NameTM != objectToUpdate.NameTM)
                    objectToUpdate.NameTM = objectToUpdateDTO.NameTM;
                if (objectToUpdateDTO.MesToSirUnitOfMeasureKoef != objectToUpdate.MesToSirUnitOfMeasureKoef)
                    objectToUpdate.MesToSirUnitOfMeasureKoef = objectToUpdateDTO.MesToSirUnitOfMeasureKoef;

                if (objectToUpdateDTO.NeedWriteToSap != objectToUpdate.NeedWriteToSap)
                    objectToUpdate.NeedWriteToSap = objectToUpdateDTO.NeedWriteToSap;
                if (objectToUpdateDTO.NeedReadFromSap != objectToUpdate.NeedReadFromSap)
                    objectToUpdate.NeedReadFromSap = objectToUpdateDTO.NeedReadFromSap;
                if (objectToUpdateDTO.NeedReadFromMes != objectToUpdate.NeedReadFromMes)
                    objectToUpdate.NeedReadFromMes = objectToUpdateDTO.NeedReadFromMes;
                if (objectToUpdateDTO.NeedWriteToMes != objectToUpdate.NeedWriteToMes)
                    objectToUpdate.NeedWriteToMes = objectToUpdateDTO.NeedWriteToMes;
                if (objectToUpdateDTO.IsNdo != objectToUpdate.IsNdo)
                    objectToUpdate.IsNdo = objectToUpdateDTO.IsNdo;
                if (objectToUpdateDTO.IsArchive != objectToUpdate.IsArchive)
                    objectToUpdate.IsArchive = objectToUpdateDTO.IsArchive;

                _db.MesParam.Update(objectToUpdate);

                //var modifiedEntries = _db.ChangeTracker.Entries().Where(e => e.State == EntityState.Modified);
                //foreach (EntityEntry entity in modifiedEntries)
                //{
                //    foreach (var propName in entity.CurrentValues.Properties)
                //    {
                //        var current = entity.CurrentValues[propName];
                //        var original = entity.OriginalValues[propName];
                //    }
                //}

                _db.SaveChanges();
                return _mapper.Map<MesParam, MesParamDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id, UpdateMode updateMode = UpdateMode.Update)
        {
            if (id > 0)
            {
                var objectToDelete = _db.MesParam.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    if (updateMode == SD.UpdateMode.MoveToArchive)
                        objectToDelete.IsArchive = true;
                    if (updateMode == SD.UpdateMode.RestoreFromArchive)
                        objectToDelete.IsArchive = false;

                    ////typeof(T).GetProperties().Select(property => ((DisplayAttribute)property.GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault())?.Name).ToArray();

                    ////foreach (var propertyEntry in _db.Entry(objectToDelete).Properties.Select(property => ((DisplayAttribute)property..GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault())?.Name))

                    //foreach (var propertyEntry in _db.Entry(objectToDelete).Properties)
                    //{

                    //    //if (propertyEntry.Metadata.ClrType == typeof(bool))
                    //    if (propertyEntry.Metadata.Name == "IsArchive")
                    //    {
                    //        var sss = propertyEntry.Metadata.PropertyInfo.GetCustomAttributesData().Select(t => t.AttributeType == typeof(DisplayAttribute));

                    //        string? annotationName = null;

                    //        foreach (var attribute in propertyEntry.Metadata.PropertyInfo.GetCustomAttributes(false))
                    //        {
                    //            var test = attribute as DisplayAttribute;
                    //            if (test != null)
                    //            {
                    //                annotationName = test.Name;
                    //                break;
                    //            }
                    //        }



                    //        var ddd = 1;
                    //        //object[] attributes = propertyEntry.Metadata.PropertyInfo.GetCustomAttributes(true).Select(t => t.GetType().Name == "DisplayAttribute");
                    //        //object[] attributes = propertyEntry.Metadata.PropertyInfo.GetCustomAttributes(true).Select(t => t.GetType().Name == "DisplayAttribute");
                    //        //foreach (object item in attributes)
                    //        //{
                    //        //    if (item.GetType().FullName == fullName)
                    //        //        return true;
                    //        //}

                    //        //.GetCustomAttributesData().Select(t => t.AttributeType == typeof(DisplayAttribute));


                    //        //var yyy = typeof(MesParam).GetProperties().Select(property => ((DisplayAttribute)property.GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault())?.Name).ToList();
                    //        //var jjj = propertyEntry.Metadata.GetAnnotations();
                    //        //string fff = propertyEntry.Metadata.FindAnnotation("Display").Name;
                    //        //propertyEntry.CurrentValue = DateTime.Now;

                    //    }
                    //}

                    _db.MesParam.Update(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }

        public async Task<MesParamDTO> GetBySapMappingNotInArchive(int? sapEquipmentIdSource, int? sapEquipmentIdDest, int? sapMaterialId, int idForExclude)
        {
            if (sapEquipmentIdSource != null && sapEquipmentIdDest != null && sapMaterialId != null)
                if (sapEquipmentIdSource > 0 && sapEquipmentIdDest > 0 && sapMaterialId > 0)
                {
                    var objToGet = _db.MesParam
                                    .Include("MesParamSourceTypeFK")
                                    .Include("SapEquipmentSourceFK")
                                    .Include("SapEquipmentDestFK")
                                    .Include("MesMaterialFK")
                                    .Include("SapMaterialFK")
                                    .Include("MesUnitOfMeasureFK")
                                    .Include("SapUnitOfMeasureFK")
                                    .Include("MesDepartmentFK").Include("MesDepartmentFK.DepartmentParent")
                                    .Include("MesDepartmentFK.DepartmentParent.DepartmentParent")
                                    .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                                    .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                                    .Include("MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                                    .FirstOrDefaultWithNoLock(u => u.SapEquipmentIdSource == sapEquipmentIdSource
                                        && u.SapEquipmentIdDest == sapEquipmentIdDest && u.SapMaterialId == sapMaterialId
                                        && u.Id != idForExclude && u.IsArchive != true);
                    if (objToGet != null)
                    {
                        return _mapper.Map<MesParam, MesParamDTO>(objToGet);
                    }

                }
            return null;
        }

    }
}
