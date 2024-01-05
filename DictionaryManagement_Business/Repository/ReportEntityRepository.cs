using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class ReportEntityRepository : IReportEntityRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ReportEntityRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<ReportEntityDTO> Create(ReportEntityDTO objectToAddDTO)
        {

            ReportEntity objectToAdd = new ReportEntity();

            if (objectToAddDTO.Id == null || objectToAddDTO.Id == Guid.Empty)
                objectToAdd.Id = Guid.NewGuid();
            else
                objectToAdd.Id = objectToAddDTO.Id;

            if (objectToAddDTO.DownloadTime == null)
                objectToAdd.DownloadTime = DateTime.Now;
            else
                objectToAdd.DownloadTime = objectToAddDTO.DownloadTime;

            objectToAdd.ReportTimeStart = objectToAddDTO.ReportTimeStart;
            objectToAdd.ReportTimeEnd = objectToAddDTO.ReportTimeEnd;
            objectToAdd.ReportDepartmentId = objectToAddDTO.ReportDepartmentId;
            objectToAdd.DownloadUserId = objectToAddDTO.DownloadUserId;
            objectToAdd.DownloadReportFileName = objectToAddDTO.DownloadReportFileName;
            objectToAdd.DownloadSuccessFlag = objectToAddDTO.DownloadSuccessFlag;

            objectToAdd.ReportTemplateId = objectToAddDTO.ReportTemplateId;

            objectToAdd.UploadTime = objectToAddDTO.UploadTime;
            objectToAdd.UploadUserId = objectToAddDTO.UploadUserId;
            objectToAdd.UploadReportFileName = objectToAddDTO.UploadReportFileName;
            objectToAdd.UploadSuccessFlag = objectToAddDTO.UploadSuccessFlag;

            var addedReportEntity = _db.ReportEntity.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<ReportEntity, ReportEntityDTO>(addedReportEntity.Entity);
        }


        public async Task<ReportEntityDTO> GetById(Guid id)
        {
            var objToGet = _db.ReportEntity
                            .Include("ReportTemplateFK").Include("ReportTemplateFK.ReportTemplateTypeFK").Include("ReportTemplateFK.MesDepartmentFK")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK")
                            .Include("ReportDepartmentFK.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DownloadUserFK")
                            .Include("UploadUserFK")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<ReportEntity, ReportEntityDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<ReportEntityDTO>> GetAll()
        {
            var hhh1 = _db.ReportEntity
                            .Include("ReportTemplateFK").Include("ReportTemplateFK.ReportTemplateTypeFK").Include("ReportTemplateFK.MesDepartmentFK")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK")
                            .Include("ReportDepartmentFK.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DownloadUserFK")
                            .Include("UploadUserFK").ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntity>, IEnumerable<ReportEntityDTO>>(hhh1);
        }

        public async Task<IEnumerable<ReportEntityDTO>> GetAllByDownloadTimeInterval(DateTime? startDownloadTime, DateTime? endDownloadTime)
        {

            if (startDownloadTime == null)
                startDownloadTime = DateTime.MinValue;
            if (startDownloadTime == null)
                startDownloadTime = DateTime.MaxValue;

            var hhh1 = _db.ReportEntity
                            .Include("ReportTemplateFK").Include("ReportTemplateFK.ReportTemplateTypeFK").Include("ReportTemplateFK.MesDepartmentFK")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK")
                            .Include("ReportDepartmentFK.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DownloadUserFK")
                            .Include("UploadUserFK")
                            .Where(u => u.DownloadTime >= startDownloadTime && u.DownloadTime <= endDownloadTime).ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntity>, IEnumerable<ReportEntityDTO>>(hhh1);

        }
        public async Task<IEnumerable<ReportEntityDTO>> GetAllByUploadTimeInterval(DateTime? startUploadTime, DateTime? endUploadTime)
        {

            if (startUploadTime == null)
                startUploadTime = DateTime.MinValue;
            if (startUploadTime == null)
                startUploadTime = DateTime.MaxValue;


            var hhh1 = _db.ReportEntity
                .Include("ReportTemplateFK").Include("ReportTemplateFK.ReportTemplateTypeFK").Include("ReportTemplateFK.MesDepartmentFK")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK")
                            .Include("ReportDepartmentFK.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DownloadUserFK")
                            .Include("UploadUserFK")
                .Where(u => u.UploadTime >= startUploadTime && u.UploadTime <= endUploadTime).ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntity>, IEnumerable<ReportEntityDTO>>(hhh1);
        }

        public async Task<ReportEntityDTO> Update(ReportEntityDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.ReportEntity
                .Include("ReportTemplateFK").Include("ReportTemplateFK.ReportTemplateTypeFK").Include("ReportTemplateFK.MesDepartmentFK")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportTemplateFK.MesDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK")
                            .Include("ReportDepartmentFK.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("ReportDepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DownloadUserFK")
                            .Include("UploadUserFK")
               .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.ReportTemplateId == null || objectToUpdateDTO.ReportTemplateId == Guid.Empty)
                {
                    objectToUpdate.ReportTemplateId = Guid.Empty;
                    objectToUpdate.ReportTemplateFK = null;
                }
                else
                {
                    if (objectToUpdate.ReportTemplateId != objectToUpdateDTO.ReportTemplateId)
                    {
                        objectToUpdate.ReportTemplateId = objectToUpdateDTO.ReportTemplateId;
                        var objectReportTemplateToUpdate = _db.ReportTemplate.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.ReportTemplateId);
                        objectToUpdate.ReportTemplateFK = objectReportTemplateToUpdate;
                    }
                }

                if (objectToUpdateDTO.ReportDepartmentId == null)
                {
                    objectToUpdate.ReportDepartmentId = null;
                    objectToUpdate.ReportDepartmentFK = null;
                }
                else
                {
                    if (objectToUpdate.ReportDepartmentId != objectToUpdateDTO.ReportDepartmentId)
                    {
                        objectToUpdate.ReportDepartmentId = objectToUpdateDTO.ReportDepartmentId;
                        var objectDataTypeToUpdate = _db.MesDepartment.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.ReportDepartmentId);
                        objectToUpdate.ReportDepartmentFK = objectDataTypeToUpdate;
                    }
                }

                if (objectToUpdateDTO.DownloadUserId == null || objectToUpdateDTO.DownloadUserId == Guid.Empty)
                {
                    objectToUpdate.DownloadUserId = Guid.Empty;
                    objectToUpdate.DownloadUserFK = null;
                }
                else
                {
                    if (objectToUpdate.DownloadUserId != objectToUpdateDTO.DownloadUserId)
                    {
                        objectToUpdate.DownloadUserId = objectToUpdateDTO.DownloadUserId;
                        var objectDataTypeToUpdate = _db.User.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.DownloadUserId);
                        objectToUpdate.DownloadUserFK = objectDataTypeToUpdate;
                    }
                }

                if (objectToUpdateDTO.UploadUserId == null)
                {
                    objectToUpdate.UploadUserId = null;
                    objectToUpdate.UploadUserFK = null;
                }
                else
                {
                    if (objectToUpdate.UploadUserId != objectToUpdateDTO.UploadUserId)
                    {
                        objectToUpdate.UploadUserId = objectToUpdateDTO.UploadUserId;
                        var objectDataTypeToUpdate = _db.User.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.UploadUserId);
                        objectToUpdate.UploadUserFK = objectDataTypeToUpdate;
                    }
                }

                if (objectToUpdateDTO.ReportTimeStart != objectToUpdate.ReportTimeStart)
                    objectToUpdate.ReportTimeStart = objectToUpdateDTO.ReportTimeStart;

                if (objectToUpdateDTO.ReportTimeEnd != objectToUpdate.ReportTimeEnd)
                    objectToUpdate.ReportTimeEnd = objectToUpdateDTO.ReportTimeEnd;

                if (objectToUpdateDTO.DownloadTime != objectToUpdate.DownloadTime)
                    objectToUpdate.DownloadTime = objectToUpdateDTO.DownloadTime;

                if (objectToUpdateDTO.DownloadReportFileName != objectToUpdate.DownloadReportFileName)
                    objectToUpdate.DownloadReportFileName = objectToUpdateDTO.DownloadReportFileName;

                if (objectToUpdateDTO.DownloadSuccessFlag != objectToUpdate.DownloadSuccessFlag)
                    objectToUpdate.DownloadSuccessFlag = objectToUpdateDTO.DownloadSuccessFlag;

                if (objectToUpdateDTO.UploadTime != objectToUpdate.UploadTime)
                    objectToUpdate.UploadTime = objectToUpdateDTO.UploadTime;

                if (objectToUpdateDTO.UploadReportFileName != objectToUpdate.UploadReportFileName)
                    objectToUpdate.UploadReportFileName = objectToUpdateDTO.UploadReportFileName;

                if (objectToUpdateDTO.UploadSuccessFlag != objectToUpdate.UploadSuccessFlag)
                    objectToUpdate.UploadSuccessFlag = objectToUpdateDTO.UploadSuccessFlag;

                _db.ReportEntity.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<ReportEntity, ReportEntityDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(Guid id)
        {
            if (id != null && id != Guid.Empty)
            {
                var objectToDelete = _db.ReportEntity.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.ReportEntity.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
