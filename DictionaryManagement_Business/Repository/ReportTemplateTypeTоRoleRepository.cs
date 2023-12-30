using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class ReportTemplateTypeToRoleRepository : IReportTemplateTypeToRoleRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ReportTemplateTypeToRoleRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<ReportTemplateTypeTоRoleDTO> Create(ReportTemplateTypeTоRoleDTO objectToAddDTO)
        {

            ReportTemplateTypeTоRole objectToAdd = new ReportTemplateTypeTоRole();

            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.ReportTemplateTypeId = objectToAddDTO.ReportTemplateTypeId;
            objectToAdd.RoleId = objectToAddDTO.RoleId;
            objectToAdd.CanDownload = objectToAddDTO.CanDownload;
            objectToAdd.CanUpload = objectToAddDTO.CanUpload;

            var addedReportTemplateTypeTоRole = _db.ReportTemplateTypeTоRole.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<ReportTemplateTypeTоRole, ReportTemplateTypeTоRoleDTO>(addedReportTemplateTypeTоRole.Entity);
        }

        public async Task<ReportTemplateTypeTоRoleDTO> Get(int reportTemplateTypeId, Guid roleId)
        {
            var objToGet = _db.ReportTemplateTypeTоRole.Include("ReportTemplateTypeFK").Include("RoleFK").
                            FirstOrDefaultWithNoLock(u => u.ReportTemplateTypeId == reportTemplateTypeId && u.RoleId == roleId);
            if (objToGet != null)
            {
                return _mapper.Map<ReportTemplateTypeTоRole, ReportTemplateTypeTоRoleDTO>(objToGet);
            }
            return null;
        }

        public async Task<ReportTemplateTypeTоRoleDTO> GetById(int id)
        {
            var objToGet = _db.ReportTemplateTypeTоRole.Include("ReportTemplateTypeFK").Include("RoleFK").
                            FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<ReportTemplateTypeTоRole, ReportTemplateTypeTоRoleDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<ReportTemplateTypeTоRoleDTO>> GetAll()
        {
            var hhh = _db.ReportTemplateTypeTоRole.Include("ReportTemplateTypeFK").Include("RoleFK").ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportTemplateTypeTоRole>, IEnumerable<ReportTemplateTypeTоRoleDTO>>(hhh);

        }

        public async Task<ReportTemplateTypeTоRoleDTO> Update(ReportTemplateTypeTоRoleDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.ReportTemplateTypeTоRole.Include("ReportTemplateTypeFK").Include("RoleFK").
                    FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (objectToUpdate.ReportTemplateTypeId != objectToUpdateDTO.ReportTemplateTypeDTOFK.Id)
                {
                    objectToUpdate.ReportTemplateTypeId = objectToUpdateDTO.ReportTemplateTypeDTOFK.Id;
                    objectToUpdate.ReportTemplateTypeFK = _mapper.Map<ReportTemplateTypeDTO, ReportTemplateType>(objectToUpdateDTO.ReportTemplateTypeDTOFK);
                }
                if (objectToUpdate.RoleId != objectToUpdateDTO.RoleDTOFK.Id)
                {
                    objectToUpdate.RoleId = objectToUpdateDTO.RoleDTOFK.Id;
                    objectToUpdate.RoleFK = _mapper.Map<RoleDTO, Role>(objectToUpdateDTO.RoleDTOFK);
                }

                if (objectToUpdate.CanDownload != objectToUpdateDTO.CanDownload)
                    objectToUpdate.CanDownload = objectToUpdateDTO.CanDownload;

                if (objectToUpdate.CanUpload != objectToUpdateDTO.CanUpload)
                    objectToUpdate.CanUpload = objectToUpdateDTO.CanUpload;

                _db.ReportTemplateTypeTоRole.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<ReportTemplateTypeTоRole, ReportTemplateTypeTоRoleDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id)
        {
            if (id > 0)
            {
                var objectToDelete = _db.ReportTemplateTypeTоRole.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.ReportTemplateTypeTоRole.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
