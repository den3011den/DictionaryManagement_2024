using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class RoleToADGroupRepository : IRoleToADGroupRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public RoleToADGroupRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<RoleToADGroupDTO> Create(RoleToADGroupDTO objectToAddDTO)
        {

            RoleToADGroup objectToAdd = new RoleToADGroup();

            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.RoleId = objectToAddDTO.RoleId;
            objectToAdd.ADGroupId = objectToAddDTO.ADGroupId;


            var addedRoleToADGroup = _db.RoleToADGroup.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<RoleToADGroup, RoleToADGroupDTO>(addedRoleToADGroup.Entity);
        }

        public async Task<RoleToADGroupDTO> Get(Guid adGroupId, Guid roleId)
        {
            var objToGet = _db.RoleToADGroup.Include("RoleFK").Include("ADGroupFK").
                            FirstOrDefault(u => u.RoleId == roleId && u.ADGroupId == adGroupId);
            if (objToGet != null)
            {
                return _mapper.Map<RoleToADGroup, RoleToADGroupDTO>(objToGet);
            }
            return null;
        }

        public async Task<RoleToADGroupDTO> GetById(int id)
        {
            var objToGet = _db.RoleToADGroup.Include("RoleFK").Include("ADGroupFK").
                            FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<RoleToADGroup, RoleToADGroupDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<RoleToADGroupDTO>> GetByRoleId(Guid roleId)
        {
            var objToGet = _db.RoleToADGroup.Include("RoleFK").Include("ADGroupFK").Where(u => u.RoleId == roleId).AsNoTracking().ToListWithNoLock();
            if (objToGet != null)
            {
                return _mapper.Map<IEnumerable<RoleToADGroup>, IEnumerable<RoleToADGroupDTO>>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<RoleToADGroupDTO>> GetAll()
        {
            var hhh = _db.RoleToADGroup.Include("RoleFK").Include("ADGroupFK").AsNoTracking().ToListWithNoLock();
            return _mapper.Map<IEnumerable<RoleToADGroup>, IEnumerable<RoleToADGroupDTO>>(hhh);

        }

        public async Task<RoleToADGroupDTO> Update(RoleToADGroupDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.RoleToADGroup.Include("RoleFK").Include("ADGroupFK").
                    FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {

                if (objectToUpdate.RoleId != objectToUpdateDTO.RoleDTOFK.Id)
                {
                    objectToUpdate.RoleId = objectToUpdateDTO.RoleDTOFK.Id;
                    objectToUpdate.RoleFK = _mapper.Map<RoleDTO, Role>(objectToUpdateDTO.RoleDTOFK);
                }
                if (objectToUpdate.ADGroupId != objectToUpdateDTO.ADGroupDTOFK.Id)
                {
                    objectToUpdate.ADGroupId = objectToUpdateDTO.ADGroupDTOFK.Id;
                    objectToUpdate.ADGroupFK = _mapper.Map<ADGroupDTO, ADGroup>(objectToUpdateDTO.ADGroupDTOFK);
                }
                _db.RoleToADGroup.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<RoleToADGroup, RoleToADGroupDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id)
        {
            if (id > 0)
            {
                var objectToDelete = _db.RoleToADGroup.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.RoleToADGroup.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }

    }
}
