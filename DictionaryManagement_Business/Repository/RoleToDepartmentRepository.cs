using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class RoleToDepartmentRepository : IRoleToDepartmentRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public RoleToDepartmentRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<RoleToDepartmentDTO> Create(RoleToDepartmentDTO objectToAddDTO)
        {

            RoleToDepartment objectToAdd = new RoleToDepartment();

            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.RoleId = objectToAddDTO.RoleId;
            objectToAdd.DepartmentId = objectToAddDTO.DepartmentId;


            var addedRoleToDepartment = _db.RoleToDepartment.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<RoleToDepartment, RoleToDepartmentDTO>(addedRoleToDepartment.Entity);
        }

        public async Task<RoleToDepartmentDTO> Get(Guid roleId, int departmentId)
        {
            var objToGet = _db.RoleToDepartment.Include("RoleFK").Include("DepartmentFK").
                            FirstOrDefaultWithNoLock(u => u.RoleId == roleId && u.DepartmentId == departmentId);
            if (objToGet != null)
            {
                return _mapper.Map<RoleToDepartment, RoleToDepartmentDTO>(objToGet);
            }
            return null;
        }

        public async Task<RoleToDepartmentDTO> GetById(int id)
        {
            var objToGet = _db.RoleToDepartment.Include("RoleFK").Include("DepartmentFK").
                            FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<RoleToDepartment, RoleToDepartmentDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<RoleToDepartmentDTO>> GetAll()
        {
            var hhh = _db.RoleToDepartment.Include("RoleFK").Include("DepartmentFK").ToListWithNoLock();
            return _mapper.Map<IEnumerable<RoleToDepartment>, IEnumerable<RoleToDepartmentDTO>>(hhh);

        }

        public async Task<RoleToDepartmentDTO> Update(RoleToDepartmentDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.RoleToDepartment.Include("RoleFK").Include("DeaprtmentFK").
                    FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {

                if (objectToUpdate.RoleId != objectToUpdateDTO.RoleDTOFK.Id)
                {
                    objectToUpdate.RoleId = objectToUpdateDTO.RoleDTOFK.Id;
                    objectToUpdate.RoleFK = _mapper.Map<RoleDTO, Role>(objectToUpdateDTO.RoleDTOFK);
                }
                if (objectToUpdate.DepartmentId != objectToUpdateDTO.DepartmentDTOFK.Id)
                {
                    objectToUpdate.DepartmentId = objectToUpdateDTO.DepartmentDTOFK.Id;
                    objectToUpdate.DepartmentFK = _mapper.Map<MesDepartmentDTO, MesDepartment>(objectToUpdateDTO.DepartmentDTOFK);
                }
                _db.RoleToDepartment.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<RoleToDepartment, RoleToDepartmentDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id)
        {
            if (id > 0)
            {
                var objectToDelete = _db.RoleToDepartment.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.RoleToDepartment.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
