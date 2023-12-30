using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class SmenaRepository : ISmenaRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SmenaRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SmenaDTO> Create(SmenaDTO objectToAddDTO)
        {

            Smena objectToAdd = new Smena();

            objectToAdd.Name = objectToAddDTO.Name;
            objectToAdd.StartTime = objectToAddDTO.StartTime;
            objectToAdd.HoursDuration = objectToAddDTO.HoursDuration;
            objectToAdd.DepartmentId = objectToAddDTO.DepartmentId;
            objectToAdd.IsArchive = objectToAddDTO.IsArchive;

            var addedSmena = _db.Smena.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<Smena, SmenaDTO>(addedSmena.Entity);
        }


        public async Task<SmenaDTO> GetById(int id)
        {
            var objToGet = _db.Smena
                            .Include("DepartmentFK")
                            .Include("DepartmentFK.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<Smena, SmenaDTO>(objToGet);
            }
            return null;
        }

        public async Task<IEnumerable<SmenaDTO>> GetAllByDepartmentId(int departmentId)
        {
            var hhh1 = _db.Smena
                            .Include("DepartmentFK")
                            .Include("DepartmentFK.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Where(u => u.DepartmentId == departmentId).ToListWithNoLock();
            return _mapper.Map<IEnumerable<Smena>, IEnumerable<SmenaDTO>>(hhh1);
        }



        public async Task<IEnumerable<SmenaDTO>> GetAll()
        {
            var hhh1 = _db.Smena
                            .Include("DepartmentFK")
                            .Include("DepartmentFK.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                            .ToListWithNoLock();
            return _mapper.Map<IEnumerable<Smena>, IEnumerable<SmenaDTO>>(hhh1);
        }


        public async Task<SmenaDTO> Update(SmenaDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.Smena
                            .Include("DepartmentFK")
                            .Include("DepartmentFK.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent")
                            .Include("DepartmentFK.DepartmentParent.DepartmentParent.DepartmentParent.DepartmentParent")
                .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {
                if (objectToUpdateDTO.DepartmentId == null || objectToUpdateDTO.DepartmentId == 0)
                {
                    objectToUpdate.DepartmentId = 0;
                    objectToUpdate.DepartmentFK = null;
                }
                else
                {
                    if (objectToUpdate.DepartmentId != objectToUpdateDTO.DepartmentId)
                    {
                        objectToUpdate.DepartmentId = objectToUpdateDTO.DepartmentId;
                        var objectDepartmentToUpdate = _db.MesDepartment.
                                FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.DepartmentId);
                        objectToUpdate.DepartmentFK = objectDepartmentToUpdate;
                    }
                }

                if (objectToUpdateDTO.Name != objectToUpdate.Name)
                    objectToUpdate.Name = objectToUpdateDTO.Name;

                if (objectToUpdateDTO.StartTime != objectToUpdate.StartTime)
                    objectToUpdate.StartTime = objectToUpdateDTO.StartTime;

                if (objectToUpdateDTO.HoursDuration != objectToUpdate.HoursDuration)
                    objectToUpdate.HoursDuration = objectToUpdateDTO.HoursDuration;

                if (objectToUpdateDTO.IsArchive != objectToUpdate.IsArchive)
                    objectToUpdate.IsArchive = objectToUpdateDTO.IsArchive;

                _db.Smena.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<Smena, SmenaDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id)
        {
            if (id != null && id != 0)
            {
                var objectToDelete = _db.Smena.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.Smena.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
