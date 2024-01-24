using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class SapEquipmentRepository : ISapEquipmentRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SapEquipmentRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SapEquipmentDTO> Create(SapEquipmentDTO objectToAddDTO)
        {
            var objectToAdd = _mapper.Map<SapEquipmentDTO, SapEquipment>(objectToAddDTO);
            var addedSapEquipment = _db.SapEquipment.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<SapEquipment, SapEquipmentDTO>(addedSapEquipment.Entity);
        }

        public async Task<SapEquipmentDTO> Get(int Id)
        {
            if (Id > 0)
            {
                var objToGet = _db.SapEquipment.FirstOrDefaultWithNoLock(u => u.Id == Id);
                if (objToGet != null)
                {
                    return _mapper.Map<SapEquipment, SapEquipmentDTO>(objToGet);
                }
            }
            return null;
        }

        public async Task<IEnumerable<SapEquipmentDTO>> GetAll(SD.SelectDictionaryScope selectDictionaryScope = SD.SelectDictionaryScope.All)
        {
            if (selectDictionaryScope == SD.SelectDictionaryScope.All)
            {
                //var fff = _db.SapEquipment;
                return _mapper.Map<IEnumerable<SapEquipment>, IEnumerable<SapEquipmentDTO>>(_db.SapEquipment.ToListWithNoLock());
            }
            if (selectDictionaryScope == SD.SelectDictionaryScope.ArchiveOnly)
                return _mapper.Map<IEnumerable<SapEquipment>, IEnumerable<SapEquipmentDTO>>(_db.SapEquipment.Where(u => u.IsArchive == true).ToListWithNoLock());
            if (selectDictionaryScope == SD.SelectDictionaryScope.NotArchiveOnly)
                return _mapper.Map<IEnumerable<SapEquipment>, IEnumerable<SapEquipmentDTO>>(_db.SapEquipment.Where(u => u.IsArchive != true).ToListWithNoLock());
            return _mapper.Map<IEnumerable<SapEquipment>, IEnumerable<SapEquipmentDTO>>(_db.SapEquipment.ToListWithNoLock());
        }

        public async Task<SapEquipmentDTO> GetByResource(string erpPlantId = "", string erpId = "")
        {
            var objToGet = _db.SapEquipment.FirstOrDefaultWithNoLock(u => ((u.ErpPlantId.Trim().ToUpper() == erpPlantId.Trim().ToUpper()) && (u.ErpId.Trim().ToUpper() == erpId.Trim().ToUpper())));
            if (objToGet != null)
            {
                return _mapper.Map<SapEquipment, SapEquipmentDTO>(objToGet);
            }
            return null;
        }
        public async Task<SapEquipmentDTO> GetByName(string name = "")
        {
            var objToGet = _db.SapEquipment.FirstOrDefaultWithNoLock(u => ((u.Name.Trim().ToUpper()) == (name.Trim().ToUpper())));
            if (objToGet != null)
            {
                return _mapper.Map<SapEquipment, SapEquipmentDTO>(objToGet);
            }
            return null;
        }


        public async Task<SapEquipmentDTO> Update(SapEquipmentDTO objectToUpdateDTO, SD.UpdateMode updateMode = SD.UpdateMode.Update)
        {
            var objectToUpdate = _db.SapEquipment.FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {
                if (updateMode == SD.UpdateMode.Update)
                {
                    if (objectToUpdate.ErpPlantId != objectToUpdateDTO.ErpPlantId)
                        objectToUpdate.ErpPlantId = objectToUpdateDTO.ErpPlantId;
                    if (objectToUpdate.ErpId != objectToUpdateDTO.ErpId)
                        objectToUpdate.ErpId = objectToUpdateDTO.ErpId;
                    if (objectToUpdate.Name != objectToUpdateDTO.Name)
                        objectToUpdate.Name = objectToUpdateDTO.Name;
                    if (objectToUpdate.IsWarehouse != objectToUpdateDTO.IsWarehouse)
                        objectToUpdate.IsWarehouse = objectToUpdateDTO.IsWarehouse;
                }
                if (updateMode == SD.UpdateMode.MoveToArchive)
                {
                    objectToUpdate.IsArchive = true;
                }
                if (updateMode == SD.UpdateMode.RestoreFromArchive)
                {
                    objectToUpdate.IsArchive = false;
                }
                _db.SapEquipment.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapEquipment, SapEquipmentDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;
        }

        public async Task<IEnumerable<SapEquipmentDTO>?> GetListByName(string name)
        {
            return _mapper.Map<IEnumerable<SapEquipment>, IEnumerable<SapEquipmentDTO>>(
                _db.SapEquipment.Where(u => !String.IsNullOrEmpty(u.Name.Trim()) && u.Name.Trim().ToUpper() == name.Trim().ToUpper()).ToListWithNoLock()
                );
        }
    }
}
