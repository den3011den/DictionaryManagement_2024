using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class SapToMesMaterialMappingRepository : ISapToMesMaterialMappingRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SapToMesMaterialMappingRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SapToMesMaterialMappingDTO> Create(SapToMesMaterialMappingDTO objectToAddDTO)
        {
            //var objectToAdd = _mapper.Map<UnitOfMeasureSapToMesMappingDTO, UnitOfMeasureSapToMesMapping>(objectToAddDTO);

            SapToMesMaterialMapping objectToAdd = new SapToMesMaterialMapping();

            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.SapMaterialId = objectToAddDTO.SapMaterialId;
            objectToAdd.MesMaterialId = objectToAddDTO.MesMaterialId;


            var addedSapToMesMaterialMapping = _db.SapToMesMaterialMapping.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<SapToMesMaterialMapping, SapToMesMaterialMappingDTO>(addedSapToMesMaterialMapping.Entity);
        }

        public async Task<SapToMesMaterialMappingDTO> Get(int sapMaterialId, int mesMaterialId)
        {
            var objToGet = _db.SapToMesMaterialMapping.Include("SapMaterial").Include("MesMaterial").
                            FirstOrDefaultWithNoLock(u => u.SapMaterialId == sapMaterialId && u.MesMaterialId == mesMaterialId);
            if (objToGet != null)
            {
                return _mapper.Map<SapToMesMaterialMapping, SapToMesMaterialMappingDTO>(objToGet);
            }
            return null;
        }

        public async Task<SapToMesMaterialMappingDTO> GetById(int id)
        {
            var objToGet = _db.SapToMesMaterialMapping.Include("SapMaterial").Include("MesMaterial").
                            FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<SapToMesMaterialMapping, SapToMesMaterialMappingDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<SapToMesMaterialMappingDTO>> GetAll()
        {
            var hhh = _db.SapToMesMaterialMapping.Include("SapMaterial").Include("MesMaterial").ToListWithNoLock();
            return _mapper.Map<IEnumerable<SapToMesMaterialMapping>, IEnumerable<SapToMesMaterialMappingDTO>>(hhh);

        }

        public async Task<SapToMesMaterialMappingDTO> Update(SapToMesMaterialMappingDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.SapToMesMaterialMapping.Include("SapMaterial").Include("MesMaterial").
                    FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);
            if (objectToUpdate != null)
            {

                if (objectToUpdate.SapMaterialId != objectToUpdateDTO.SapMaterialDTO.Id)
                {
                    objectToUpdate.SapMaterialId = objectToUpdateDTO.SapMaterialDTO.Id;
                    objectToUpdate.SapMaterial = _mapper.Map<SapMaterialDTO, SapMaterial>(objectToUpdateDTO.SapMaterialDTO);
                }
                if (objectToUpdate.MesMaterialId != objectToUpdateDTO.MesMaterialDTO.Id)
                {
                    objectToUpdate.MesMaterialId = objectToUpdateDTO.MesMaterialDTO.Id;
                    objectToUpdate.MesMaterial = _mapper.Map<MesMaterialDTO, MesMaterial>(objectToUpdateDTO.MesMaterialDTO);
                }
                _db.SapToMesMaterialMapping.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapToMesMaterialMapping, SapToMesMaterialMappingDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<int> Delete(int id)
        {
            if (id > 0)
            {
                var objectToDelete = _db.SapToMesMaterialMapping.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.SapToMesMaterialMapping.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;

        }
    }
}
