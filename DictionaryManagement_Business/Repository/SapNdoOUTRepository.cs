using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;

namespace DictionaryManagement_Business.Repository
{
    public class SapNdoOUTRepository : ISapNdoOUTRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SapNdoOUTRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<SapNdoOUTDTO> Create(SapNdoOUTDTO objectToAddDTO)
        {

            SapNdoOUT objectToAdd = new SapNdoOUT();

            objectToAdd.Id = objectToAddDTO.Id;

            if (objectToAddDTO.AddTime == null)
                objectToAdd.AddTime = DateTime.Now;
            else
                objectToAdd.AddTime = objectToAddDTO.AddTime;

            objectToAdd.ValueTime = objectToAddDTO.ValueTime;
            objectToAdd.Value = objectToAddDTO.Value;
            objectToAdd.TagName = objectToAddDTO.TagName;
            objectToAdd.SapGone = objectToAddDTO.SapGone;
            objectToAdd.SapGoneTime = objectToAddDTO.SapGoneTime;

            var addedSapNdoOUT = _db.SapNdoOUT.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<SapNdoOUT, SapNdoOUTDTO>(addedSapNdoOUT.Entity);
        }


        public async Task<SapNdoOUTDTO> GetById(Int64 id)
        {
            var objToGet = _db.SapNdoOUT
                            .FirstOrDefaultWithNoLock(u => u.Id == id);
            if (objToGet != null)
            {
                return _mapper.Map<SapNdoOUT, SapNdoOUTDTO>(objToGet);
            }
            return null;
        }


        public async Task<IEnumerable<SapNdoOUTDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime, string intervalMode)
        {

            if (startTime == null)
                startTime = DateTime.MinValue;
            if (startTime == null)
                endTime = DateTime.MaxValue;

            switch (intervalMode.Trim().ToUpper())
            {
                case "ADDTIME":
                    var hhh1 = _db.SapNdoOUT
                        .Where(u => u.AddTime >= startTime && u.AddTime <= endTime).ToListWithNoLock();
                    return _mapper.Map<IEnumerable<SapNdoOUT>, IEnumerable<SapNdoOUTDTO>>(hhh1);

                case "VALUETIME":
                    var hhh2 = _db.SapNdoOUT
                        .Where(u => u.ValueTime >= startTime && u.ValueTime <= endTime).ToListWithNoLock();
                    return _mapper.Map<IEnumerable<SapNdoOUT>, IEnumerable<SapNdoOUTDTO>>(hhh2);
                default:
                    return null;
            }

        }


        public async Task<SapNdoOUTDTO> Update(SapNdoOUTDTO objectToUpdateDTO)
        {
            var objectToUpdate = _db.SapNdoOUT
               .FirstOrDefaultWithNoLock(u => u.Id == objectToUpdateDTO.Id);

            if (objectToUpdate != null)
            {

                if (objectToUpdateDTO.AddTime != objectToUpdate.AddTime)
                    objectToUpdate.AddTime = objectToUpdateDTO.AddTime;

                if (objectToUpdateDTO.ValueTime != objectToUpdate.ValueTime)
                    objectToUpdate.ValueTime = objectToUpdateDTO.ValueTime;

                if (objectToUpdateDTO.Value != objectToUpdate.Value)
                    objectToUpdate.Value = objectToUpdateDTO.Value;

                if (objectToUpdateDTO.TagName != objectToUpdate.TagName)
                    objectToUpdate.TagName = objectToUpdateDTO.TagName;

                if (objectToUpdateDTO.SapGone != objectToUpdate.SapGone)
                    objectToUpdate.SapGone = objectToUpdateDTO.SapGone;

                if (objectToUpdateDTO.SapGoneTime != objectToUpdate.SapGoneTime)
                    objectToUpdate.SapGoneTime = objectToUpdateDTO.SapGoneTime;


                _db.SapNdoOUT.Update(objectToUpdate);
                _db.SaveChanges();
                return _mapper.Map<SapNdoOUT, SapNdoOUTDTO>(objectToUpdate);
            }
            return objectToUpdateDTO;

        }

        public async Task<Int64> Delete(Int64 id)
        {
            if (id != null && id != 0)
            {
                var objectToDelete = _db.SapNdoOUT.FirstOrDefaultWithNoLock(u => u.Id == id);
                if (objectToDelete != null)
                {
                    _db.SapNdoOUT.Remove(objectToDelete);
                    return _db.SaveChanges();
                }
            }
            return 0;
        }
    }
}
