using DictionaryManagement_DataAccess.Data.IntDB;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISapToMesMaterialMappingRepository
    {
        public Task<SapToMesMaterialMappingDTO> Get(int sapMaterialId, int mesMaterialId);
        public Task<SapToMesMaterialMappingDTO> GetById(int id);
        public Task<IEnumerable<SapToMesMaterialMappingDTO>> GetAll();
        public Task<SapToMesMaterialMappingDTO> Update(SapToMesMaterialMappingDTO objDTO);
        public Task<SapToMesMaterialMappingDTO> Create(SapToMesMaterialMappingDTO objectToAddDTO);
        public Task<int> Delete(int Id);
    }
}
