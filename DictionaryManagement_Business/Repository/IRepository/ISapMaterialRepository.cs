using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISapMaterialRepository
    {
        public Task<SapMaterialDTO> Get(int Id);
        public Task<IEnumerable<SapMaterialDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<SapMaterialDTO> Update(SapMaterialDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<SapMaterialDTO> Create(SapMaterialDTO objectToAddDTO);
        public Task<SapMaterialDTO> GetByCode(string code = "");
        public Task<SapMaterialDTO> GetByName(string name = "");
        public Task<SapMaterialDTO> GetByShortName(string shortName = "");

    }
}
