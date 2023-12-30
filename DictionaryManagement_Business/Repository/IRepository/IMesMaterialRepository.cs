using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IMesMaterialRepository
    {
        public Task<MesMaterialDTO> Get(int Id);
        public Task<IEnumerable<MesMaterialDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<MesMaterialDTO> Update(MesMaterialDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<MesMaterialDTO> Create(MesMaterialDTO objectToAddDTO);
        public Task<MesMaterialDTO> GetByCode(string code = "");
        public Task<MesMaterialDTO> GetByName(string name = "");
        public Task<MesMaterialDTO> GetByShortName(string shortName = "");
    }
}
