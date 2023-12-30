using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IADGroupRepository
    {
        public Task<ADGroupDTO> Get(Guid Id);
        public Task<IEnumerable<ADGroupDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<ADGroupDTO> Update(ADGroupDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<ADGroupDTO> Create(ADGroupDTO objectToAddDTO);
        public Task<ADGroupDTO> GetByName(string name = "");

    }
}
