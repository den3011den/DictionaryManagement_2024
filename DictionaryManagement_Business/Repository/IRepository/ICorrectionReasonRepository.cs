using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ICorrectionReasonRepository
    {
        public Task<CorrectionReasonDTO> Get(int Id);
        public Task<IEnumerable<CorrectionReasonDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<CorrectionReasonDTO> Update(CorrectionReasonDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<CorrectionReasonDTO> Create(CorrectionReasonDTO objectToAddDTO);
        public Task<CorrectionReasonDTO> GetByName(string Name);
    }
}
