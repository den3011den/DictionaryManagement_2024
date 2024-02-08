using DictionaryManagement_Models.IntDBModels;
using static DictionaryManagement_Common.SD;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISapEquipmentRepository
    {
        public Task<SapEquipmentDTO> Get(int Id);
        public Task<IEnumerable<SapEquipmentDTO>> GetAll(SelectDictionaryScope selectDictionaryScope = SelectDictionaryScope.All);
        public Task<SapEquipmentDTO> Update(SapEquipmentDTO objDTO, UpdateMode updateMode = UpdateMode.Update);
        public Task<SapEquipmentDTO> Create(SapEquipmentDTO objectToAddDTO);
        public Task<SapEquipmentDTO> GetByResource(string erpPlantId = "", string erpId = "");
        public Task<SapEquipmentDTO> GetByName(string name = "");
        public Task<IEnumerable<SapEquipmentDTO>?> GetListByName(string name = "");

    }
}
