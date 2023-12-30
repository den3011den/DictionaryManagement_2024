using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISmenaRepository
    {

        public Task<IEnumerable<SmenaDTO>> GetAllByDepartmentId(int departmentId);
        public Task<SmenaDTO> GetById(int id);
        public Task<IEnumerable<SmenaDTO>> GetAll();
        public Task<SmenaDTO> Update(SmenaDTO objDTO);
        public Task<SmenaDTO> Create(SmenaDTO objectToAddDTO);
        public Task<int> Delete(int id);
    }
}
