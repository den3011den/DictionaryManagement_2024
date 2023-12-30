using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IRoleToDepartmentRepository
    {
        public Task<RoleToDepartmentDTO> Get(Guid roleId, int departmentId);
        public Task<RoleToDepartmentDTO> GetById(int id);
        public Task<IEnumerable<RoleToDepartmentDTO>> GetAll();
        public Task<RoleToDepartmentDTO> Update(RoleToDepartmentDTO objDTO);
        public Task<RoleToDepartmentDTO> Create(RoleToDepartmentDTO objectToAddDTO);
        public Task<int> Delete(int Id);
    }
}
