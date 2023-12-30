using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface IReportTemplateTypeToRoleRepository
    {
        public Task<ReportTemplateTypeTоRoleDTO> Get(int reportTemplateTypeId, Guid roleId);
        public Task<ReportTemplateTypeTоRoleDTO> GetById(int id);
        public Task<IEnumerable<ReportTemplateTypeTоRoleDTO>> GetAll();
        public Task<ReportTemplateTypeTоRoleDTO> Update(ReportTemplateTypeTоRoleDTO objDTO);
        public Task<ReportTemplateTypeTоRoleDTO> Create(ReportTemplateTypeTоRoleDTO objectToAddDTO);
        public Task<int> Delete(int Id);
    }
}
