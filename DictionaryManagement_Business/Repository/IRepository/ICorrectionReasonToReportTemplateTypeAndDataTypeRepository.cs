using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ICorrectionReasonToReportTemplateTypeAndDataTypeRepository
    {
        public Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> GetByCorrectionReasonIdAndReportTemplateTypeIdAndDataType(int correctionReasonId, int reportTemplateTypeId, int dataTypeId);
        public Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> GetById(int id);
        public Task<IEnumerable<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>?> GetByCorrectionReasonId(int correctionReasonId);
        public Task<IEnumerable<CorrectionReasonToReportTemplateTypeAndDataTypeDTO>> GetAll();
        public Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> Update(CorrectionReasonToReportTemplateTypeAndDataTypeDTO objDTO);
        public Task<CorrectionReasonToReportTemplateTypeAndDataTypeDTO?> Create(CorrectionReasonToReportTemplateTypeAndDataTypeDTO objectToAddDTO);
        public Task<int> Delete(int Id);
    }
}
