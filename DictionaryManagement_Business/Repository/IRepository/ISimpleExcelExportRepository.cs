
using ClosedXML.Excel;
using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Repository.IRepository
{
    public interface ISimpleExcelExportRepository
    {
        public Task<string> GenerateExcelReportEntity(string filename, IEnumerable<ReportEntityDTO> data);
        public Task<string> GenerateExcelMesParam(string filename, IEnumerable<MesParamDTO> data);
        public Task<string> GenerateExcelSapEquipment(string filename, IEnumerable<SapEquipmentDTO> data);
        public Task<string> GenerateExcelSapMaterial(string filename, IEnumerable<SapMaterialDTO> data);
        public Task<string> GenerateExcelUsers(string filename, IEnumerable<UserDTO> data);
        public Task<string> GenerateExcelADGroup(string filename, IEnumerable<ADGroupDTO> data);
        public Task<string> GenerateExcelRole(string filename, IEnumerable<RoleVMDTO> data);
        public Task<string> GenerateExcelMesDepartments(string filename, IEnumerable<MesDepartmentVMDTO> mesDepartmentVMDTOList, int maxLevel);
        public Task<string> GenerateExcelMesNdoStocks(string filename, IEnumerable<MesNdoStocksDTO> data);
        public Task<string> GenerateExcelSapNdoOUT(string filename, IEnumerable<SapNdoOUTDTO> data);
        public Task<string> GenerateExcelMesMovements(string filename, IEnumerable<MesMovementsDTO> data);
        public Task<string> GenerateExcelSapMovementsOUT(string filename, IEnumerable<SapMovementsOUTDTO> data);
        public Task<string> GenerateExcelSapMovementsIN(string filename, IEnumerable<SapMovementsINDTO> data);
        public Task<string> GenerateExcelLogEvent(string filename, IEnumerable<LogEventDTO> data);
        public Task<string> GenerateExcelMesMaterial(string filename, IEnumerable<MesMaterialDTO> data);
        public Task<string> GenerateExcelSettings(string filename, IEnumerable<SettingsDTO> data);
        public Task<string> GenerateExcelSapUnitOfMeasure(string filename, IEnumerable<SapUnitOfMeasureDTO> data);
        public Task<string> GenerateExcelMesUnitOfMeasure(string filename, IEnumerable<MesUnitOfMeasureDTO> data);
        public Task<string> GenerateExcelCorrectionReason(string filename, IEnumerable<CorrectionReasonDTO> data);
        public Task<string> GenerateExcelMesParamSourceType(string filename, IEnumerable<MesParamSourceTypeDTO> data);
        public Task<string> GenerateExcelDataType(string filename, IEnumerable<DataTypeDTO> data);
        public Task<string> GenerateExcelDataSource(string filename, IEnumerable<DataSourceDTO> data);
        public Task<string> GenerateExcelReportTemplateType(string filename, IEnumerable<ReportTemplateTypeDTO> data);
        public Task<string> GenerateExcelReportTemplate(string filename, IEnumerable<ReportTemplateDTO> data);
        public Task<string> GenerateExcelLogEventType(string filename, IEnumerable<LogEventTypeDTO> data);
        public Task<string> GenerateExcelSmena(string filename, IEnumerable<SmenaDTO> data);
        public Task<string> GenerateExcelScheduler(string filename, IEnumerable<SchedulerDTO> data);
        public Task<Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook?>> GetSheetData(ReportEntityDTO? reportEntityDTO, string sheetSettingName, XLWorkbook? workbook);

    }
}
