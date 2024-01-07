using ClosedXML.Excel;

namespace DictionaryManagement_Server.Extensions.Repository.IRepository
{
    public interface ILoadFromExcelRepository
    {
        public Task<string> SapMaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> MesMaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> SapEquipmentReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> MesParamReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
    }
}
