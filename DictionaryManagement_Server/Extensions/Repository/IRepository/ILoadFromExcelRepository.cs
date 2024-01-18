using ClosedXML.Excel;
using DictionaryManagement_Business.Repository;
using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Server.Extensions.Repository.IRepository
{
    public interface ILoadFromExcelRepository
    {
        public Task<string> MaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> SapEquipmentReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> MesParamReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> MesNdoStocksReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MesNdoStocksDTO>? mesNdoStocksList);
        public Task<string> SapNdoOUTReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapNdoOUTDTO>? sapNdoOUTList);

        public Task<bool> MaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository);
        public Task<bool> SapEquipmentExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository);
        public Task<bool> MesParamExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository);


    }
}
