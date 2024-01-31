using ClosedXML.Excel;
using DictionaryManagement_Business.Repository;
using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Server.Extensions.Repository.IRepository
{
    public interface ILoadFromExcelRepository
    {
        public Task<string> MaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MaterialDTO>? materialList);
        public Task<string> SapEquipmentReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapEquipmentDTO>? sapEquipmentList);
        public Task<string> MesParamReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MesParamDTO>? mesParamList);
        public Task<string> MesNdoStocksReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MesNdoStocksDTO>? mesNdoStocksList);
        public Task<string> SapNdoOUTReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapNdoOUTDTO>? sapNdoOUTList);
        public Task<string> SapUnitOfMeasureReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);

        public Task<string> MesMovementsReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MesMovementsDTO>? mesMovementsListDTO);
        public Task<string> SapMovementsOUTReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapMovementsOUTDTO>? sapMovementsOUTListDTO);
        public Task<string> SapMovementsINReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapMovementsINDTO>? sapMovementsINListDTO);
        public Task<string> UserReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<UserDTO>? userListDTO);
        public Task<string> ADGroupReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<ADGroupDTO>? adGroupListDTO);

        public Task<bool> MaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository);
        public Task<bool> SapEquipmentExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository);
        public Task<bool> MesParamExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository);

        public Task<bool> SapNdoOUTExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> MesNdoStocksExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> MesMovementsExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> SapMovementsOUTExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> SapMovementsINExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> UsersExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> ADGroupsExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

    }
}
