﻿using ClosedXML.Excel;
using DictionaryManagement_Business.Repository;

namespace DictionaryManagement_Server.Extensions.Repository.IRepository
{
    public interface ILoadFromExcelRepository
    {
        public Task<string> SapMaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> MesMaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> SapEquipmentReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);
        public Task<string> MesParamReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet);

        public Task<bool> SapMaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository);

        public Task<bool> MesMaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository);

        public Task<bool> SapEquipmentExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

        public Task<bool> MesParamExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository);

    }
}
