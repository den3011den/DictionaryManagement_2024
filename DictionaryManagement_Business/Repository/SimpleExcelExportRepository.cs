using AutoMapper;
using ClosedXML.Excel;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using DocumentFormat.OpenXml;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using System.Globalization;

namespace DictionaryManagement_Business.Repository
{
    public class SimpleExcelExportRepository : ISimpleExcelExportRepository
    {
        private readonly ISettingsRepository _settingsRepository;
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public SimpleExcelExportRepository(ISettingsRepository settingsRepository, IntDBApplicationDbContext db, IMapper mapper)
        {
            _settingsRepository = settingsRepository;
            _db = db;
            _mapper = mapper;
        }

        public async Task<string> GenerateExcelReportEntity(string filename, IEnumerable<ReportEntityDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("ReportEntity");

                int excelRowNum = 1;
                int excelColNum = 1;


                ws.Cell(excelRowNum, excelColNum).Value = "ИД экземпляра отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД шаблона отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Тип шаблона отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Начало интервала отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Окончание интервала отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД производства у экземпляра отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование производства у экземпляра отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД производства у шаблона отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование производства у шаблона отчёта";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время скачивания (Download)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Кто скачал (Download)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время загрузки (Upload)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Кто загрузил (Upload)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Загружен в СИР в режиме переотправки в SAP (ResendMode)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Даты переотправки в SAP (ReportEntityResendDates)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (ReportEntityDTO reportEntity in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportTemplateId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportTemplateDTOFK.ReportTemplateTypeDTOFK.Name.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportTimeStart == null ? "" : ((DateTime)reportEntity.ReportTimeStart).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportTimeEnd == null ? "" : ((DateTime)reportEntity.ReportTimeEnd).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportDepartmentId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportDepartmentDTOFK == null ? "" : reportEntity.ReportDepartmentDTOFK.ToStringHierarchyShortName.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportTemplateDTOFK.MesDepartmentDTOFK.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ReportTemplateDTOFK.MesDepartmentDTOFK.ToStringHierarchyShortName.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.DownloadTime == null ? "" : ((DateTime)reportEntity.DownloadTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.DownloadUserDTOFK == null ? "" : reportEntity.DownloadUserDTOFK.UserName.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.UploadTime == null ? "" : ((DateTime)reportEntity.UploadTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.UploadUserDTOFK == null ? "" : reportEntity.UploadUserDTOFK.UserName.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntity.ResendMode == true ? "Да" : "";

                    string reportEntityResendDatesListString = "";
                    if (reportEntity.ReportEntityResendDatesListDTO != null)
                    {
                        int datesInRowCount = 1;
                        foreach (ReportEntityResendDatesDTO reportEntityResendDatesDTO in reportEntity.ReportEntityResendDatesListDTO.OrderBy(u => u.ResendDate))
                        {
                            if (String.IsNullOrEmpty(reportEntityResendDatesListString))
                                reportEntityResendDatesListString = reportEntityResendDatesDTO.ResendDate.ToString("dd.MM.yyyy");
                            else
                            {
                                if (datesInRowCount % 4 == 0)
                                {
                                    reportEntityResendDatesListString = reportEntityResendDatesListString + "\n" + reportEntityResendDatesDTO.ResendDate.ToString("dd.MM.yyyy");
                                    datesInRowCount = 1;
                                }
                                else
                                {
                                    reportEntityResendDatesListString = reportEntityResendDatesListString + "     " + reportEntityResendDatesDTO.ResendDate.ToString("dd.MM.yyyy");
                                }
                            }
                            datesInRowCount++;
                        }
                    }

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportEntityResendDatesListString;
                    ws.Cell(excelRowNum, excelColNum).Style.Alignment.WrapText = true;

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelMesParam(string filename, IEnumerable<MesParamDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesParam");

                int excelRowNum = 1;
                int excelColNum = 1;


                ws.Cell(excelRowNum, excelColNum).Value = "ИД тэга СИР (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код тэга (Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование тэга (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Точка измерения (TI)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование точки измерения (NameTI)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Технологическое место (TM)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование технологического места (NameTM)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Источник (MesParamSourceType.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Тэг источника (MesParamSourceLink)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код производства (DepartmentId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование производства (MesDepartment.ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД источника Sap (SapEquipmentIdSource)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода-источника Sap (SapEquipment.ErpPlantId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код склада/ресурса-источника Sap (SapEquipment.ErpId)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование источника Sap (SapEquipment.ErpPlantId + ErpId + Name)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД приёмника Sap (SapEquipmentIdDest)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода-приёмника Sap (SapEquipment.ErpPlantId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код склада/ресурса-приёмника Sap (SapEquipment.ErpId)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование приёмника Sap (SapEquipment.ErpPlantId + ErpId + Name)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД материала Sap (SapMaterialId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код АСВ НСИ материала Sap (SapMaterial.Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование материала Sap (SapMaterial.Name)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД ед.изм. Sap (SapUnitOfMeasureId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование ед.изм. Sap (SapUnitOfMeasure.ShortName)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время запроса данных в прошлое в днях (DaysRequestInPast)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Читать из SAP (NeedReadFromSap)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Передавать в SAP (NeedWriteToSap)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Читать из MES (NeedReadFromMes)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Передавать в MES (NeedWriteToMes)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является тэгом НДО (IsNdo)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Коэф. пересчёта данных по тэгу из ед.изм. MES в ед.изм СИР (MesToSirUnitOfMeasureKoef)";

                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (MesParamDTO mesParamDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.Name == null ? "" : mesParamDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.TI == null ? "" : mesParamDTO.TI;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.NameTI == null ? "" : mesParamDTO.NameTI;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.TM == null ? "" : mesParamDTO.TM;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.NameTM == null ? "" : mesParamDTO.NameTM;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.Description == null ? "" : mesParamDTO.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.MesParamSourceTypeDTOFK == null ? "" : mesParamDTO.MesParamSourceTypeDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.MesParamSourceLink == null ? "" : mesParamDTO.MesParamSourceLink;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.DepartmentId == null ? "" : mesParamDTO.DepartmentId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.MesDepartmentDTOFK == null ? "" : mesParamDTO.MesDepartmentDTOFK.ToStringHierarchyShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentIdSource == null ? "" : mesParamDTO.SapEquipmentIdSource.ToString();
                    excelColNum++;

                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentSourceDTOFK == null ? "" : mesParamDTO.SapEquipmentSourceDTOFK.ErpPlantId;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentSourceDTOFK == null ? "" : mesParamDTO.SapEquipmentSourceDTOFK.ErpId;
                    excelColNum++;

                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentSourceDTOFK == null ? "" : mesParamDTO.SapEquipmentSourceDTOFK.ErpPlantId + "|" + mesParamDTO.SapEquipmentSourceDTOFK.ErpId + " " + mesParamDTO.SapEquipmentSourceDTOFK.Name;

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentIdDest == null ? "" : mesParamDTO.SapEquipmentIdDest.ToString();
                    excelColNum++;

                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentDestDTOFK == null ? "" : mesParamDTO.SapEquipmentDestDTOFK.ErpPlantId;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentDestDTOFK == null ? "" : mesParamDTO.SapEquipmentDestDTOFK.ErpId;
                    excelColNum++;

                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapEquipmentDestDTOFK == null ? "" : mesParamDTO.SapEquipmentDestDTOFK.ErpPlantId + "|" + mesParamDTO.SapEquipmentDestDTOFK.ErpId + " " + mesParamDTO.SapEquipmentDestDTOFK.Name;

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapMaterialId == null ? "" : mesParamDTO.SapMaterialId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapMaterialDTOFK == null ? "" : mesParamDTO.SapMaterialDTOFK.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapMaterialDTOFK == null ? "" : mesParamDTO.SapMaterialDTOFK.Name;

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapUnitOfMeasureId == null ? "" : mesParamDTO.SapUnitOfMeasureId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.SapUnitOfMeasureDTOFK == null ? "" : mesParamDTO.SapUnitOfMeasureDTOFK.ShortName;

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.DaysRequestInPast == null ? "" : mesParamDTO.DaysRequestInPast.ToString();

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.NeedReadFromSap == true ? "Да" : "";

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.NeedWriteToSap == true ? "Да" : "";

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.NeedReadFromMes == true ? "Да" : "";

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.NeedWriteToMes == true ? "Да" : "";

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.IsNdo == true ? "Да" : "";

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.MesToSirUnitOfMeasureKoef == null ? "" : mesParamDTO.MesToSirUnitOfMeasureKoef.ToString();

                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelSapEquipment(string filename, IEnumerable<SapEquipmentDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("SapEquipment");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода SAP (ErpPlantId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код ресурса/склада SAP (ErpId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является складом (IsWarehouse)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SapEquipmentDTO sapEquipmentDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTO.ErpPlantId;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTO.ErpId;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTO.IsWarehouse == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelSapMaterial(string filename, IEnumerable<SapMaterialDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("SapMaterial");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код АСВ НСИ (Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сокр. наименование (ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SapMaterialDTO sapMaterialDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = sapMaterialDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMaterialDTO.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMaterialDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMaterialDTO.ShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMaterialDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelUsers(string filename, IEnumerable<UserDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("Users");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Логин (Login)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (UserName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Синхронизируется с AD (IsSyncWithAD)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сервисный пользователь (IsServiceUser)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время последней синхронизации с AD (SyncWithADGroupsLastTime)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (UserDTO userDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.Login;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.UserName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.Description == null ? "" : userDTO.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.IsSyncWithAD == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.IsServiceUser == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.IsArchive == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = userDTO.SyncWithADGroupsLastTime.ToString("dd.MM.yyyy HH:mm:ss.fff");

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelADGroup(string filename, IEnumerable<ADGroupDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("ADGroups");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (ADGroupDTO adGroupDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = adGroupDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = adGroupDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = adGroupDTO.Description == null ? "" : adGroupDTO.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = adGroupDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelRole(string filename, IEnumerable<RoleVMDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var wsRole = wbook.AddWorksheet("Roles");
                var wsUserToRole = wbook.AddWorksheet("UserToRole");
                var wsReportTemplateTypeTоRole = wbook.AddWorksheet("ReportTemplateTypeTоRole");
                var wsRoleToADGroup = wbook.AddWorksheet("RoleToADGroup");
                var wsRoleToDepartment = wbook.AddWorksheet("RoleToDepartment");

                int excelRowNum = 1;
                int excelColNum = 1;

                wsRole.Cell(excelRowNum, excelColNum).Value = "ИД (Id)";
                excelColNum++;
                wsRole.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                wsRole.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                wsRole.Cell(excelRowNum, excelColNum).Value = "Админка полный доступ (IsAdmin)";
                excelColNum++;
                wsRole.Cell(excelRowNum, excelColNum).Value = "Админка только чтение (IsAdminReadOnly)";
                excelColNum++;
                wsRole.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";

                wsRole.Row(excelRowNum).Style.Font.SetBold(true);
                wsRole.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                wsRole.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 1;
                excelColNum = 1;

                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "ИД связки UserToRole (UserToRole.Id)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "ИД роли СИР (Role.Id)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Наименование роли СИР (Role.Name)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Описание роли СИР (Role.Description)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "В архиве (Role.IsArchive)";

                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "ИД пользователя (User.Id)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Логин пользователя (User.Login)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Наименование пользователя (User.UserName)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Описание пользователя (User.Description)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Синхронизируется с AD (User.IsSyncWithAD)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Сервисный пользователь (User.IsServiceUser)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "Время последней синхронизации с AD (User.SyncWithADGroupsLastTime)";
                excelColNum++;
                wsUserToRole.Cell(excelRowNum, excelColNum).Value = "В архиве (User.IsArchive)";

                wsUserToRole.Row(excelRowNum).Style.Font.SetBold(true);
                wsUserToRole.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                wsUserToRole.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 1;
                excelColNum = 1;

                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "ИД связки ReportTemplateTypeTоRole (ReportTemplateTypeTоRole.Id)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "ИД роли СИР (Role.Id)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "Наименование роли СИР (Role.Name)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "Описание роли СИР (Role.Description)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "В архиве (Role.IsArchive)";

                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "ИД типа шаблона отчёта (ReportTemplateType.Id)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "Наименование типа шаблона отчёта (ReportTemplateType.Name)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "Расчёт автоматически (ReportTemplateType.NeedAutoCalc)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "Право на чтение (ReportTemplateTypeTоRole.CanDownload)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "Право на запись (ReportTemplateTypeTоRole.CanUpload)";
                excelColNum++;
                wsReportTemplateTypeTоRole.Cell(excelRowNum, excelColNum).Value = "В архиве (ReportTemplateType.IsArchive)";

                wsReportTemplateTypeTоRole.Row(excelRowNum).Style.Font.SetBold(true);
                wsReportTemplateTypeTоRole.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                wsReportTemplateTypeTоRole.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 1;
                excelColNum = 1;

                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "ИД связки RoleToADGroup (RoleToADGroup.Id)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "ИД роли СИР (Role.Id)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "Наименование роли СИР (Role.Name)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "Описание роли СИР (Role.Description)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "В архиве (Role.IsArchive)";

                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "ИД группы AD (ADGroup.Id)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "Наименование AD группы (ADGroup.Name)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "Описание AD группы (ADGroup.Description)";
                excelColNum++;
                wsRoleToADGroup.Cell(excelRowNum, excelColNum).Value = "В архиве (ADGroup.IsArchive)";

                wsRoleToADGroup.Row(excelRowNum).Style.Font.SetBold(true);
                wsRoleToADGroup.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                wsRoleToADGroup.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 1;
                excelColNum = 1;

                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "ИД связки RoleToDepartment (RoleToDepartment.Id)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "ИД роли СИР (Role.Id)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "Наименование роли СИР (Role.Name)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "Описание роли СИР (Role.Description)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "В архиве (Role.IsArchive)";

                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "ИД производства (MesDepartment.Id)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "Код производства (MesDepartment.MesCode)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "Наименование производства (MesDepartment.Name)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "Сокр. наименование производства (MesDepartment.ShortName)";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "Сокр. наименование производства - полный путь";
                excelColNum++;
                wsRoleToDepartment.Cell(excelRowNum, excelColNum).Value = "В архиве (MesDepartment.IsArchive)";

                wsRoleToDepartment.Row(excelRowNum).Style.Font.SetBold(true);
                wsRoleToDepartment.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                wsRoleToDepartment.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;

                int wsRoleRowNum = 2;
                int wsRoleColNum = 1;

                int wsUserToRoleRowNum = 2;
                int wsUserToRoleColNum = 1;

                int wsReportTemplateTypeTоRoleRowNum = 2;
                int wsReportTemplateTypeTоRoleColNum = 1;

                int wsRoleToADGroupRowNum = 2;
                int wsRoleToADGroupColNum = 1;

                int wsRoleToDepartmentRowNum = 2;
                int wsRoleToDepartmentColNum = 1;

                foreach (RoleVMDTO roleVMDTO in data)
                {
                    wsRoleColNum = 1;

                    wsRole.Cell(wsRoleRowNum, wsRoleColNum).Value = roleVMDTO.Id.ToString();
                    wsRoleColNum++;
                    wsRole.Cell(wsRoleRowNum, wsRoleColNum).Value = roleVMDTO.Name;
                    wsRoleColNum++;
                    wsRole.Cell(wsRoleRowNum, wsRoleColNum).Value = roleVMDTO.Description == null ? "" : roleVMDTO.Description;
                    wsRoleColNum++;
                    wsRole.Cell(wsRoleRowNum, wsRoleColNum).Value = roleVMDTO.IsAdmin == true ? "Да" : "";
                    wsRoleColNum++;
                    wsRole.Cell(wsRoleRowNum, wsRoleColNum).Value = roleVMDTO.IsAdminReadOnly == true ? "Да" : "";
                    wsRoleColNum++;
                    wsRole.Cell(wsRoleRowNum, wsRoleColNum).Value = roleVMDTO.IsArchive == true ? "Да" : "";

                    if (roleVMDTO.UserToRoleDTOs != null)
                    {
                        foreach (UserToRoleDTO userToRoleDTO in roleVMDTO.UserToRoleDTOs)
                        {
                            wsUserToRoleColNum = 1;

                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.Id.ToString();
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.RoleDTOFK.Id.ToString();
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.RoleDTOFK.Name;
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.RoleDTOFK.Description == null ? "" : userToRoleDTO.RoleDTOFK.Description;
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.RoleDTOFK.IsArchive == true ? "Да" : "";

                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.Id.ToString();
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.Login;
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.UserName;
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.Description == null ? "" : userToRoleDTO.UserDTOFK.Description;
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.IsSyncWithAD == true ? "Да" : "";
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.IsServiceUser == true ? "Да" : "";
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.SyncWithADGroupsLastTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                            wsUserToRoleColNum++;
                            wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Value = userToRoleDTO.UserDTOFK.IsArchive == true ? "Да" : "";

                            wsUserToRoleRowNum++;
                        }
                    }

                    if (roleVMDTO.ReportTemplateTypeTоRoleDTOs != null)
                    {
                        foreach (ReportTemplateTypeTоRoleDTO reportTemplateTypeTоRoleDTO in roleVMDTO.ReportTemplateTypeTоRoleDTOs)
                        {
                            wsReportTemplateTypeTоRoleColNum = 1;

                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.Id;
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.RoleDTOFK.Id.ToString();
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.RoleDTOFK.Name;
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.RoleDTOFK.Description == null ? "" : reportTemplateTypeTоRoleDTO.RoleDTOFK.Description;
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.RoleDTOFK.IsArchive == true ? "Да" : "";

                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.ReportTemplateTypeDTOFK.Id.ToString();
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.ReportTemplateTypeDTOFK.Name;
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.ReportTemplateTypeDTOFK.NeedAutoCalc == true ? "Да" : "";
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.CanDownload == true ? "Да" : "";
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.CanUpload == true ? "Да" : "";
                            wsReportTemplateTypeTоRoleColNum++;
                            wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Value = reportTemplateTypeTоRoleDTO.ReportTemplateTypeDTOFK.IsArchive == true ? "Да" : "";

                            wsReportTemplateTypeTоRoleRowNum++;
                        }
                    }

                    if (roleVMDTO.RoleToADGroupDTOs != null)
                    {
                        foreach (RoleToADGroupDTO roleToADGroupDTO in roleVMDTO.RoleToADGroupDTOs)
                        {
                            wsRoleToADGroupColNum = 1;

                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.Id.ToString();
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.RoleDTOFK.Id.ToString();
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.RoleDTOFK.Name;
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.RoleDTOFK.Description == null ? "" : roleToADGroupDTO.RoleDTOFK.Description;
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.RoleDTOFK.IsArchive == true ? "Да" : "";

                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.ADGroupDTOFK.Id.ToString();
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.ADGroupDTOFK.Name;
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.ADGroupDTOFK.Description == null ? "" : roleToADGroupDTO.ADGroupDTOFK.Description;
                            wsRoleToADGroupColNum++;
                            wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Value = roleToADGroupDTO.ADGroupDTOFK.IsArchive == true ? "Да" : "";

                            wsRoleToADGroupRowNum++;
                        }
                    }


                    if (roleVMDTO.RoleToDepartmentDTOs != null)
                    {
                        foreach (RoleToDepartmentDTO roleToDepartmentDTO in roleVMDTO.RoleToDepartmentDTOs)
                        {

                            wsRoleToDepartmentColNum = 1;

                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.Id.ToString();
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.RoleDTOFK.Id.ToString();
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.RoleDTOFK.Name;
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.RoleDTOFK.Description == null ? "" : roleToDepartmentDTO.RoleDTOFK.Description;
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.RoleDTOFK.IsArchive == true ? "Да" : "";

                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.DepartmentDTOFK.Id.ToString();
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.DepartmentDTOFK.MesCode == null ? "" : roleToDepartmentDTO.DepartmentDTOFK.MesCode;
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.DepartmentDTOFK.Name;
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.DepartmentDTOFK.ShortName;
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.DepartmentDTOFK.ToStringHierarchyShortName;
                            wsRoleToDepartmentColNum++;
                            wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Value = roleToDepartmentDTO.DepartmentDTOFK.IsArchive == true ? "Да" : "";

                            wsRoleToDepartmentRowNum++;
                        }
                    }

                    wsRoleRowNum++;
                }

                for (var j = 1; j <= wsRoleColNum; j++)
                    wsRole.Column(j).AdjustToContents();

                for (var j = 1; j <= wsUserToRoleColNum; j++)
                    wsUserToRole.Column(j).AdjustToContents();

                for (var j = 1; j <= wsReportTemplateTypeTоRoleColNum; j++)
                    wsReportTemplateTypeTоRole.Column(j).AdjustToContents();

                for (var j = 1; j <= wsRoleToADGroupColNum; j++)
                    wsRoleToADGroup.Column(j).AdjustToContents();

                for (var j = 1; j <= wsRoleToADGroupColNum; j++)
                    wsRoleToDepartment.Column(j).AdjustToContents();

                var rangeRole = wsRole.Range(wsRole.Cell(1, 1).Address, wsRole.Cell(wsRoleRowNum, wsRoleColNum).Address);
                rangeRole.SetAutoFilter();

                var rangeUserToRole = wsUserToRole.Range(wsUserToRole.Cell(1, 1).Address, wsUserToRole.Cell(wsUserToRoleRowNum, wsUserToRoleColNum).Address);
                rangeUserToRole.SetAutoFilter();

                var rangeReportTemplateTypeTоRole = wsReportTemplateTypeTоRole.Range(wsReportTemplateTypeTоRole.Cell(1, 1).Address,
                    wsReportTemplateTypeTоRole.Cell(wsReportTemplateTypeTоRoleRowNum, wsReportTemplateTypeTоRoleColNum).Address);
                rangeReportTemplateTypeTоRole.SetAutoFilter();

                var rangeRoleToADGroup = wsRoleToADGroup.Range(wsRoleToADGroup.Cell(1, 1).Address,
                    wsRoleToADGroup.Cell(wsRoleToADGroupRowNum, wsRoleToADGroupColNum).Address);
                rangeRoleToADGroup.SetAutoFilter();

                var rangeRoleToDepartment = wsRoleToDepartment.Range(wsRoleToDepartment.Cell(1, 1).Address,
                    wsRoleToDepartment.Cell(wsRoleToDepartmentRowNum, wsRoleToDepartmentColNum).Address);
                rangeRoleToDepartment.SetAutoFilter();


                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<Tuple<IXLWorksheet, int>> AddAllMesDepartmentToExcel(IEnumerable<MesDepartmentVMDTO>? topLevelList, IXLWorksheet? ws, int excelRowNum, int maxLevel)
        {

            if (topLevelList != null)
            {
                foreach (var topLevelItem in topLevelList)
                {
                    MesDepartmentVMDTO? parentDepartmentVMDTO = topLevelItem;

                    int breakCount = 0;
                    while (parentDepartmentVMDTO != null && breakCount <= 100)
                    {
                        //if ((parentDepartmentVMDTO.DepLevel <= 0 || parentDepartmentVMDTO.DepLevel >= 16384))
                        //{ 
                        //int a = 3;
                        //}
                        //else
                        ws.Cell(excelRowNum, parentDepartmentVMDTO.DepLevel).Value = parentDepartmentVMDTO.ShortName;
                        parentDepartmentVMDTO = parentDepartmentVMDTO.DepartmentParentVMDTO;
                        breakCount++;
                    }


                    int j = maxLevel + 1;

                    ws.Cell(excelRowNum, j).Value = topLevelItem.DepLevel.ToString();
                    j++;
                    ws.Cell(excelRowNum, j).Value = topLevelItem.Id.ToString();
                    j++;
                    ws.Cell(excelRowNum, j).Value = topLevelItem.MesCode == null ? "" : topLevelItem.MesCode;
                    j++;
                    ws.Cell(excelRowNum, j).Value = topLevelItem.Name == null ? "" : topLevelItem.Name;
                    j++;
                    ws.Cell(excelRowNum, j).Value = topLevelItem.ShortName == null ? "" : topLevelItem.ShortName;
                    j++;
                    ws.Cell(excelRowNum, j).Value = topLevelItem.IsArchive == true ? "Да" : "";

                    excelRowNum++;

                    Tuple<IXLWorksheet, int> tmp = await AddAllMesDepartmentToExcel(topLevelItem.ChildrenDTO, ws, excelRowNum, maxLevel);
                    ws = tmp.Item1;
                    excelRowNum = tmp.Item2;
                }
            }
            return new Tuple<IXLWorksheet, int>(ws, excelRowNum);

        }

        public async Task<string> GenerateExcelMesDepartments(string filename, IEnumerable<MesDepartmentVMDTO>? mesDepartmentVMDTOList, int maxLevel)
        {
            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesDepartment");

                int excelRowNum = 1;
                int excelColNum = 1;

                for (int j = 1; j <= maxLevel; j++)
                {
                    ws.Cell(excelRowNum, j).Value = "Производство - Уровень " + j.ToString();
                }

                excelColNum = maxLevel + 1;

                ws.Cell(excelRowNum, excelColNum).Value = "Уровень";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код пр-ва (Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сокр. наименование (ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                excelRowNum++;

                Tuple<IXLWorksheet, int> tmp = await AddAllMesDepartmentToExcel(mesDepartmentVMDTOList, ws, excelRowNum, maxLevel);
                ws = tmp.Item1;

                for (var jjj = 1; jjj <= excelColNum; jjj++)
                    ws.Column(jjj).AdjustToContents();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelMesNdoStocks(string filename, IEnumerable<MesNdoStocksDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesNdoStocks");

                int excelRowNum = 1;
                int excelColNum = 1;


                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ид тэга СИР (MesParamId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код тэга СИР (MesParam.Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование тэга СИР (MesParam.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления записи (AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ид пользователя (AddUserId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Логин пользователя (User.Login)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ФИО пользователя (User.UserName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения (ValueTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение (Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Разность (ValueDifference)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД экземпляра отчёта (ReportGuid)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи в витрине SAP (SapNdoOutId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления в витрину (SapNdoOUT.AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Имя тэга в витрине (SapNdoOUT.TagName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения в витрине (SapNdoOUT.ValueTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение в витрине (SapNdoOUT.Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Sap забрал значение (SapNdoOUT.SapGone)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время Sap забрал значение (SapNdoOUT.SapGoneTime)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (MesNdoStocksDTO mesNdoStocksDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.MesParamId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.MesParamDTOFK.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.MesParamDTOFK.Name == null ? "" : mesNdoStocksDTO.MesParamDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.AddUserId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.AddUserDTOFK.Login;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.AddUserDTOFK.UserName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.ValueTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.ValueDifference == null ? "" : mesNdoStocksDTO.ValueDifference.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.ReportGuid == null ? "" : mesNdoStocksDTO.ReportGuid.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOutId == null ? "" : mesNdoStocksDTO.SapNdoOutId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOUTDTOFK == null ? "" : mesNdoStocksDTO.SapNdoOUTDTOFK.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOUTDTOFK == null ? "" : mesNdoStocksDTO.SapNdoOUTDTOFK.TagName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOUTDTOFK == null ? "" : mesNdoStocksDTO.SapNdoOUTDTOFK.ValueTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOUTDTOFK == null ? "" : mesNdoStocksDTO.SapNdoOUTDTOFK.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOUTDTOFK == null ? "" : (mesNdoStocksDTO.SapNdoOUTDTOFK.SapGone == true ? "Да" : "");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesNdoStocksDTO.SapNdoOUTDTOFK == null ? "" : (mesNdoStocksDTO.SapNdoOUTDTOFK.SapGoneTime == null ? "" : ((DateTime)mesNdoStocksDTO.SapNdoOUTDTOFK.SapGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff"));

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelSapNdoOUT(string filename, IEnumerable<SapNdoOUTDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("SapNdoOUT");

                int excelRowNum = 1;
                int excelColNum = 1;


                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Имя тэга (TagName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления записи (AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения (ValueTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение (Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Sap забрал значение (SapGone)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время Sap забрал значение (SapGoneTime)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SapNdoOUTDTO sapNdoOUTDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.TagName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.ValueTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.SapGone == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapNdoOUTDTO.SapGoneTime == null ? "" : ((DateTime)sapNdoOUTDTO.SapGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelMesMovements(string filename, IEnumerable<MesMovementsDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesMovements");

                int excelRowNum = 1;
                int excelColNum = 1;


                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления записи (AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД пользователя (AddUserId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Имя пользователя (User.UserName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД тэга СИР (MesParamId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код тэга СИР (MesParam.Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Имя тэга СИР (MesParam.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Источник (DataSource.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Тип (DataType.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД экземпляра отчёта (ReportGuid)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи в витрине SAP (выход) (SapMovementOutId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ушло в SAP (выход) (SapMovementsOUT.SapGoneTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ошибка SAP (выход) (SapMovementsOUT.SapErrorMessage)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи в витрине SAP (вход) (SapMovementINId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Пришло в витрину SAP (вход) (SapMovementsIN.AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения в витрине SAP (вход) (SapMovementsIN.SapDocumentPostTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ушло из СИР в MES (MesGoneTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД пред. записи (PreviousRecordId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Передавать в Sap (NeedWriteToSap)";


                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Причина корректировки (CorrectionReason.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Комментарий корректировки (MesMovementsComment.CorrectionComment)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (MesMovementsDTO mesMovementsDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.AddUserId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.AddUserDTOFK.UserName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.MesParamId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.MesParamDTOFK.Code == null ? "" : mesMovementsDTO.MesParamDTOFK.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.MesParamDTOFK.Name == null ? "" : mesMovementsDTO.MesParamDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.DataSourceDTOFK == null ? "" : mesMovementsDTO.DataSourceDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.DataTypeDTOFK == null ? "" : mesMovementsDTO.DataTypeDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.ReportGuid == null ? "" : mesMovementsDTO.ReportGuid.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.SapMovementOutId == null ? "" : mesMovementsDTO.SapMovementOutId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.SapMovementsOUTDTOFK == null ? "" : mesMovementsDTO.SapMovementsOUTDTOFK.SapGoneTime == null ? "" : mesMovementsDTO.SapMovementsOUTDTOFK.SapGoneTime.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.SapMovementsOUTDTOFK == null ? "" : mesMovementsDTO.SapMovementsOUTDTOFK.SapErrorMessage == null ? "" : mesMovementsDTO.SapMovementsOUTDTOFK.SapErrorMessage;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.SapMovementInId == null ? "" : mesMovementsDTO.SapMovementInId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.SapMovementsINDTOFK == null ? "" : mesMovementsDTO.SapMovementsINDTOFK.AddTime == null ? "" : mesMovementsDTO.SapMovementsINDTOFK.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.SapMovementsINDTOFK == null ? "" : mesMovementsDTO.SapMovementsINDTOFK.SapDocumentPostTime == null ? "" : mesMovementsDTO.SapMovementsINDTOFK.SapDocumentPostTime.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.MesGoneTime == null ? "" : ((DateTime)mesMovementsDTO.MesGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.PreviousRecordId == null ? "" : mesMovementsDTO.PreviousRecordId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMovementsDTO.NeedWriteToSap == true ? "Да" : "Нет";


                    string correctionNames = "";
                    string correctionComments = "";

                    foreach (var correctionItem in mesMovementsDTO.MesMovementsCommentListDTO)
                    {
                        if (correctionItem.CorrectionReasonDTOFK != null)
                            correctionNames = correctionNames + correctionItem.CorrectionReasonDTOFK.Name + "\n";
                        if (!correctionItem.CorrectionComment.IsNullOrEmpty())
                            correctionComments = correctionComments + correctionItem.CorrectionComment + "\n";
                    }
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = correctionNames;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = correctionComments;

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelSapMovementsOUT(string filename, IEnumerable<SapMovementsOUTDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("SapMovementsOUT");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Партия (BatchNo)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления записи (AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код материала SAP (SapMaterialCode)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. материала SAP (SapMaterial.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код тэга СИР (MesParam.Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. тэга СИР (MesParam.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. тэга СИР (MesParam.Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода источника SAP (ErpPlantIdSource)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код ресурса источника SAP (ErpIdSource)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. источника SAP (SapEquipment.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является складом (IsWarehouseSource)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода приёмника SAP (ErpPlantIdDest)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код ресурса приёмника SAP (ErpIdDest)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. приёмника SAP (SapEquipment.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является складом (IsWarehouseDest)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения (ValueTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение (Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Корректировка (Correction2Previous)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Согласованные (IsReconciled)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ед.изм. SAP (SapUnitOfMeasure)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ушло в SAP (SapGone)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Когда ушло в SAP (SapGoneTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ошибка SAP (SapError)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сообщение ошибки SAP (SapErrorMessage)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи в архиве данных (MesMovementId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения в архиве данных (MesMovements.ValueTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение в архиве данных (MesMovements.Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД пред. записи (PreviousRecordId)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SapMovementsOUTDTO sapMovementsOUTDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.BatchNo == null ? "" : sapMovementsOUTDTO.BatchNo;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapMaterialCode;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapMaterialDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.MesParamDTOFK.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.MesParamDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.MesParamDTOFK.Description == null ? "" : sapMovementsOUTDTO.MesParamDTOFK.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.ErpPlantIdSource == null ? "" : sapMovementsOUTDTO.ErpPlantIdSource;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.ErpIdSource == null ? "" : sapMovementsOUTDTO.ErpIdSource;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapEquipmentSourceDTOFK == null ? "" : sapMovementsOUTDTO.SapEquipmentSourceDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.IsWarehouseSource == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.ErpPlantIdDest == null ? "" : sapMovementsOUTDTO.ErpPlantIdDest;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.ErpIdDest == null ? "" : sapMovementsOUTDTO.ErpIdDest;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapEquipmentDestDTOFK == null ? "" : sapMovementsOUTDTO.SapEquipmentDestDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.IsWarehouseDest == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.ValueTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.Correction2Previous.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.IsReconciled == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapUnitOfMeasure == null ? "" : sapMovementsOUTDTO.SapUnitOfMeasure;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapGone == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapGoneTime == null ? "" : ((DateTime)sapMovementsOUTDTO.SapGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapError == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.SapErrorMessage == null ? "" : sapMovementsOUTDTO.SapErrorMessage;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.MesMovementId == null ? "" : sapMovementsOUTDTO.MesMovementId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.MesMovementsDTOFK == null ? "" : sapMovementsOUTDTO.MesMovementsDTOFK.ValueTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.MesMovementsDTOFK == null ? "" : sapMovementsOUTDTO.MesMovementsDTOFK.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTDTO.PreviousRecordId == null ? "" : sapMovementsOUTDTO.PreviousRecordId.ToString();
                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelSapMovementsIN(string filename, IEnumerable<SapMovementsINDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("SapMovementsIN");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (ErpId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления записи (AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления записи (BatchNo)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время ввода док-та в SAP (SapDocumentEnterTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код материала SAP (SapMaterialCode)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД материала SAP в СИР (SapMaterial.Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. материала SAP в СИР (SapMaterial.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода источника SAP (ErpPlantIdSource)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код ресурса источника SAP (ErpIdSource)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД источника SAP в СИР (SapEquipment.Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. источника SAP в СИР (SapEquipment.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является складом (IsWarehouseSource)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код завода приёмника SAP (ErpPlantIdDest)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код ресурса приёмника SAP (ErpIdDest)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД приёмника SAP в СИР (SapEquipment.Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. приёмника SAP в СИР (SapEquipment.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является складом (IsWarehouseDest)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ед.изм. SAP (SapUnitOfMeasure)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения (SapDocumentPostTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение (Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сторно (IsStorno)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ушло в СИР (MesGone)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Когда ушло в СИР (MesGoneTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ошибка СИР (MesError)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сообщение ошибки СИР (MesErrorMessage)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Тип движения (MoveType)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи в архиве данных (MesMovementId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД тэга СИР (MesParam.Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код тэга СИР (MesParam.Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наим. тэга СИР (MesParam.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время значения в архиве данных (MesMovements.ValueTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение в архиве данных (MesMovements.Value)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД пред. записи (PreviousErpId)";

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SapMovementsINDTO sapMovementsINDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.ErpId;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.BatchNo == null ? "" : sapMovementsINDTO.BatchNo;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapDocumentEnterTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapMaterialCode;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapMaterialDTOFK == null ? "" : sapMovementsINDTO.SapMaterialDTOFK.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapMaterialDTOFK == null ? "" : sapMovementsINDTO.SapMaterialDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.ErpPlantIdSource;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.ErpIdSource;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapEquipmentSourceDTOFK == null ? "" : sapMovementsINDTO.SapEquipmentSourceDTOFK.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapEquipmentSourceDTOFK == null ? "" : sapMovementsINDTO.SapEquipmentSourceDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.IsWarehouseSource == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.ErpPlantIdDest;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.ErpIdDest;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapEquipmentDestDTOFK == null ? "" : sapMovementsINDTO.SapEquipmentDestDTOFK.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapEquipmentDestDTOFK == null ? "" : sapMovementsINDTO.SapEquipmentDestDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.IsWarehouseDest == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapUnitOfMeasure;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.SapDocumentPostTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.IsStorno == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesGone == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesGoneTime == null ? "" : ((DateTime)sapMovementsINDTO.MesGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesError == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesErrorMessage == null ? "" : sapMovementsINDTO.MesErrorMessage;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MoveType == null ? "" : sapMovementsINDTO.MoveType;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesMovementId == null ? "" : sapMovementsINDTO.MesMovementId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesMovementDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.MesParamDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.MesParamDTOFK.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesMovementDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.MesParamDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.MesParamDTOFK.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesMovementDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.MesParamDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.MesParamDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesMovementDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.ValueTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.MesMovementDTOFK == null ? "" : sapMovementsINDTO.MesMovementDTOFK.Value.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapMovementsINDTO.PreviousErpId == null ? "" : sapMovementsINDTO.PreviousErpId;

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelLogEvent(string filename, IEnumerable<LogEventDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("LogEvent");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Тип (LogEventTypeId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время (EventTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Пользователь (UserId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Стар.значение (OldValue)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Нов.значение (NewValue)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД экз. отчёта (ReportEntityId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Крит. (IsCritical)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Ошибка (IsError)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Предупреждение (IsWarning)";
                excelColNum++;


                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                bool firstIteration = true;

                excelRowNum = 2;
                foreach (LogEventDTO logEventDTO in data)
                {
                    excelColNum = 1;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.LogEventTypeDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.EventTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.UserDTOFK.UserNameAndLogin;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.OldValue == null ? "" : logEventDTO.OldValue.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.NewValue == null ? "" : logEventDTO.NewValue.ToString();
                    excelColNum++;
                    if (firstIteration)
                    {
                        ws.Column(excelColNum).Style.Alignment.WrapText = true;
                        firstIteration = false;
                        ws.Column(excelColNum).Width = 80;
                    }

                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.IsCritical == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.IsError == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventDTO.IsWarning == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelMesMaterial(string filename, IEnumerable<MesMaterialDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesMaterial");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Код (Code)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сокр. наименование (ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (MesMaterialDTO mesMaterialDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = mesMaterialDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMaterialDTO.Code;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMaterialDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMaterialDTO.ShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesMaterialDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelSettings(string filename, IEnumerable<SettingsDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("Settings");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Значение (Value)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SettingsDTO settingsDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = settingsDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = settingsDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = settingsDTO.Description == null ? "" : settingsDTO.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = settingsDTO.Value;

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelSapUnitOfMeasure(string filename, IEnumerable<SapUnitOfMeasureDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("SapUnitOfMeasure");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сокр.наименование (ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SapUnitOfMeasureDTO sapUnitOfMeasureDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTO.ShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelMesUnitOfMeasure(string filename, IEnumerable<MesUnitOfMeasureDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesUnitOfMeasure");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Сокр.наименование (ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (MesUnitOfMeasureDTO mesUnitOfMeasureDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = mesUnitOfMeasureDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesUnitOfMeasureDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesUnitOfMeasureDTO.ShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesUnitOfMeasureDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelCorrectionReason(string filename, IEnumerable<CorrectionReasonDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("CorrectionReason");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (CorrectionReasonDTO correctionReasonDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = correctionReasonDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = correctionReasonDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = correctionReasonDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelMesParamSourceType(string filename, IEnumerable<MesParamSourceTypeDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("MesParamSourceType");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (MesParamSourceTypeDTO mesParamSourceTypeDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = mesParamSourceTypeDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamSourceTypeDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = mesParamSourceTypeDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelDataType(string filename, IEnumerable<DataTypeDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("DataType");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Приоритет (выше значение - выше приоритет) (Priority)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Является типом результирующих данных авторасчётов (IsAutoCalcDestDataType)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (DataTypeDTO dataTypeDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = dataTypeDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = dataTypeDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = dataTypeDTO.Priority == null ? "" : dataTypeDTO.Priority.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = dataTypeDTO.IsAutoCalcDestDataType == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = dataTypeDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelDataSource(string filename, IEnumerable<DataSourceDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("DataSource");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (DataSourceDTO dataSourceDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = dataSourceDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = dataSourceDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = dataSourceDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelReportTemplateType(string filename, IEnumerable<ReportTemplateTypeDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("ReportTemplateType");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Требует авторасчёта (NeedAutoCalc)";
                excelColNum++;


                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (ReportTemplateTypeDTO reportTemplateTypeDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateTypeDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateTypeDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateTypeDTO.IsArchive == true ? "Да" : "";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateTypeDTO.NeedAutoCalc == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelReportTemplate(string filename, IEnumerable<ReportTemplateDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("ReportTemplate");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время добавления (AddTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД кем добавлен (AddUserId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Кем добавлен (User.UserName + User.Login)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Описание (Description)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД типа шаблона (ReportTemplateTypeId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Тип шаблона (ReportTemplateType.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД результирующего типа данных (DestDataTypeId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Результирующий тип данных (DataType.Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД производства (DepartmentId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Производство (MesDepartment.ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Имя файла (TemplateFileName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (ReportTemplateDTO reportTemplateDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.AddUserId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.AddUserDTOFK == null ? "" : reportTemplateDTO.AddUserDTOFK.UserName + " (" + reportTemplateDTO.AddUserDTOFK.Login + ")";
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.Description;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.ReportTemplateTypeId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.ReportTemplateTypeDTOFK == null ? "" : reportTemplateDTO.ReportTemplateTypeDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.DestDataTypeId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.DestDataTypeDTOFK == null ? "" : reportTemplateDTO.DestDataTypeDTOFK.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.DepartmentId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.MesDepartmentDTOFK == null ? "" : reportTemplateDTO.MesDepartmentDTOFK.ToStringHierarchyShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.TemplateFileName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = reportTemplateDTO.IsArchive == true ? "Да" : "";
                    excelColNum++;
                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelLogEventType(string filename, IEnumerable<LogEventTypeDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("LogEventType");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (LogEventTypeDTO logEventTypeDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = logEventTypeDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventTypeDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = logEventTypeDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }


        public async Task<string> GenerateExcelSmena(string filename, IEnumerable<SmenaDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("Smena");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Наименование (Name)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "ИД производства (DepartmentId)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Производство (Department.ShortName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время начала (StartTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Продолжительность в часах (HoursDuration)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "В архиве (IsArchive)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SmenaDTO smenaDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = smenaDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = smenaDTO.Name;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = smenaDTO.DepartmentId.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = smenaDTO.DepartmentDTOFK == null ? "" : smenaDTO.DepartmentDTOFK.ToStringHierarchyShortName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = (DateTime.MinValue + smenaDTO.StartTime).ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = smenaDTO.HoursDuration.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = smenaDTO.IsArchive == true ? "Да" : "";

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<string> GenerateExcelScheduler(string filename, IEnumerable<SchedulerDTO> data)
        {

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string fullfilepath = System.IO.Path.Combine(pathVar, filename);

            using var wbook = new XLWorkbook();
            {

                var ws = wbook.AddWorksheet("Scheduler");

                int excelRowNum = 1;
                int excelColNum = 1;

                ws.Cell(excelRowNum, excelColNum).Value = "ИД записи (Id)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Модуль (ModuleName)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время начала (StartTime)";
                excelColNum++;
                ws.Cell(excelRowNum, excelColNum).Value = "Время последнего выполнения (LastExecuted)";
                excelColNum++;

                ws.Row(excelRowNum).Style.Font.SetBold(true);
                ws.Row(excelRowNum).Style.Fill.BackgroundColor = XLColor.LightCyan;

                ws.SheetView.FreezeRows(excelRowNum);

                excelRowNum = 2;
                foreach (SchedulerDTO schedulerDTO in data)
                {
                    excelColNum = 1;

                    ws.Cell(excelRowNum, excelColNum).Value = schedulerDTO.Id.ToString();
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = schedulerDTO.ModuleName;
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = (DateTime.MinValue + schedulerDTO.StartTime).ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                    excelColNum++;
                    ws.Cell(excelRowNum, excelColNum).Value = schedulerDTO.LastExecuted == null ? "" : ((DateTime)schedulerDTO.LastExecuted).ToString("dd.MM.yyyy HH:mm:ss.fff");

                    excelRowNum++;
                }

                for (var j = 1; j <= excelColNum; j++)
                    ws.Column(j).AdjustToContents();

                var range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(excelRowNum, excelColNum).Address);
                range.SetAutoFilter();

                wbook.SaveAs(fullfilepath);
                if (wbook != null)
                    wbook.Dispose();
            }
            return fullfilepath;
        }

        public async Task<Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook?>> GetSheetData(ReportEntityDTO? reportEntityDTO, string sheetSettingName, XLWorkbook? workbook)
        {
            if (reportEntityDTO == null)
            {
                return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Пустой объект экземпляра отчёта", workbook);
            }

            if (workbook == null)
            {
                string pathVar = "";
                if (reportEntityDTO.UploadSuccessFlag)
                    pathVar = (await _settingsRepository.GetByName(SD.ReportUploadPathSettingName)).Value;
                else
                    pathVar = (await _settingsRepository.GetByName(SD.ReportDownloadPathSettingName)).Value;

                string fileName = "";
                if (String.IsNullOrEmpty(reportEntityDTO.DownloadReportFileName))
                {
                    fileName = reportEntityDTO.Id.ToString() + ".xlsx";
                }
                else
                {
                    fileName = reportEntityDTO.DownloadReportFileName;
                }
                string file = System.IO.Path.Combine(pathVar, fileName);
                var extension = Path.GetExtension(fileName);
                if (System.IO.File.Exists(file))
                {
                    try
                    {
                        workbook = new XLWorkbook(file);
                    }
                    catch (Exception ex1)
                    {
                        return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Не удалось загрузить файл: " + file
                            + " Возможно не является файлом формата xlsx", workbook);
                    }
                }
                else
                {
                    return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Не удалось найти файл " + file, workbook);
                }

            }

            if (workbook != null)
            {
                workbook.CalculateMode = XLCalculateMode.Manual;

                if (workbook.IsProtected)
                {
                    try
                    {
                        workbook.Unprotect(SD.ExcelWorkBookProtectionPassword);
                    }
                    catch (Exception exx1)
                    {
                        return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Не удалось снять защиту с книги с помощью пароля: " + SD.ExcelWorkBookProtectionPassword, workbook);
                    }
                }
                string? sheetName = (await _settingsRepository.GetByName(sheetSettingName)).Value;

                if (sheetName.IsNullOrEmpty())
                {
                    return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(),
                            "Не удалось найти настройку в таблице Settings или у неё пустое значение в Value: " + SD.ReportTagLibrarySheetSettingName, workbook);
                }

                IXLWorksheet worksheet = null;
                try
                {
                    worksheet = workbook.Worksheet(sheetName);
                }
                catch (Exception ex2)
                {
                    return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Не удалось загрузить лист: " + sheetName, workbook);
                }

                if (worksheet == null)
                {
                    return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Не найден лист: " + sheetName, workbook);
                }
                ExcelSheetWithSirTagsDTOList reportList = new ExcelSheetWithSirTagsDTOList();

                var headerRows = worksheet.Row(1);

                reportList.Column1Name = headerRows.Cell(1).CachedValue.ToString();
                reportList.Column2Name = headerRows.Cell(2).CachedValue.ToString();
                reportList.Column3Name = headerRows.Cell(3).CachedValue.ToString();
                reportList.Column4Name = headerRows.Cell(4).CachedValue.ToString();
                reportList.Column5Name = headerRows.Cell(5).CachedValue.ToString();
                reportList.Column6Name = headerRows.Cell(6).CachedValue.ToString();
                reportList.Column7Name = headerRows.Cell(7).CachedValue.ToString();
                reportList.Column8Name = headerRows.Cell(8).CachedValue.ToString();
                reportList.Column9Name = headerRows.Cell(9).CachedValue.ToString();
                reportList.Column10Name = headerRows.Cell(10).CachedValue.ToString();
                reportList.Column11Name = headerRows.Cell(11).CachedValue.ToString();
                reportList.Column12Name = headerRows.Cell(12).CachedValue.ToString();

                IEnumerable<IXLRangeRow>? rows = null;

                try
                {
                    rows = worksheet.RangeUsed().RowsUsed().Skip(1);
                }
                catch (Exception ex2)
                {
                    return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "Не удалось получить строки листа: " + sheetName, workbook);
                }


                //bool notImplementedThirdColumnForEmbReport = (reportEntityDTO.ReportTemplateDTOFK.ReportTemplateTypeDTOFK.Name.Trim().ToUpper() == "ОТЧЁТ ЭМБ"
                //    && sheetSettingName == SD.ReportOutputSheetSettingName);

                bool notImplementedThirdColumnForEmbReport = false;

                foreach (var row in rows)
                {
                    ExcelSheetWithSirTagsDTO rowItem = new ExcelSheetWithSirTagsDTO();

                    rowItem.Column1 = await GetCellValue(row.Cell(1));
                    rowItem.Column2 = await GetCellValue(row.Cell(2));
                    if (notImplementedThirdColumnForEmbReport)
                        rowItem.Column3 = "Формула не поддерживается";
                    else
                        rowItem.Column3 = await GetCellValue(row.Cell(3));
                    rowItem.Column4 = await GetCellValue(row.Cell(4));
                    rowItem.Column5 = await GetCellValue(row.Cell(5));
                    rowItem.Column6 = await GetCellValue(row.Cell(6));
                    rowItem.Column7 = await GetCellValue(row.Cell(7));
                    rowItem.Column8 = await GetCellValue(row.Cell(8));
                    rowItem.Column9 = await GetCellValue(row.Cell(9));
                    rowItem.Column10 = await GetCellValue(row.Cell(10));
                    rowItem.Column11 = await GetCellValue(row.Cell(11));
                    rowItem.Column12 = await GetCellValue(row.Cell(12));
                    rowItem.MesParamFoundFlag = false;


                    reportList.excelSheetWithSirTagsDTOList.Add(rowItem);
                }

                IEnumerable<ExcelSheetWithSirTagsDTO>? returnResult;

                if (reportList.Column1Name.Trim().ToUpper() == "MESPARAMCODE")
                {

                    returnResult = (from reportListAlias in reportList.excelSheetWithSirTagsDTOList
                                    join MP_prom1 in _db.MesParam.AsNoTracking().ToListWithNoLock() on
                                         reportListAlias.Column1 equals MP_prom1.Code
                                    into MP_prom2
                                    from MP in MP_prom2.DefaultIfEmpty()
                                    select
                                            new ExcelSheetWithSirTagsDTO
                                            {
                                                Column1 = reportListAlias.Column1,
                                                Column2 = reportListAlias.Column2,
                                                Column3 = reportListAlias.Column3,
                                                Column4 = reportListAlias.Column4,
                                                Column5 = reportListAlias.Column5,
                                                Column6 = reportListAlias.Column6,
                                                Column7 = reportListAlias.Column7,
                                                Column8 = reportListAlias.Column8,
                                                Column9 = reportListAlias.Column9,
                                                Column10 = reportListAlias.Column10,
                                                Column11 = reportListAlias.Column11,
                                                Column12 = reportListAlias.Column12,
                                                MesParamDTOFK = _mapper.Map<MesParam, MesParamDTO>(MP),
                                                MesParamFoundFlag = (MP == null ? true : false),
                                            }).ToList();
                }
                else
                {
                    returnResult = reportList.excelSheetWithSirTagsDTOList;
                }
                reportList.excelSheetWithSirTagsDTOList = (List<ExcelSheetWithSirTagsDTO>)returnResult
                        .Where(u => ((!u.Column1.IsNullOrEmpty()) || (!u.Column2.IsNullOrEmpty())
                         || (!u.Column3.IsNullOrEmpty()) || (!u.Column4.IsNullOrEmpty()) || (!u.Column5.IsNullOrEmpty())
                         || (!u.Column6.IsNullOrEmpty()) || (!u.Column7.IsNullOrEmpty()) || (!u.Column8.IsNullOrEmpty())
                         || (!u.Column9.IsNullOrEmpty()) || (!u.Column10.IsNullOrEmpty()) || (!u.Column11.IsNullOrEmpty())
                          || (!u.Column12.IsNullOrEmpty())
                        )).ToList();

                return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(reportList, "", workbook);
            }
            return new Tuple<ExcelSheetWithSirTagsDTOList, string, XLWorkbook>(new ExcelSheetWithSirTagsDTOList(), "", workbook);
        }

        public async Task<string> GetCellValue(IXLCell? cell)
        {
            string retVar = "";
            if (cell == null)
                return "";
            retVar = cell.CachedValue.ToString();

            if (!string.IsNullOrEmpty(retVar))
                return retVar;

            if (cell.NeedsRecalculation)
            {

                switch (cell.DataType)
                {
                    case XLDataType.Text:
                        {
                            string ggg;
                            if (cell.TryGetValue<string>(out ggg))
                                retVar = ggg;
                            else
                                retVar = "Формула не поддерживается";
                            break;
                        }
                    case XLDataType.Boolean:
                        {
                            bool ggg;
                            if (cell.TryGetValue<bool>(out ggg))
                                retVar = ggg.ToString();
                            else
                                retVar = "Формула не поддерживается";
                            break;
                        }
                    case XLDataType.DateTime:
                        {
                            DateTime ggg;
                            if (cell.TryGetValue<DateTime>(out ggg))
                                retVar = ggg.ToString();
                            else
                                retVar = "Формула не поддерживается";
                            break;
                        }
                    case XLDataType.Number:
                        {
                            Decimal ggg;
                            if (cell.TryGetValue<Decimal>(out ggg))
                                retVar = ggg.ToString();
                            else
                                retVar = "Формула не поддерживается";
                            break;
                        }
                    case XLDataType.Blank:
                        {
                            var ttt = cell.Style.DateFormat.NumberFormatId;
                            if (ttt == 14 || ttt == -1)
                            {
                                DateTime ggg;
                                if (cell.TryGetValue<DateTime>(out ggg))
                                    retVar = ggg.ToString();
                                else
                                    retVar = "Формула не поддерживается";
                            }
                            else
                            {
                                string ggg;
                                if (cell.TryGetValue<string>(out ggg))
                                    retVar = ggg;
                                else
                                    retVar = "Формула не поддерживается";
                            }
                            break;
                        }
                }
            }
            return retVar;
        }
    }
}
