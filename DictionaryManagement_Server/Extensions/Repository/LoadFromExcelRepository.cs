using ClosedXML.Excel;
using DictionaryManagement_Business.Repository;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_Models.IntDBModels;
using DictionaryManagement_Server.Extensions.Repository.IRepository;
using Radzen;

namespace DictionaryManagement_Server.Extensions.Repository
{
    public class LoadFromExcelRepository : ILoadFromExcelRepository
    {
        private readonly ISapMaterialRepository _sapMaterialRepository;
        private readonly IMesMaterialRepository _mesMaterialRepository;
        private readonly ISapEquipmentRepository _sapEquipmentRepository;
        private readonly ILogEventRepository _logEventRepository;
        private readonly IMesParamRepository _mesParamRepository;
        private readonly IMesParamSourceTypeRepository _mesParamSourceTypeRepository;
        private readonly IMesDepartmentRepository _mesDepartmentRepository;
        private readonly ISapUnitOfMeasureRepository _sapUnitOfMeasureRepository;
        private readonly IADGroupRepository _adGroupRepository;
        private readonly IUserRepository _userRepository;

        public LoadFromExcelRepository(ISapMaterialRepository sapMaterialRepository, IMesMaterialRepository mesMaterialRepository,
            ISapEquipmentRepository sapEquipmentRepository,
            ILogEventRepository logEventRepository, IMesParamRepository mesParamRepository,
            IMesParamSourceTypeRepository mesParamSourceTypeRepository,
            IMesDepartmentRepository mesDepartmentRepository,
            ISapUnitOfMeasureRepository sapUnitOfMeasureRepository,
            IADGroupRepository adGroupRepository,
            IUserRepository userRepository)
        {
            _sapMaterialRepository = sapMaterialRepository;
            _mesMaterialRepository = mesMaterialRepository;
            _sapEquipmentRepository = sapEquipmentRepository;
            _logEventRepository = logEventRepository;
            _mesParamRepository = mesParamRepository;
            _mesParamSourceTypeRepository = mesParamSourceTypeRepository;
            _mesDepartmentRepository = mesDepartmentRepository;
            _sapUnitOfMeasureRepository = sapUnitOfMeasureRepository;
            _adGroupRepository = adGroupRepository;
            _userRepository = userRepository;
        }

        public async Task<string> MaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Материалов " + worksheet.Name.Substring(0, 3);
            await loadFromExcelPage.RefreshSate();

            IEnumerable<MaterialDTO> MaterialDTOList;
            if (worksheet.Name.Equals("SapMaterial"))
            {
                MaterialDTOList = (IEnumerable<MaterialDTO>)(await _sapMaterialRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.Id);
            }
            else
                MaterialDTOList = (IEnumerable<MaterialDTO>)(await _sapMaterialRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.Id);

            int recordCount = MaterialDTOList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var materialDTOItem in MaterialDTOList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 2;
                worksheet.Cell(excelRowNum, excelColNum).Value = materialDTOItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = materialDTOItem.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = materialDTOItem.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = materialDTOItem.ShortName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = materialDTOItem.IsArchive == true ? "Да" : "Нет";
                excelRowNum++;
            }
            return worksheet.Name + "_Example_with_data_";
        }


        public async Task<string> SapEquipmentReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Ресурсов SAP)";
            await loadFromExcelPage.RefreshSate();

            IEnumerable<SapEquipmentDTO> sapEquipmentDTOList = (await _sapEquipmentRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.Id);

            int recordCount = sapEquipmentDTOList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var sapEquipmentDTOItem in sapEquipmentDTOList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 2;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTOItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTOItem.ErpPlantId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTOItem.ErpId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTOItem.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTOItem.IsWarehouse == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapEquipmentDTOItem.IsArchive == true ? "Да" : "Нет";
                excelRowNum++;
            }
            return "SapEquipment_Example_with_data_";
        }


        public async Task<string> MesParamReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {

            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Тэгов СИР)";
            await loadFromExcelPage.RefreshSate();

            IEnumerable<MesParamDTO> mesParamDTOList = (await _mesParamRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.Id);

            int recordCount = mesParamDTOList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var mesParamDTOItem in mesParamDTOList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 2;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.Name == null ? "" : mesParamDTOItem.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.Description == null ? "" : mesParamDTOItem.Description;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.MesParamSourceTypeDTOFK == null ? "" : mesParamDTOItem.MesParamSourceTypeDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.MesParamSourceLink == null ? "" : mesParamDTOItem.MesParamSourceLink;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.MesDepartmentDTOFK == null ? "" : mesParamDTOItem.MesDepartmentDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.MesDepartmentDTOFK == null ? "" : mesParamDTOItem.MesDepartmentDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentSourceDTOFK == null ? "" : mesParamDTOItem.SapEquipmentSourceDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentSourceDTOFK == null ? "" : mesParamDTOItem.SapEquipmentSourceDTOFK.ErpPlantId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentSourceDTOFK == null ? "" : mesParamDTOItem.SapEquipmentSourceDTOFK.ErpId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentSourceDTOFK == null ? "" : mesParamDTOItem.SapEquipmentSourceDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentDestDTOFK == null ? "" : mesParamDTOItem.SapEquipmentDestDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentDestDTOFK == null ? "" : mesParamDTOItem.SapEquipmentDestDTOFK.ErpPlantId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentDestDTOFK == null ? "" : mesParamDTOItem.SapEquipmentDestDTOFK.ErpId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapEquipmentDestDTOFK == null ? "" : mesParamDTOItem.SapEquipmentDestDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapMaterialDTOFK == null ? "" : mesParamDTOItem.SapMaterialDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapMaterialDTOFK == null ? "" : mesParamDTOItem.SapMaterialDTOFK.Code.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapMaterialDTOFK == null ? "" : mesParamDTOItem.SapMaterialDTOFK.Name.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.SapUnitOfMeasureDTOFK == null ? "" : mesParamDTOItem.SapUnitOfMeasureDTOFK.Name.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.DaysRequestInPast == null ? "" : mesParamDTOItem.DaysRequestInPast.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.TI == null ? "" : mesParamDTOItem.TI;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.NameTI == null ? "" : mesParamDTOItem.NameTI;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.TM == null ? "" : mesParamDTOItem.TM;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.NameTM == null ? "" : mesParamDTOItem.NameTM;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.MesToSirUnitOfMeasureKoef == null ? "" : mesParamDTOItem.MesToSirUnitOfMeasureKoef.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.NeedWriteToSap == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.NeedReadFromSap == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.NeedReadFromMes == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.NeedWriteToMes == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.IsNdo == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesParamDTOItem.IsArchive == true ? "Да" : "Нет";

                excelRowNum++;
            }

            return "MesParam_Example_with_data_";
        }


        public async Task<string> MesNdoStocksReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MesNdoStocksDTO>? mesNdoStocksList)
        {

            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка записей Архива данных НДО)";
            await loadFromExcelPage.RefreshSate();

            if (mesNdoStocksList == null || mesNdoStocksList.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените интервал дат или фильтры в отображаемом списке записей Архива данных НДО");
                return "MesNdoStocks_Example_with_data_";
            }

            int recordCount = mesNdoStocksList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var mesNdoStocksItem in mesNdoStocksList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.AddUserId.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.AddUserDTOFK.UserName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.MesParamDTOFK.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.ValueTime.ToString("dd.MM.yyyy HH:mm:ss");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.Value.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.ValueDifference.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.ReportGuid == null ? "" : mesNdoStocksItem.ReportGuid.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesNdoStocksItem.SapNdoOutId == null ? "" : mesNdoStocksItem.SapNdoOutId.ToString();
                excelRowNum++;
            }

            return "MesNdoStocks_Example_with_data_";
        }


        public async Task<string> SapNdoOUTReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapNdoOUTDTO>? sapNdoOUTList)
        {

            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка записей витрины SAP НДО-выход)";
            await loadFromExcelPage.RefreshSate();

            if (sapNdoOUTList == null || sapNdoOUTList.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените интервал дат или фильтры в отображаемом списке записей витрины SAP НДО-выход");
                return "SapNdoOUT_Example_with_data_";
            }

            int recordCount = sapNdoOUTList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var sapNdoOUTItem in sapNdoOUTList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.TagName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.ValueTime.ToString("dd.MM.yyyy HH:mm:ss");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.Value.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.SapGone == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapNdoOUTItem.SapGoneTime == null ? "" : ((DateTime)sapNdoOUTItem.SapGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelRowNum++;
            }

            return "SapNdoOUT_Example_with_data_";
        }



        public async Task<string> MesMovementsReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<MesMovementsDTO>? mesMovementsListDTO)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка записей Архива данных)";
            await loadFromExcelPage.RefreshSate();

            if (mesMovementsListDTO == null || mesMovementsListDTO.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените интервал дат или фильтры в отображаемом списке записей Архива данных");
                return "MesMovements_Example_with_data_";
            }

            int recordCount = mesMovementsListDTO.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var mesMovementsItemDTO in mesMovementsListDTO)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.AddUserDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.AddUserDTOFK.UserName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.MesParamDTOFK.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.ValueTime.ToString("dd.MM.yyyy HH:mm:ss");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.Value.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.SapMovementInId == null ? "" : mesMovementsItemDTO.SapMovementInId.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.SapMovementOutId == null ? "" : mesMovementsItemDTO.SapMovementOutId.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.DataSourceDTOFK == null ? "" : mesMovementsItemDTO.DataSourceDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.DataTypeDTOFK == null ? "" : mesMovementsItemDTO.DataTypeDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.ReportGuid == null ? "" : mesMovementsItemDTO.ReportGuid.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.MesGone == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.NeedWriteToSap == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMovementsItemDTO.PreviousRecordId == null ? "" : mesMovementsItemDTO.PreviousRecordId.ToString();

                excelRowNum++;
            }

            return "MesMovements_Example_with_data_";
        }

        public async Task<string> SapMovementsOUTReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapMovementsOUTDTO>? sapMovementsOUTListDTO)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка записей витрины SAP Движения-выход)";
            await loadFromExcelPage.RefreshSate();

            if (sapMovementsOUTListDTO == null || sapMovementsOUTListDTO.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените интервал дат или фильтры в отображаемом списке записей витрины SAP Движения-выход");
                return "SapMovementsOUT_Example_with_data_";
            }

            int recordCount = sapMovementsOUTListDTO.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var sapMovementsOUTItemDTO in sapMovementsOUTListDTO)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.BatchNo == null ? "" : sapMovementsOUTItemDTO.BatchNo;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.MesParamId.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.MesParamDTOFK.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.MesParamDTOFK.Name == null ? "" : sapMovementsOUTItemDTO.MesParamDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapMaterialDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapMaterialCode;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapMaterialDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapEquipmentSourceDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.ErpPlantIdSource;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.ErpIdSource;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapEquipmentSourceDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.IsWarehouseSource == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapEquipmentDestDTOFK.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.ErpPlantIdDest;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.ErpIdDest;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapEquipmentDestDTOFK.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.IsWarehouseDest == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.ValueTime.ToString("dd.MM.yyyy HH:mm:ss");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.Value.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.Correction2Previous.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.IsReconciled == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapUnitOfMeasure == null ? "" : sapMovementsOUTItemDTO.SapUnitOfMeasure;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapUnitOfMeasure == null ? "" : sapMovementsOUTItemDTO.SapUnitOfMeasure;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapGone == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapGoneTime == null ? "" : ((DateTime)sapMovementsOUTItemDTO.SapGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.SapErrorMessage == null ? "" : sapMovementsOUTItemDTO.SapErrorMessage;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.MesMovementId == null ? "" : sapMovementsOUTItemDTO.MesMovementId.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsOUTItemDTO.PreviousRecordId == null ? "" : sapMovementsOUTItemDTO.PreviousRecordId.ToString();

                excelRowNum++;
            }
            return "SapMovementsOUT_Example_with_data_";
        }

        public async Task<string> SapMovementsINReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<SapMovementsINDTO>? sapMovementsINListDTO)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка записей витрины SAP Движения-вход)";
            await loadFromExcelPage.RefreshSate();

            if (sapMovementsINListDTO == null || sapMovementsINListDTO.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените интервал дат или фильтры в отображаемом списке записей витрины SAP Движения-вход");
                return "SapMovementsIN_Example_with_data_";
            }

            int recordCount = sapMovementsINListDTO.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var sapMovementsINItemOUT in sapMovementsINListDTO)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.ErpId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.AddTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.SapDocumentEnterTime.ToString("dd.MM.yyyy HH:mm:ss");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.BatchNo == null ? "" : sapMovementsINItemOUT.BatchNo;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.SapMaterialCode;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.ErpPlantIdSource;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.ErpIdSource;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.IsWarehouseSource == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.ErpPlantIdDest;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.ErpIdDest;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.IsWarehouseDest == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.SapDocumentPostTime.ToString("dd.MM.yyyy HH:mm:ss");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.Value.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.SapUnitOfMeasure;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.IsStorno == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.MesGone == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.MesGoneTime == null ? "" : ((DateTime)sapMovementsINItemOUT.MesGoneTime).ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.MesError == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.MesErrorMessage == null ? "" : sapMovementsINItemOUT.MesErrorMessage;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.MesMovementId == null ? "" : sapMovementsINItemOUT.MesMovementId.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.PreviousErpId == null ? "" : sapMovementsINItemOUT.PreviousErpId;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMovementsINItemOUT.MoveType == null ? "" : sapMovementsINItemOUT.MoveType;

                excelRowNum++;
            }

            return "SapMovementsIN_Example_with_data_";
        }

        public async Task<string> UserReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<UserDTO>? userListDTO)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка отображаемых записей Справочника пользователей)";
            await loadFromExcelPage.RefreshSate();

            if (userListDTO == null || userListDTO.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените фильтры в отображаемом списке Справочника пользователей");
                return "User_Example_with_data_";
            }

            int recordCount = userListDTO.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var userListItemDTO in userListDTO)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.Login;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.UserName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.Description == null ? "" : userListItemDTO.Description;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.IsSyncWithAD == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.SyncWithADGroupsLastTime == null ? "" : userListItemDTO.SyncWithADGroupsLastTime.ToString("dd.MM.yyyy HH:mm:ss.fff");
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.IsServiceUser == true ? "Да" : "Нет";
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = userListItemDTO.IsArchive == true ? "Да" : "Нет";

                excelRowNum++;
            }

            return "User_Example_with_data_";
        }
        public async Task<string> ADGroupReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet, IEnumerable<ADGroupDTO>? adGroupListDTO)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка записей справочника Группы AD)";
            await loadFromExcelPage.RefreshSate();

            if (adGroupListDTO == null || adGroupListDTO.Count() <= 0)
            {
                await loadFromExcelPage.ShowSwal("warning", "Выгрузился пустой список. Измените фильтры в отображаемом списке Справочника групп AD");
                return "ADGroup_Example_with_data_";
            }

            int recordCount = adGroupListDTO.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var adGroupItemDTO in adGroupListDTO)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 3;
                worksheet.Cell(excelRowNum, excelColNum).Value = adGroupItemDTO.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = adGroupItemDTO.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = adGroupItemDTO.Description;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = adGroupItemDTO.IsArchive == true ? "Да" : "Нет";

                excelRowNum++;
            }

            return "ADGroup_Example_with_data_";
        }


        public async Task<bool> MaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;
            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в справочник Материалов " + worksheet.Name.Substring(0, 3));
            await loadFromExcelPage.RefreshSate();

            loadFromExcelPage.haveChanges = false;

            int resultColumnNumber = 7;
            int rowNumber = 9;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                var rowVar = worksheet.Row(rowNumber);

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;


                string idVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string codeVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string nameVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string shortNameVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string isArchiveVarString = rowVar.Cell(6).CachedValue.ToString().Trim();

                string resultString = "";
                int idVarInt = 0;

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(codeVarString) && String.IsNullOrEmpty(nameVarString)
                        && String.IsNullOrEmpty(shortNameVarString) && String.IsNullOrEmpty(isArchiveVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(codeVarString))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3. И ИД записи, и Код материала пустые. Изменения не применялись.";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                    rowNumber++;
                    continue;
                }

                string isUpdateOrAddMode = "NONE";
                bool needCheckCode = false;
                bool needCheckName = false;
                bool needCheckShortName = false;

                MaterialDTO? foundMaterialDTO = null;
                MaterialDTO changedMaterialDTO = new MaterialDTO();
                if (!String.IsNullOrEmpty(idVarString))  // если указан id элемента, то рассматриваем только вариант редактирования уже существующей записи
                {
                    isUpdateOrAddMode = "UPDATE";
                    try
                    {
                        idVarInt = int.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        idVarInt = 0;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не удалось получить целое число ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    if (worksheet.Name.Equals("SapMaterial"))
                        foundMaterialDTO = await _sapMaterialRepository.Get(idVarInt);
                    else
                        foundMaterialDTO = await _mesMaterialRepository.Get(idVarInt);
                    if (foundMaterialDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не найдена запись в справочнике с ИД записи равным " + idVarInt.ToString() + ".Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    changedMaterialDTO.Id = idVarInt;
                    if (String.IsNullOrEmpty(codeVarString))
                    {
                        changedMaterialDTO.Code = foundMaterialDTO.Code;
                    }
                    else
                    {
                        if (foundMaterialDTO.Code.Equals(codeVarString))
                            changedMaterialDTO.Code = foundMaterialDTO.Code;
                        else
                        {
                            needCheckCode = true;
                            changedMaterialDTO.Code = codeVarString;
                        }
                    }

                    if (String.IsNullOrEmpty(nameVarString))
                    {
                        changedMaterialDTO.Name = foundMaterialDTO.Name;
                    }
                    else
                    {
                        if (foundMaterialDTO.Name.Equals(nameVarString))
                            changedMaterialDTO.Name = foundMaterialDTO.Name;
                        else
                        {
                            needCheckName = true;
                            changedMaterialDTO.Name = nameVarString;
                        }
                    }

                    if (String.IsNullOrEmpty(shortNameVarString))
                    {
                        changedMaterialDTO.ShortName = foundMaterialDTO.ShortName;
                    }
                    else
                    {
                        if (foundMaterialDTO.ShortName.Equals(shortNameVarString))
                        {
                            changedMaterialDTO.ShortName = foundMaterialDTO.ShortName;
                        }
                        else
                        {
                            needCheckShortName = true;
                            changedMaterialDTO.ShortName = shortNameVarString;
                        }
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(codeVarString))
                    {
                        if (worksheet.Name.Equals("SapMaterial"))
                            foundMaterialDTO = await _sapMaterialRepository.GetByCode(codeVarString);
                        else
                            foundMaterialDTO = await _mesMaterialRepository.GetByCode(codeVarString);
                    }
                    if (foundMaterialDTO == null)  // добавление
                    {
                        changedMaterialDTO.Code = codeVarString;
                        isUpdateOrAddMode = "ADD";
                        needCheckCode = true;
                        needCheckName = true;
                        needCheckShortName = true;

                        changedMaterialDTO.Code = codeVarString;
                        changedMaterialDTO.Name = nameVarString;
                        changedMaterialDTO.ShortName = shortNameVarString;
                    }
                    else  // редактирование
                    {
                        isUpdateOrAddMode = "UPDATE";
                        needCheckCode = false;
                        needCheckName = true;
                        needCheckShortName = true;

                        changedMaterialDTO.Id = foundMaterialDTO.Id;
                        changedMaterialDTO.Code = foundMaterialDTO.Code;
                        changedMaterialDTO.Name = nameVarString;
                        changedMaterialDTO.ShortName = shortNameVarString;
                    }
                }

                if (String.IsNullOrEmpty(changedMaterialDTO.Name))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4. Наименование материала не может быть пустым. Изменения не применялись.";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 4 }, resultString);
                    rowNumber++;
                    continue;
                }

                if (String.IsNullOrEmpty(changedMaterialDTO.ShortName))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Сокращённое наименование материала не может быть пустым. Изменения не применялись.";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 5 }, resultString);
                    rowNumber++;
                    continue;
                }


                if (needCheckCode)
                {
                    MaterialDTO? objectForCheckCode;
                    if (worksheet.Name.Equals("SapMaterial"))
                        objectForCheckCode = _sapMaterialRepository.GetByCode(codeVarString).Result;
                    else
                        objectForCheckCode = _mesMaterialRepository.GetByCode(codeVarString).Result;

                    if (objectForCheckCode != null)
                    {
                        bool isBad = false;
                        if (isUpdateOrAddMode == "UPDATE")
                        {
                            if (objectForCheckCode.Id != foundMaterialDTO.Id)
                            {
                                isBad = true;
                            }
                        }
                        else  // если добавление
                        {
                            isBad = true;
                        }
                        if (isBad)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 3. Уже есть запись с кодом материала " + codeVarString
                                + ". ИД записи: " + objectForCheckCode.Id.ToString() +
                                ". Наименование: " + objectForCheckCode.ShortName + ". Изменения не применялись."; ;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 3 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }
                }


                if (needCheckName)
                {
                    MaterialDTO? objectForCheckName;
                    if (worksheet.Name.Equals("SapMaterial"))
                        objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                    else
                        objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;

                    if (objectForCheckName != null)
                    {
                        bool isBad = false;
                        if (isUpdateOrAddMode == "UPDATE")
                        {
                            if (objectForCheckName.Id != foundMaterialDTO.Id)
                            {
                                isBad = true;
                            }
                        }
                        else  // это добавление записи
                        {
                            isBad = true;
                        }
                        if (isBad)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 4. Уже есть запись с наименованием " + nameVarString
                                        + ". ИД записи: " + objectForCheckName.Id.ToString() +
                                        ". Код: " + objectForCheckName.Code + ". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 4 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }
                }

                if (needCheckShortName)
                {
                    MaterialDTO? objectForCheckShortName;
                    if (worksheet.Name.Equals("SapMaterial"))
                        objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                    else
                        objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;

                    if (objectForCheckShortName != null)
                    {
                        bool isBad = false;
                        if (isUpdateOrAddMode == "UPDATE")
                        {
                            if (objectForCheckShortName.Id != foundMaterialDTO.Id)
                            {
                                isBad = true;
                            }
                        }
                        else
                        {
                            isBad = true;
                        }
                        if (isBad)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Уже есть запись с сокр. наименованием " + shortNameVarString
                                        + ". ИД записи: " + objectForCheckShortName.Id.ToString() +
                                        ". Код: " + objectForCheckShortName.Code + ". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 5 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }
                }

                switch (isUpdateOrAddMode)
                {
                    case "ADD":
                        {
                            if (String.IsNullOrEmpty(isArchiveVarString))
                                changedMaterialDTO.IsArchive = false;
                            else
                                changedMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                            MaterialDTO? newMaterialDTO;
                            if (worksheet.Name.Equals("SapMaterial"))
                            {
                                newMaterialDTO = await _sapMaterialRepository.Create(new SapMaterialDTO(changedMaterialDTO));
                                await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: null, newObject: new SapMaterialDTO(newMaterialDTO), "Добавление материала SAP", "Материал SAP: ", _authorizationRepository);
                            }
                            else
                            {
                                newMaterialDTO = await _mesMaterialRepository.Create(new MesMaterialDTO(changedMaterialDTO));
                                await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: null, newObject: new MesMaterialDTO(newMaterialDTO), "Добавление материала MES", "Материал MES: ", _authorizationRepository);
                            }

                            loadFromExcelPage.haveChanges = true;
                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Материал добавлен с кодом " + newMaterialDTO.Id.ToString();
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);

                            loadFromExcelPage.console.Log(resultString);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "UPDATE":
                        {
                            if (String.IsNullOrEmpty(isArchiveVarString))
                            {
                                changedMaterialDTO.IsArchive = false;
                            }
                            else
                            {
                                changedMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                            }

                            if (worksheet.Name.Equals("SapMaterial"))
                                await _sapMaterialRepository.Update(new SapMaterialDTO(changedMaterialDTO), SD.UpdateMode.Update);
                            else
                                await _mesMaterialRepository.Update(new MesMaterialDTO(changedMaterialDTO), SD.UpdateMode.Update);

                            if (changedMaterialDTO.IsArchive != foundMaterialDTO.IsArchive)
                            {
                                SD.UpdateMode updMode;
                                if (changedMaterialDTO.IsArchive == true)
                                {
                                    updMode = SD.UpdateMode.MoveToArchive;
                                }
                                else
                                {
                                    updMode = SD.UpdateMode.RestoreFromArchive;
                                }
                                if (worksheet.Name.Equals("SapMaterial"))
                                {
                                    await _sapMaterialRepository.Update(new SapMaterialDTO(changedMaterialDTO), updMode);
                                    await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: new SapMaterialDTO(foundMaterialDTO), newObject: new SapMaterialDTO(changedMaterialDTO), "Изменение материала SAP", "Материал SAP: ", _authorizationRepository);
                                }
                                else
                                {
                                    await _mesMaterialRepository.Update(new MesMaterialDTO(changedMaterialDTO), updMode);
                                    await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: new MesMaterialDTO(foundMaterialDTO), newObject: new MesMaterialDTO(changedMaterialDTO), "Изменение материала MES", "Материал MES: ", _authorizationRepository);
                                }
                            }

                            loadFromExcelPage.haveChanges = true;
                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            resultString = "!!! Для строки " + rowNumber.ToString() + " определен не предусмотренный режим обработки = " + isUpdateOrAddMode + ". Изменения не производились.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 1 }, resultString);
                            rowNumber++;
                            continue;
                        }

                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа " + worksheet.Name + " в справочник Материалов " + worksheet.Name.Substring(0, 3));
            await loadFromExcelPage.RefreshSate();
            return haveErrors;
        }

        public async Task<bool> SapEquipmentExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист SapEquipment загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа SapEquipment в справочник Ресурсов SAP");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string idVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string erpPlantIdVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string erpIdVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string nameVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string isWarehouseVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string isArchiveVarString = rowVar.Cell(7).CachedValue.ToString().Trim();
                string resultString = "";
                int idVarInt = 0;

                int resultColumnNumber = 8;

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(erpPlantIdVarString) && String.IsNullOrEmpty(erpIdVarString)
                        && String.IsNullOrEmpty(nameVarString)
                        && String.IsNullOrEmpty(isWarehouseVarString) && String.IsNullOrEmpty(isArchiveVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                if (String.IsNullOrEmpty(idVarString) && (String.IsNullOrEmpty(erpPlantIdVarString) || String.IsNullOrEmpty(erpIdVarString)))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3, 4. Пустые: ИД записи, и одно из полей \"Код завода SAP\" или . \"Код ресурса/склада SAP\". Изменения не применялись.";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 2, 3, 4 }, resultString);
                    rowNumber++;
                    continue;
                }

                string isUpdateOrAddMode = "NONE";
                bool needCheckErpPlantIdPlusErpId = false;
                bool needCheckName = false;
                string duplicateNameString = "";

                SapEquipmentDTO? foundSapEquipmentDTO = null;
                SapEquipmentDTO changedSapEquipmentDTO = new SapEquipmentDTO();
                if (!String.IsNullOrEmpty(idVarString))  // если указан id элемента, то рассматриваем только вариант редактирования уже существующей записи
                {
                    isUpdateOrAddMode = "UPDATE";
                    try
                    {
                        idVarInt = int.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        idVarInt = 0;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не удалось получить целое число ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    if (!((String.IsNullOrEmpty(erpPlantIdVarString) && String.IsNullOrEmpty(erpIdVarString)) ||
                        (!String.IsNullOrEmpty(erpPlantIdVarString) && !String.IsNullOrEmpty(erpIdVarString))))
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3, 4. Если указан ИД записи (режим редактирования записи), то \"Код завода SAP\" и \"Код ресурса/склада SAP\" должны быть" +
                            " или оба пусты (тогда останутся без изменений), или оба заполнены. Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 3, 4 }, resultString);
                        rowNumber++;
                        continue;
                    }


                    foundSapEquipmentDTO = await _sapEquipmentRepository.Get(idVarInt);

                    if (foundSapEquipmentDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не найдена запись в справочнике с ИД записи равным " + idVarInt.ToString() + ".Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    changedSapEquipmentDTO.Id = idVarInt;
                    if (String.IsNullOrEmpty(erpPlantIdVarString + erpIdVarString))
                    {
                        changedSapEquipmentDTO.ErpPlantId = foundSapEquipmentDTO.ErpPlantId;
                        changedSapEquipmentDTO.ErpId = foundSapEquipmentDTO.ErpId;
                    }
                    else
                    {
                        if ((foundSapEquipmentDTO.ErpPlantId + foundSapEquipmentDTO.ErpId).Equals(erpPlantIdVarString + erpIdVarString))
                        {
                            changedSapEquipmentDTO.ErpPlantId = foundSapEquipmentDTO.ErpPlantId;
                            changedSapEquipmentDTO.ErpId = foundSapEquipmentDTO.ErpId;
                        }
                        else
                        {
                            needCheckErpPlantIdPlusErpId = true;
                            changedSapEquipmentDTO.ErpPlantId = erpPlantIdVarString;
                            changedSapEquipmentDTO.ErpId = erpIdVarString;
                        }
                    }

                    if (String.IsNullOrEmpty(nameVarString))
                    {
                        changedSapEquipmentDTO.Name = foundSapEquipmentDTO.Name;
                    }
                    else
                    {
                        if (foundSapEquipmentDTO.Name.Equals(nameVarString))
                            changedSapEquipmentDTO.Name = foundSapEquipmentDTO.Name;
                        else
                        {
                            needCheckName = true;
                            changedSapEquipmentDTO.Name = nameVarString;
                        }
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(erpPlantIdVarString + erpIdVarString))
                    {
                        foundSapEquipmentDTO = await _sapEquipmentRepository.GetByResource(erpPlantIdVarString, erpIdVarString);

                        if (foundSapEquipmentDTO == null)  // добавление
                        {
                            isUpdateOrAddMode = "ADD";
                            needCheckErpPlantIdPlusErpId = true;
                            needCheckName = true;
                            changedSapEquipmentDTO.ErpPlantId = erpPlantIdVarString;
                            changedSapEquipmentDTO.ErpId = erpIdVarString;
                            changedSapEquipmentDTO.Name = nameVarString;

                            if (String.IsNullOrEmpty(changedSapEquipmentDTO.Name))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Наименование Ресурса SAP не может быть пустым. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 5 }, resultString);
                                rowNumber++;
                                continue;
                            }


                        }
                        else  // редактирование
                        {
                            isUpdateOrAddMode = "UPDATE";
                            needCheckErpPlantIdPlusErpId = false;
                            needCheckName = true;
                            changedSapEquipmentDTO.Id = foundSapEquipmentDTO.Id;
                            changedSapEquipmentDTO.ErpPlantId = foundSapEquipmentDTO.ErpPlantId;
                            changedSapEquipmentDTO.ErpId = foundSapEquipmentDTO.ErpId;
                            changedSapEquipmentDTO.Name = nameVarString;
                        }
                    }
                }

                if (needCheckErpPlantIdPlusErpId)
                {
                    SapEquipmentDTO? objectForCheckErpPlantIdPlusErpId = _sapEquipmentRepository.GetByResource(erpPlantIdVarString, erpIdVarString).Result;

                    if (objectForCheckErpPlantIdPlusErpId != null)
                    {
                        bool isBad = false;
                        if (isUpdateOrAddMode == "UPDATE")
                        {
                            if (objectForCheckErpPlantIdPlusErpId.Id != foundSapEquipmentDTO.Id)
                            {
                                isBad = true;
                            }
                        }
                        else  // если добавление
                        {
                            isBad = true;
                        }
                        if (isBad)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 3, 4. Уже есть запись с сочетанием \"Код завода SAP\" + \"Код ресурса/склада SAP\" равным " +
                                "\"" + erpPlantIdVarString + "\" + \"" + erpIdVarString + "\". ИД записи: " + objectForCheckErpPlantIdPlusErpId.Id.ToString() +
                                ". Наименование: " + objectForCheckErpPlantIdPlusErpId.Name + ". Изменения не применялись."; ;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 3, 4 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }
                }

                if (needCheckName)
                {
                    SapEquipmentDTO? objectForCheckName = _sapEquipmentRepository.GetByName(nameVarString).Result;

                    if (objectForCheckName != null)
                    {
                        bool isBad = false;
                        if (isUpdateOrAddMode == "UPDATE")
                        {
                            if (objectForCheckName.Id != foundSapEquipmentDTO.Id)
                            {
                                isBad = true;
                            }
                        }
                        else  // это добавление записи
                        {
                            isBad = true;
                        }
                        if (isBad)
                        {
                            duplicateNameString = "Предупреждение: уже был ресурс с таким наименованием.";
                        }
                    }
                }

                switch (isUpdateOrAddMode)
                {
                    case "ADD":
                        {
                            if (String.IsNullOrEmpty(isArchiveVarString))
                                changedSapEquipmentDTO.IsArchive = false;
                            else
                                changedSapEquipmentDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                            if (String.IsNullOrEmpty(isWarehouseVarString))
                                changedSapEquipmentDTO.IsWarehouse = false;
                            else
                                changedSapEquipmentDTO.IsWarehouse = isWarehouseVarString.ToUpper().Equals("ДА") ? true : false;

                            SapEquipmentDTO? newSapEquipmentDTO = await _sapEquipmentRepository.Create(changedSapEquipmentDTO);

                            await _logEventRepository.ToLog<SapEquipmentDTO>(oldObject: null, newObject: newSapEquipmentDTO, "Добавление ресурса SAP", "Ресурс SAP: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Ресурс Sap добавлен с кодом " + newSapEquipmentDTO.Id.ToString()
                                + ". " + duplicateNameString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            if (String.IsNullOrEmpty(duplicateNameString))
                            {
                                worksheet.Cell(rowNumber, 1).Value = "ОК";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 1).Value = "!!! ОК";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 5).Style.Font.FontColor = XLColor.YellowGreen;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "UPDATE":
                        {
                            if (String.IsNullOrEmpty(isArchiveVarString))
                            {
                                changedSapEquipmentDTO.IsArchive = false;
                            }
                            else
                            {
                                changedSapEquipmentDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                            }

                            if (String.IsNullOrEmpty(isWarehouseVarString))
                                changedSapEquipmentDTO.IsWarehouse = false;
                            else
                                changedSapEquipmentDTO.IsWarehouse = isWarehouseVarString.ToUpper().Equals("ДА") ? true : false;

                            await _sapEquipmentRepository.Update(changedSapEquipmentDTO, SD.UpdateMode.Update);

                            if (changedSapEquipmentDTO.IsArchive != foundSapEquipmentDTO.IsArchive)
                            {
                                SD.UpdateMode updMode;
                                if (changedSapEquipmentDTO.IsArchive == true)
                                {
                                    updMode = SD.UpdateMode.MoveToArchive;
                                }
                                else
                                {
                                    updMode = SD.UpdateMode.RestoreFromArchive;
                                }
                                await _sapEquipmentRepository.Update(changedSapEquipmentDTO, updMode);
                                await _logEventRepository.ToLog<SapEquipmentDTO>(oldObject: foundSapEquipmentDTO, newObject: changedSapEquipmentDTO, "Изменение Ресурса SAP", "Материал SAP: ", _authorizationRepository);
                            }

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. " + duplicateNameString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            if (String.IsNullOrEmpty(duplicateNameString))
                            {
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 1).Value = "!!! ОК";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            resultString = "!!! Для строки " + rowNumber.ToString() + " определен не предусмотренный режим обработки = " + isUpdateOrAddMode + ". Изменения не производились.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 1 }, resultString);
                            rowNumber++;
                            continue;
                        }
                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа SapEquipment в справочник Ресурсов SAP");
            await loadFromExcelPage.RefreshSate();

            return haveErrors;
        }


        public async Task<bool> MesParamExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {

            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в справочник Тэгов СИР");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;

            int resultColumnNumber = 34;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string idVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string codeVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string nameVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string descriptionVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string mesParamSourceTypeNameVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string mesParamSourceLinkVarString = rowVar.Cell(7).CachedValue.ToString().Trim();
                string departmentIdVarString = rowVar.Cell(8).CachedValue.ToString().Trim();
                string departmentNameVarString = rowVar.Cell(9).CachedValue.ToString().Trim();

                string sapEquipmentIdSourceVarString = rowVar.Cell(10).CachedValue.ToString().Trim();
                string erpPlantIdSourceVarString = rowVar.Cell(11).CachedValue.ToString().Trim();
                string erpIdSourceVarString = rowVar.Cell(12).CachedValue.ToString().Trim();
                string erpNameSourceVarString = rowVar.Cell(13).CachedValue.ToString().Trim();

                string sapEquipmentIdDestVarString = rowVar.Cell(14).CachedValue.ToString().Trim();
                string erpPlantIdDestVarString = rowVar.Cell(15).CachedValue.ToString().Trim();
                string erpIdDestVarString = rowVar.Cell(16).CachedValue.ToString().Trim();
                string erpNameDestVarString = rowVar.Cell(17).CachedValue.ToString().Trim();

                string sapMaterialIdVarString = rowVar.Cell(18).CachedValue.ToString().Trim();
                string sapMaterialCodeVarString = rowVar.Cell(19).CachedValue.ToString().Trim();
                string sapMaterialNameVarString = rowVar.Cell(20).CachedValue.ToString().Trim();

                string sapUnitOfMeasureNameVarString = rowVar.Cell(21).CachedValue.ToString().Trim();
                string daysRequestInPastVarString = rowVar.Cell(22).CachedValue.ToString().Trim();

                string TIVarString = rowVar.Cell(23).CachedValue.ToString().Trim();
                string nameTIVarString = rowVar.Cell(24).CachedValue.ToString().Trim();

                string TMVarString = rowVar.Cell(25).CachedValue.ToString().Trim();
                string nameTMVarString = rowVar.Cell(26).CachedValue.ToString().Trim();

                string mesToSirUnitOfMeasureKoefVarString = rowVar.Cell(27).CachedValue.ToString().Trim();

                string needWriteToSapVarString = rowVar.Cell(28).CachedValue.ToString().Trim();
                string needReadFromSapVarString = rowVar.Cell(29).CachedValue.ToString().Trim();

                string needReadFromMesVarString = rowVar.Cell(30).CachedValue.ToString().Trim();
                string needWriteToMesVarString = rowVar.Cell(31).CachedValue.ToString().Trim();

                string isNdoVarString = rowVar.Cell(32).CachedValue.ToString().Trim();
                string isArchiveVarString = rowVar.Cell(33).CachedValue.ToString().Trim();


                string resultString = "";
                int idVarInt = 0;


                // монструозный if. Наверное лучше какой-нить Dictionary. Но пока так.
                // Нет данных во всех 32-х колонках - выходим
                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(codeVarString) && String.IsNullOrEmpty(nameVarString)
                   && String.IsNullOrEmpty(descriptionVarString)
                   && String.IsNullOrEmpty(mesParamSourceTypeNameVarString) && String.IsNullOrEmpty(mesParamSourceLinkVarString))
                    if (String.IsNullOrEmpty(departmentIdVarString) && String.IsNullOrEmpty(departmentNameVarString) && String.IsNullOrEmpty(sapEquipmentIdSourceVarString)
                       && String.IsNullOrEmpty(erpPlantIdSourceVarString)
                       && String.IsNullOrEmpty(erpIdSourceVarString) && String.IsNullOrEmpty(erpNameSourceVarString))
                        if (String.IsNullOrEmpty(sapEquipmentIdDestVarString) && String.IsNullOrEmpty(erpPlantIdDestVarString) && String.IsNullOrEmpty(erpIdDestVarString)
                           && String.IsNullOrEmpty(erpNameDestVarString)
                           && String.IsNullOrEmpty(sapMaterialIdVarString) && String.IsNullOrEmpty(sapMaterialCodeVarString) && String.IsNullOrEmpty(sapMaterialNameVarString))
                            if (String.IsNullOrEmpty(sapUnitOfMeasureNameVarString) && String.IsNullOrEmpty(daysRequestInPastVarString) && String.IsNullOrEmpty(TIVarString)
                               && String.IsNullOrEmpty(nameTIVarString)
                               && String.IsNullOrEmpty(TMVarString) && String.IsNullOrEmpty(nameTMVarString) && String.IsNullOrEmpty(mesToSirUnitOfMeasureKoefVarString))
                                if (String.IsNullOrEmpty(needWriteToSapVarString) && String.IsNullOrEmpty(needReadFromSapVarString) && String.IsNullOrEmpty(needReadFromMesVarString)
                                   && String.IsNullOrEmpty(needWriteToMesVarString)
                                   && String.IsNullOrEmpty(isNdoVarString) && String.IsNullOrEmpty(isArchiveVarString))
                                {
                                    isEmptyString = true;
                                    continue;
                                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(codeVarString))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3. Пустые одновременно поля \"ИД записи\" и \"Код тэга СИР\". Изменения не применялись.";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                    rowNumber++;
                    continue;
                }

                string isUpdateOrAddMode = "NONE";
                bool needCheckCode = true;
                bool needCheckMesParamSourceLink = true;
                string duplicateMesParamSourceLink = "";

                IEnumerable<MesDepartmentDTO>? foundMesDepartmentList = null;
                MesDepartmentDTO? foundMesDepartment = null;
                MesParamSourceTypeDTO? foundMesParamSourceTypeDTO = null;
                SapEquipmentDTO? foundSapEquipmentSourceDTO = null;
                SapEquipmentDTO? foundSapEquipmentDestDTO = null;
                SapMaterialDTO? foundSapMaterialDTO = null;
                SapUnitOfMeasureDTO? foundSapUnitOfMeasureDTO = null;

                MesParamDTO? foundMesParamDTO = null;
                MesParamDTO changedMesParamDTO = new MesParamDTO();
                if (!String.IsNullOrEmpty(idVarString))  // если указан id элемента, то рассматриваем только вариант редактирования уже существующей записи
                {
                    isUpdateOrAddMode = "UPDATE";
                    try
                    {
                        idVarInt = int.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        idVarInt = 0;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не удалось получить целое число ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }

                    foundMesParamDTO = await _mesParamRepository.GetById(idVarInt);

                    if (foundMesParamDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не найдена запись в справочнике с ИД записи равным " + idVarInt.ToString() + ". Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    changedMesParamDTO.Id = foundMesParamDTO.Id;
                    if (String.IsNullOrEmpty(codeVarString))
                    {
                        changedMesParamDTO.Code = foundMesParamDTO.Code;
                    }
                    else
                    {
                        changedMesParamDTO.Code = codeVarString;
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(codeVarString))
                    {
                        foundMesParamDTO = await _mesParamRepository.GetByCode(codeVarString);

                        if (foundMesParamDTO == null)  // добавление
                        {
                            isUpdateOrAddMode = "ADD";
                            changedMesParamDTO.Code = codeVarString;
                        }
                        else  // редактирование
                        {
                            isUpdateOrAddMode = "UPDATE";
                            needCheckCode = true;

                            changedMesParamDTO.Id = foundMesParamDTO.Id;
                            changedMesParamDTO.Code = codeVarString;
                        }
                    }
                }

                if ((!String.IsNullOrEmpty(mesParamSourceTypeNameVarString) && String.IsNullOrEmpty(mesParamSourceLinkVarString)) ||
                    (String.IsNullOrEmpty(mesParamSourceTypeNameVarString) && !String.IsNullOrEmpty(mesParamSourceLinkVarString)))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 6, 7. Поля \"Источник\" и \"Тэг/ИД источника\" могут быть или оба проставлены, или оба пустые. Изменения не применялись.";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 6, 7 }, resultString);
                    rowNumber++;
                    continue;

                }
                if (needCheckCode)
                {
                    MesParamDTO? objectForCheckCode;
                    objectForCheckCode = _mesParamRepository.GetByCode(codeVarString).Result;

                    if (objectForCheckCode != null)
                    {
                        bool isBad = false;
                        if (isUpdateOrAddMode == "UPDATE")
                        {
                            if (objectForCheckCode.Id != foundMesParamDTO.Id)
                            {
                                isBad = true;
                            }
                        }
                        else  // если добавление
                        {
                            isBad = true;
                        }
                        if (isBad)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 3. Уже есть запись с кодом Тэга СИР " + codeVarString
                                + ". ИД записи: " + objectForCheckCode.Id.ToString() +
                                ". Наименование: " + objectForCheckCode.Name + ". Изменения не применялись."; ;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 3 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }
                }


                if (needCheckMesParamSourceLink)
                {
                    if (!String.IsNullOrEmpty(mesParamSourceLinkVarString) && !String.IsNullOrEmpty(mesParamSourceTypeNameVarString))
                    {

                        MesParamDTO? objectForCheckMesParamSourceLink = _mesParamRepository.GetByMesParamSourceLink(mesParamSourceLinkVarString).GetAwaiter().GetResult();

                        if (objectForCheckMesParamSourceLink != null)
                        {
                            bool isBad = false;
                            if (isUpdateOrAddMode == "UPDATE")
                            {
                                if (objectForCheckMesParamSourceLink.Id != foundMesParamDTO.Id)
                                {
                                    isBad = true;
                                }
                            }
                            else  // это добавление записи
                            {
                                isBad = true;
                            }
                            if (isBad)
                            {
                                duplicateMesParamSourceLink = "Предупреждение: уже был тэг СИР с таким \"Тэг/ИД источника.\"";
                            }
                        }
                    }
                }

                changedMesParamDTO.Name = nameVarString;
                changedMesParamDTO.Description = descriptionVarString;

                if (String.IsNullOrEmpty(mesParamSourceTypeNameVarString))
                {
                    changedMesParamDTO.MesParamSourceLink = null;
                    changedMesParamDTO.MesParamSourceTypeDTOFK = null;
                }
                else
                {
                    foundMesParamSourceTypeDTO = _mesParamSourceTypeRepository.GetByName(mesParamSourceTypeNameVarString).GetAwaiter().GetResult();
                    if (foundMesParamSourceTypeDTO == null)
                    {
                        changedMesParamDTO.MesParamSourceType = null;
                        changedMesParamDTO.MesParamSourceTypeDTOFK = null;
                    }
                    else
                    {
                        changedMesParamDTO.MesParamSourceType = foundMesParamSourceTypeDTO.Id;
                        changedMesParamDTO.MesParamSourceTypeDTOFK = foundMesParamSourceTypeDTO;
                    }
                }

                if (String.IsNullOrEmpty(mesParamSourceLinkVarString))
                {
                    changedMesParamDTO.MesParamSourceLink = null;
                }
                else
                {
                    changedMesParamDTO.MesParamSourceLink = mesParamSourceLinkVarString;
                }

                if (String.IsNullOrEmpty(departmentIdVarString) && String.IsNullOrEmpty(departmentNameVarString))
                {
                    changedMesParamDTO.DepartmentId = null;
                    changedMesParamDTO.MesDepartmentDTOFK = null;
                }
                else
                {
                    if (!String.IsNullOrEmpty(departmentIdVarString))
                    {
                        int departmentIdVarInt;
                        try
                        {
                            departmentIdVarInt = int.Parse(departmentIdVarString);
                        }
                        catch (Exception ex)
                        {
                            haveErrors = true;
                            departmentIdVarInt = 0;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 8. Не удалось получить целое число ИД производства." +
                                " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 8 }, resultString);
                            rowNumber++;
                            continue;
                        }

                        foundMesDepartment = _mesDepartmentRepository.GetById(departmentIdVarInt).GetAwaiter().GetResult();

                        if (foundMesDepartment == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 8. Не удалось найти производство с \"ИД производства\" равным " + departmentIdVarInt.ToString() + ". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 8 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundMesDepartment == null)
                    {
                        if (!String.IsNullOrEmpty(departmentNameVarString))
                        {
                            foundMesDepartmentList = _mesDepartmentRepository.GetByNameList(departmentNameVarString).GetAwaiter().GetResult();
                            if (foundMesDepartmentList != null)
                            {
                                if (foundMesDepartmentList.Count() > 1)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9. Найдено более одного производства с наименованием равным " + departmentNameVarString + ". Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                foundMesDepartment = foundMesDepartmentList.First();
                            }
                            if (foundMesDepartment == null)
                            {
                                foundMesDepartmentList = null;
                                foundMesDepartmentList = _mesDepartmentRepository.GetByShortNameList(departmentNameVarString).GetAwaiter().GetResult();
                                if (foundMesDepartmentList != null)
                                {
                                    if (foundMesDepartmentList.Count() > 1)
                                    {
                                        haveErrors = true;
                                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 9. Найдено более одного производства с сокращённым наименованием равным " + departmentNameVarString + ". Изменения не применялись.";
                                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 9 }, resultString);
                                        rowNumber++;
                                        continue;
                                    }
                                    foundMesDepartment = foundMesDepartmentList.First();
                                }
                            }
                        }
                    }
                    if (foundMesDepartment == null && (!String.IsNullOrEmpty(departmentIdVarString) || !String.IsNullOrEmpty(departmentNameVarString)))
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 8, 9. не найдено производство. Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 8, 9 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    else
                    {
                        changedMesParamDTO.DepartmentId = foundMesDepartment.Id;
                        changedMesParamDTO.MesDepartmentDTOFK = foundMesDepartment;
                    }
                }

                if (String.IsNullOrEmpty(sapEquipmentIdSourceVarString) && String.IsNullOrEmpty(erpPlantIdSourceVarString) && String.IsNullOrEmpty(erpIdSourceVarString) && String.IsNullOrEmpty(erpNameSourceVarString))
                {
                    changedMesParamDTO.SapEquipmentIdSource = null;
                    changedMesParamDTO.SapEquipmentSourceDTOFK = null;
                }
                else
                {
                    if (!String.IsNullOrEmpty(sapEquipmentIdSourceVarString))
                    {
                        int sapEquipmentIdSourceVarInt;
                        try
                        {
                            sapEquipmentIdSourceVarInt = int.Parse(sapEquipmentIdSourceVarString);
                        }
                        catch (Exception ex)
                        {
                            haveErrors = true;
                            sapEquipmentIdSourceVarInt = 0;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 10. Не удалось получить целое число \"ИД ресурса-источника SAP\"" +
                                " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 10 }, resultString);
                            rowNumber++;
                            continue;
                        }

                        foundSapEquipmentSourceDTO = _sapEquipmentRepository.Get(sapEquipmentIdSourceVarInt).GetAwaiter().GetResult();

                        if (foundSapEquipmentSourceDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 10. Не найден Ресурс-источник SAP c \"ИД ресурса-источника SAP\" равным " + sapEquipmentIdSourceVarInt.ToString() + ". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 10 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapEquipmentSourceDTO == null && (!String.IsNullOrEmpty(erpPlantIdSourceVarString) || !String.IsNullOrEmpty(erpIdSourceVarString)))
                    {
                        if (!String.IsNullOrEmpty(erpPlantIdSourceVarString) && !String.IsNullOrEmpty(erpIdSourceVarString))
                        {
                            foundSapEquipmentSourceDTO = await _sapEquipmentRepository.GetByResource(erpPlantIdSourceVarString, erpIdSourceVarString);
                            if (foundSapEquipmentSourceDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 11, 12. Ресурс SAP с \"Кодом завода ресурса-источника SAP\" равным " + erpPlantIdSourceVarString +
                                    " и \"Кодом ресурса/склада ресурса-источника SAP\" равным " + erpIdSourceVarString + " не найден в справочнике Ресурсов SAP. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 11, 12 }, resultString);
                                rowNumber++;
                                continue;
                            }
                        }
                        else
                        {
                            if (!String.IsNullOrEmpty(erpPlantIdSourceVarString) || !String.IsNullOrEmpty(erpIdSourceVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 11, 12. Поля \"Код завода ресурса-источника SAP\" и \"Код ресурса/склада ресурса-источника SAP\" должны быть или оба пустые, или оба заполнены. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 11, 12 }, resultString);
                                rowNumber++;
                                continue;
                            }
                        }
                    }

                    if (foundSapEquipmentSourceDTO == null && !String.IsNullOrEmpty(erpNameSourceVarString))
                    {
                        foundSapEquipmentSourceDTO = await _sapEquipmentRepository.GetByName(erpNameSourceVarString);
                        if (foundSapEquipmentSourceDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 13. Ресурс-источник SAP с наименованием равным " + erpNameSourceVarString +
                                " не найден в справочнике Ресурсы SAP. Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 13 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapEquipmentSourceDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 10, 11, 12, 13. Ресурс-источник SAP не найден в справочнике Ресурсы SAP. Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[4] { 10, 11, 12, 13 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    else
                    {
                        changedMesParamDTO.SapEquipmentIdSource = foundSapEquipmentSourceDTO.Id;
                        changedMesParamDTO.SapEquipmentSourceDTOFK = foundSapEquipmentSourceDTO;
                    }
                }

                if (String.IsNullOrEmpty(sapEquipmentIdDestVarString) && String.IsNullOrEmpty(erpPlantIdDestVarString) && String.IsNullOrEmpty(erpIdDestVarString) && String.IsNullOrEmpty(erpNameDestVarString))
                {
                    changedMesParamDTO.SapEquipmentIdDest = null;
                    changedMesParamDTO.SapEquipmentDestDTOFK = null;
                }
                else
                {
                    if (!String.IsNullOrEmpty(sapEquipmentIdDestVarString))
                    {
                        int sapEquipmentIdDestVarInt;
                        try
                        {
                            sapEquipmentIdDestVarInt = int.Parse(sapEquipmentIdDestVarString);
                        }
                        catch (Exception ex)
                        {
                            haveErrors = true;
                            sapEquipmentIdDestVarInt = 0;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 14. Не удалось получить целое число \"ИД ресурса-приёмника SAP\"" +
                                " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 14 }, resultString);
                            rowNumber++;
                            continue;
                        }

                        foundSapEquipmentDestDTO = _sapEquipmentRepository.Get(sapEquipmentIdDestVarInt).GetAwaiter().GetResult();

                        if (foundSapEquipmentDestDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 14. Не найден Ресурс-приёмник SAP c \"ИД ресурса-приёмника SAP\" равным " + sapEquipmentIdDestVarInt.ToString() + ". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 14 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapEquipmentDestDTO == null && (!String.IsNullOrEmpty(erpPlantIdDestVarString) || !String.IsNullOrEmpty(erpIdDestVarString)))
                    {
                        if (!String.IsNullOrEmpty(erpPlantIdDestVarString) && !String.IsNullOrEmpty(erpIdDestVarString))
                        {
                            foundSapEquipmentDestDTO = await _sapEquipmentRepository.GetByResource(erpPlantIdDestVarString, erpIdDestVarString);
                            if (foundSapEquipmentDestDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 15, 16. Ресурс SAP с \"Кодом завода ресурса-приёмника SAP\" равным " + erpPlantIdDestVarString +
                                    " и \"Кодом ресурса/склада ресурса-приёмника SAP\" равным " + erpIdDestVarString + " не найден в справочнике Ресурсов SAP. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 15, 16 }, resultString);
                                rowNumber++;
                                continue;
                            }
                        }
                        else
                        {
                            if (!String.IsNullOrEmpty(erpPlantIdDestVarString) || !String.IsNullOrEmpty(erpIdDestVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 15, 16. Поля \"Код завода ресурса-приёмника SAP\" и \"Код ресурса/склада ресурса-приёмника SAP\" должны быть или оба пустые, или оба заполнены. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 15, 16 }, resultString);
                                rowNumber++;
                                continue;
                            }
                        }
                    }

                    if (foundSapEquipmentDestDTO == null && !String.IsNullOrEmpty(erpNameDestVarString))
                    {
                        foundSapEquipmentDestDTO = await _sapEquipmentRepository.GetByName(erpNameDestVarString);
                        if (foundSapEquipmentDestDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 17. Ресурс-приёмник SAP с наименованием равным " + erpNameDestVarString +
                                " не найден в справочнике Ресурсы SAP. Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 17 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapEquipmentDestDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 14, 15, 16, 17. Ресурс-приёмник SAP не найден в справочнике Ресурсы SAP. Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[4] { 14, 15, 16, 17 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    else
                    {
                        changedMesParamDTO.SapEquipmentIdDest = foundSapEquipmentDestDTO.Id;
                        changedMesParamDTO.SapEquipmentDestDTOFK = foundSapEquipmentDestDTO;
                    }
                }

                if (String.IsNullOrEmpty(sapMaterialIdVarString) && String.IsNullOrEmpty(sapMaterialCodeVarString) && String.IsNullOrEmpty(sapMaterialNameVarString))
                {
                    changedMesParamDTO.SapMaterialId = null;
                    changedMesParamDTO.SapMaterialDTOFK = null;
                }
                else
                {
                    if (!String.IsNullOrEmpty(sapMaterialIdVarString))
                    {
                        int sapMaterialIdVarInt;
                        try
                        {
                            sapMaterialIdVarInt = int.Parse(sapMaterialIdVarString);
                        }
                        catch (Exception ex)
                        {
                            haveErrors = true;
                            sapMaterialIdVarInt = 0;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 18. Не удалось получить целое число \"ИД материала SAP\"" +
                                " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 18 }, resultString);
                            rowNumber++;
                            continue;
                        }

                        foundSapMaterialDTO = _sapMaterialRepository.Get(sapMaterialIdVarInt).GetAwaiter().GetResult();

                        if (foundSapMaterialDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 18. Не найден Материал SAP c \"ИД материала SAP\" равным " + sapMaterialIdVarInt.ToString() + ". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 18 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapMaterialDTO == null && !String.IsNullOrEmpty(sapMaterialCodeVarString))
                    {
                        foundSapMaterialDTO = await _sapMaterialRepository.GetByCode(sapMaterialCodeVarString);
                        if (foundSapMaterialDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 19. Материал SAP с кодом равным " + sapMaterialCodeVarString +
                                " не найден в справочнике Материалы SAP. Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 19 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapMaterialDTO == null && !String.IsNullOrEmpty(sapMaterialNameVarString))
                    {
                        foundSapMaterialDTO = await _sapMaterialRepository.GetByName(sapMaterialNameVarString);
                        if (foundSapMaterialDTO == null)
                        {
                            foundSapMaterialDTO = await _sapMaterialRepository.GetByShortName(sapMaterialNameVarString);
                            if (foundSapMaterialDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 20. Материал SAP с наименованием или сокр. наименованием " + sapMaterialNameVarString +
                                    " не найден в справочнике Материалы SAP. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 20 }, resultString);
                                rowNumber++;
                                continue;
                            }
                        }
                    }
                    if (foundSapMaterialDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 18, 19, 20. Материал SAP не найден в справочнике Материалы SAP. Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 18, 19, 20 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    else
                    {
                        changedMesParamDTO.SapMaterialId = foundSapMaterialDTO.Id;
                        changedMesParamDTO.SapMaterialDTOFK = foundSapMaterialDTO;
                    }
                }


                if ((changedMesParamDTO.SapEquipmentIdSource != null) || (changedMesParamDTO.SapEquipmentIdDest != null) || (changedMesParamDTO.SapMaterialId != null))
                {
                    if ((changedMesParamDTO.SapEquipmentIdSource != null) && (changedMesParamDTO.SapEquipmentIdDest != null) && (changedMesParamDTO.SapMaterialId != null))
                    {
                        var foundBySapMapping = await _mesParamRepository.GetBySapMappingNotInArchive(changedMesParamDTO.SapEquipmentIdSource, changedMesParamDTO.SapEquipmentIdDest, changedMesParamDTO.SapMaterialId, 0);

                        if (foundBySapMapping != null)
                        {
                            bool isBad = false;
                            if (isUpdateOrAddMode == "UPDATE")
                            {
                                if (foundBySapMapping.Id != foundMesParamDTO.Id)
                                {
                                    isBad = true;
                                }
                            }
                            else  // это добавление записи
                            {
                                isBad = true;
                            }
                            if (isBad == true)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ". Уже есть не архивный Тэг СИР с таким же маппингом \"Ресурс-источник SAP + Ресурс-приёмник SAP + Материал SAP\" ( КОД: " + foundBySapMapping.Code.ToString() + " НАИМЕНОВАНИЕ: " + foundBySapMapping.Name + " ИД: " + foundBySapMapping.Id.ToString() + "). Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber,
                                    new int[11] { 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 }, resultString);
                                rowNumber++;
                                continue;
                            }
                        }
                    }
                    else
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ". Не полный мэппинг Тэга СИР с SAP по связке параметров \"Источник SAP + Приёмник SAP + Материал SAP\". Должны быть проставлены или все 3 параметра, или ни одного. Изменения не применялись.";
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber,
                                new int[11] { 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 }, resultString);
                        rowNumber++;
                        continue;
                    }
                }

                if (String.IsNullOrEmpty(sapUnitOfMeasureNameVarString))
                {
                    changedMesParamDTO.SapUnitOfMeasureId = null;
                    changedMesParamDTO.SapUnitOfMeasureDTOFK = null;
                }
                else
                {
                    foundSapUnitOfMeasureDTO = _sapUnitOfMeasureRepository.GetByName(sapUnitOfMeasureNameVarString).GetAwaiter().GetResult();
                    if (foundSapUnitOfMeasureDTO == null)
                    {
                        foundSapUnitOfMeasureDTO = _sapUnitOfMeasureRepository.GetByShortName(sapUnitOfMeasureNameVarString).GetAwaiter().GetResult();
                        if (foundSapUnitOfMeasureDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 21. Единица измерения SAP с наименованием или сокр. наименованием " + sapUnitOfMeasureNameVarString +
                                " не найдена в справочнике Единиц измерения SAP. Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 21 }, resultString);
                            rowNumber++;
                            continue;
                        }
                    }
                    if (foundSapUnitOfMeasureDTO != null)
                    {
                        changedMesParamDTO.SapUnitOfMeasureId = foundSapUnitOfMeasureDTO.Id;
                        changedMesParamDTO.SapUnitOfMeasureDTOFK = foundSapUnitOfMeasureDTO;
                    }
                }

                if (String.IsNullOrEmpty(daysRequestInPastVarString))
                {
                    changedMesParamDTO.DaysRequestInPast = null;
                }
                else
                {
                    int daysRequestInPastVarInt;
                    try
                    {
                        daysRequestInPastVarInt = int.Parse(daysRequestInPastVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        daysRequestInPastVarInt = 0;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 22. Не удалось получить целое число \"Глубина опроса в днях\"" +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 22 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    changedMesParamDTO.DaysRequestInPast = daysRequestInPastVarInt;
                }

                changedMesParamDTO.TI = String.IsNullOrEmpty(TIVarString) ? null : TIVarString;
                changedMesParamDTO.NameTI = String.IsNullOrEmpty(nameTIVarString) ? null : nameTIVarString;
                changedMesParamDTO.TM = String.IsNullOrEmpty(TMVarString) ? null : TMVarString;
                changedMesParamDTO.NameTM = String.IsNullOrEmpty(nameTMVarString) ? null : nameTMVarString;

                if (String.IsNullOrEmpty(mesToSirUnitOfMeasureKoefVarString))
                {
                    changedMesParamDTO.MesToSirUnitOfMeasureKoef = 1;
                }
                else
                {
                    decimal mesToSirUnitOfMeasureKoefVarInt;
                    try
                    {
                        mesToSirUnitOfMeasureKoefVarInt = decimal.Parse(mesToSirUnitOfMeasureKoefVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        mesToSirUnitOfMeasureKoefVarInt = 0;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 27. Не удалось получить число \"Коэффициент пересчёта данных по тэгу ед. изм. MES в ед. изм. СИР\"" +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 27 }, resultString);
                        rowNumber++;
                        continue;
                    }
                    changedMesParamDTO.MesToSirUnitOfMeasureKoef = mesToSirUnitOfMeasureKoefVarInt;
                }

                changedMesParamDTO.NeedWriteToSap = needWriteToSapVarString.ToUpper() == "ДА" ? true : false;
                changedMesParamDTO.NeedReadFromSap = needReadFromSapVarString.ToUpper() == "ДА" ? true : false;
                changedMesParamDTO.NeedReadFromMes = needReadFromMesVarString.ToUpper() == "ДА" ? true : false;
                changedMesParamDTO.NeedWriteToMes = needWriteToMesVarString.ToUpper() == "ДА" ? true : false;
                changedMesParamDTO.IsNdo = isNdoVarString.ToUpper() == "ДА" ? true : false;

                if ((changedMesParamDTO.IsNdo == true) && ((bool)changedMesParamDTO.NeedReadFromSap || (bool)changedMesParamDTO.NeedWriteToMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Параметр НДО\", \"Читать из SAP\", \"Передавать в MES\". Тэг СИР с признаком \"Параметр НДО\" не может иметь признаки \"Читать из SAP\" или \"Передавать в MES\"." +
                        " Изменения не применялись";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 29, 31, 32 }, resultString);
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedReadFromMes && (bool)changedMesParamDTO.NeedWriteToMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Читать из MES\", \"Передавать в MES\". Тэг СИР не может одновременно иметь включенными признаки \"Читать из MES\" и \"Передавать в MES\"." +
                        " Изменения не применялись";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 29, 31 }, resultString);
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedReadFromSap && (bool)changedMesParamDTO.NeedWriteToSap))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Передавать в SAP\", \"Читать из SAP\". Тэг СИР не может одновременно иметь включенными признаки \"Передавать в SAP\" и \"Читать из SAP\"." +
                        " Изменения не применялись";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 28, 29 }, resultString);
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedWriteToSap && (bool)changedMesParamDTO.NeedWriteToMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Передавать в SAP\", \"Передавать в MES\". Тэг СИР не может одновременно иметь включенными признаки \"Передавать в SAP\" и \"Передавать в MES\"." +
                        " Изменения не применялись";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 28, 31 }, resultString);
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedReadFromSap && (bool)changedMesParamDTO.NeedReadFromMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Читать из SAP\", \"Читать из MES\". Тэг СИР не может одновременно иметь включенными признаки \"Читать из SAP\" и \"Читать из MES\"." +
                        " Изменения не применялись";
                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 29, 30 }, resultString);
                    rowNumber++;
                    continue;
                }

                switch (isUpdateOrAddMode)
                {
                    case "ADD":
                        {
                            if (String.IsNullOrEmpty(isArchiveVarString))
                                changedMesParamDTO.IsArchive = false;
                            else
                                changedMesParamDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                            MesParamDTO? newMesParamDTO = await _mesParamRepository.Create(changedMesParamDTO);

                            await _logEventRepository.ToLog<MesParamDTO>(oldObject: null, newObject: newMesParamDTO, "Добавление тэга СИР", "Тэг СИР: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Тэг СИР добавлен с ИД " + newMesParamDTO.Id.ToString()
                                + ". " + duplicateMesParamSourceLink;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            if (String.IsNullOrEmpty(duplicateMesParamSourceLink))
                            {
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 1).Value = "!!! OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "UPDATE":
                        {
                            changedMesParamDTO.Id = foundMesParamDTO.Id;
                            if (String.IsNullOrEmpty(isArchiveVarString))
                            {
                                changedMesParamDTO.IsArchive = false;
                            }
                            else
                            {
                                changedMesParamDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                            }
                            await _mesParamRepository.Update(changedMesParamDTO);

                            if (changedMesParamDTO.IsArchive != foundMesParamDTO.IsArchive)
                            {
                                SD.UpdateMode updMode;
                                if (changedMesParamDTO.IsArchive == true)
                                {
                                    updMode = SD.UpdateMode.MoveToArchive;
                                }
                                else
                                {
                                    updMode = SD.UpdateMode.RestoreFromArchive;
                                }
                                await _mesParamRepository.Delete(changedMesParamDTO.Id, updMode);
                                await _logEventRepository.ToLog<MesParamDTO>(oldObject: foundMesParamDTO, newObject: changedMesParamDTO, "Изменение тэга СИР", "Тэг СИР: ", _authorizationRepository);
                            }

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. " + duplicateMesParamSourceLink;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            if (String.IsNullOrEmpty(duplicateMesParamSourceLink))
                            {
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            resultString = "!!! Для строки " + rowNumber.ToString() + " определен не предусмотренный режим обработки = " + isUpdateOrAddMode + ". Изменения не производились.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 1 }, resultString);
                            rowNumber++;
                            continue;
                        }
                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа SapEquipment в справочник Ресурсов SAP");
            await loadFromExcelPage.RefreshSate();

            return haveErrors;

        }

        public async Task<bool> SapNdoOUTExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            return false;
        }

        public async Task<bool> MesNdoStocksExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            return false;
        }

        public async Task<bool> MesMovementsExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            return false;
        }

        public async Task<bool> SapMovementsOUTExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            return false;
        }

        public async Task<bool> UsersExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в Справочник пользователей");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string actionVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string idVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string loginVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string userNameVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string descriptionVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string isSyncWithADVarString = rowVar.Cell(7).CachedValue.ToString().Trim();
                string syncWithADGroupsLastTimeVarString = rowVar.Cell(8).CachedValue.ToString().Trim();
                string isServiceUserVarString = rowVar.Cell(9).CachedValue.ToString().Trim();
                string isArchiveVarString = rowVar.Cell(10).CachedValue.ToString().Trim();

                string resultString = "";
                Guid idVarGuid = Guid.Empty;

                int resultColumnNumber = 11;

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(loginVarString) && String.IsNullOrEmpty(userNameVarString)
                        && String.IsNullOrEmpty(descriptionVarString) && String.IsNullOrEmpty(isSyncWithADVarString)
                        && String.IsNullOrEmpty(syncWithADGroupsLastTimeVarString) && String.IsNullOrEmpty(isServiceUserVarString)
                        && String.IsNullOrEmpty(isArchiveVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                UserDTO? foundUserDTO = null;
                UserDTO changedUserDTO = new UserDTO();

                if (!String.IsNullOrEmpty(idVarString))
                {
                    try
                    {
                        idVarGuid = Guid.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        idVarGuid = Guid.Empty;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД Пользователя\"). Не удалось получить ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                        rowNumber++;
                        continue;
                    }
                }

                switch (actionVarString.Trim().ToUpper())
                {
                    case "ИЗМЕНИТЬ":
                        {
                            if (idVarGuid == Guid.Empty)
                            {
                                haveErrors = true;
                                idVarGuid = Guid.Empty;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД Пользователя\"). В режиме \"Изменить\" должен быть указан ИД Пользователя. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            foundUserDTO = await _userRepository.Get(idVarGuid);
                            if (foundUserDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД Пользователя\"). Не найден Пользователь с ИД: " + idVarGuid.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            changedUserDTO.Id = idVarGuid;

                            if (String.IsNullOrEmpty(loginVarString))
                                changedUserDTO.Login = foundUserDTO.Login;
                            else
                                changedUserDTO.Login = loginVarString;

                            UserDTO? objectForCheckLogin = _userRepository.GetByLogin(changedUserDTO.Login).Result;

                            if (objectForCheckLogin != null)
                            {
                                if (objectForCheckLogin.Id != objectForCheckLogin.Id)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Логин\"). Уже есть запись с таким логином. ИД записи: " + objectForCheckLogin.Id.ToString() + ". Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }

                            if (String.IsNullOrEmpty(userNameVarString))
                                changedUserDTO.UserName = foundUserDTO.UserName;
                            else
                                changedUserDTO.UserName = userNameVarString;

                            UserDTO? objectForCheckUserName = _userRepository.GetByUserName(changedUserDTO.UserName).Result;
                            if (objectForCheckUserName != null)
                            {
                                if (objectForCheckUserName.Id != objectForCheckUserName.Id)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ФИО\"). Уже есть запись с таким ФИО. ИД записи: " + objectForCheckUserName.Id.ToString() + ". Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }

                            changedUserDTO.Description = descriptionVarString;

                            if (String.IsNullOrEmpty(isSyncWithADVarString))
                            {
                                changedUserDTO.IsSyncWithAD = false;
                            }
                            else
                            {
                                changedUserDTO.IsSyncWithAD = isSyncWithADVarString.ToUpper().Equals("ДА") ? true : false;
                            }

                            if (!String.IsNullOrEmpty(syncWithADGroupsLastTimeVarString))
                            {
                                try
                                {
                                    changedUserDTO.SyncWithADGroupsLastTime = DateTime.Parse(syncWithADGroupsLastTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarGuid = Guid.Empty;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время последней синхронизации пользователя с AD\"). Не удалось получить Время последней синхронизации пользователя с AD." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }
                            else
                                changedUserDTO.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

                            if (String.IsNullOrEmpty(isServiceUserVarString))
                            {
                                changedUserDTO.IsServiceUser = false;
                            }
                            else
                            {
                                changedUserDTO.IsServiceUser = isServiceUserVarString.ToUpper().Equals("ДА") ? true : false;
                            }

                            if (String.IsNullOrEmpty(isArchiveVarString))
                            {
                                changedUserDTO.IsArchive = false;
                            }
                            else
                            {
                                changedUserDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                            }
                            await _userRepository.Update(changedUserDTO, SD.UpdateMode.Update);

                            if (changedUserDTO.IsArchive != foundUserDTO.IsArchive)
                            {
                                SD.UpdateMode updMode;
                                if (changedUserDTO.IsArchive == true)
                                {
                                    updMode = SD.UpdateMode.MoveToArchive;
                                }
                                else
                                {
                                    updMode = SD.UpdateMode.RestoreFromArchive;
                                }
                                await _userRepository.Update(changedUserDTO, updMode);
                                await _logEventRepository.ToLog<UserDTO>(oldObject: foundUserDTO, newObject: changedUserDTO, "Изменение пользователя", "Пользователь: ", _authorizationRepository);
                            }

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. ";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;

                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            break;
                        }
                    case "ДОБАВИТЬ":
                        {
                            if (idVarGuid != Guid.Empty)
                            {
                                UserDTO? objectForCheckId = _userRepository.Get(idVarGuid).Result;
                                if (objectForCheckId != null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД Пользователя\"). Уже есть запись с таким ИД пользователя. Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedUserDTO.Id = idVarGuid;
                            }

                            if (String.IsNullOrEmpty(loginVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Логин\"). В режиме добавления Логин не может быть пустым. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            UserDTO? objectForCheckLogin = _userRepository.GetByLogin(loginVarString).Result;

                            if (objectForCheckLogin != null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Логин\"). Уже есть запись с таким логином. ИД записи: " + objectForCheckLogin.Id.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                rowNumber++;
                                continue;
                            }
                            changedUserDTO.Login = loginVarString;

                            if (String.IsNullOrEmpty(loginVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Логин\"). В режиме добавления Логин не может быть пустым. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            UserDTO? objectForCheckUserName = _userRepository.GetByUserName(userNameVarString).Result;

                            if (objectForCheckUserName != null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ФИО\"). Уже есть запись с таким ФИО. ИД записи: " + objectForCheckUserName.Id.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                rowNumber++;
                                continue;
                            }
                            changedUserDTO.UserName = userNameVarString;
                            changedUserDTO.Description = descriptionVarString;

                            if (String.IsNullOrEmpty(isSyncWithADVarString))
                                changedUserDTO.IsSyncWithAD = false;
                            else
                                changedUserDTO.IsSyncWithAD = isSyncWithADVarString.ToUpper().Equals("ДА") ? true : false;

                            if (!String.IsNullOrEmpty(syncWithADGroupsLastTimeVarString))
                            {
                                try
                                {
                                    changedUserDTO.SyncWithADGroupsLastTime = DateTime.Parse(syncWithADGroupsLastTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarGuid = Guid.Empty;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время последней синхронизации пользователя с AD\"). Не удалось получить Время последней синхронизации пользователя с AD." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }
                            else
                                changedUserDTO.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;

                            if (changedUserDTO.SyncWithADGroupsLastTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue)
                                changedUserDTO.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue;
                            if (changedUserDTO.SyncWithADGroupsLastTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue)
                                changedUserDTO.SyncWithADGroupsLastTime = (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue;

                            if (String.IsNullOrEmpty(isServiceUserVarString))
                                changedUserDTO.IsServiceUser = false;
                            else
                                changedUserDTO.IsServiceUser = isServiceUserVarString.ToUpper().Equals("ДА") ? true : false;

                            if (String.IsNullOrEmpty(isArchiveVarString))
                                changedUserDTO.IsArchive = false;
                            else
                                changedUserDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                            UserDTO? newUserDTO = await _userRepository.Create(changedUserDTO);

                            await _logEventRepository.ToLog<UserDTO>(oldObject: null, newObject: newUserDTO, "Добавление пользователя", "Пользователь: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Пользователь добавлен. ИД " + newUserDTO.Id.ToString();
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "ОК";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 3).Value = newUserDTO.Id.ToString();
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);

                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            haveErrors = true;
                            idVarGuid = Guid.Empty;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 2 (\"Действие\"). Не предусмотренное значение действия = " + actionVarString + ". Для Справочника пользователей допустимы действия \"Добавить\" или \"Изменить\". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                            rowNumber++;
                            continue;
                        }
                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа " + worksheet.Name + " в Справочник пользователей");
            await loadFromExcelPage.RefreshSate();

            return haveErrors;
        }


        public async Task<bool> ADGroupsExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в справочник Групп AD");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string actionVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string idVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string nameVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string descriptionVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string isArchiveVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string resultString = "";
                Guid idVarGuid = Guid.Empty;

                int resultColumnNumber = 7;

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(nameVarString) && String.IsNullOrEmpty(descriptionVarString)
                        && String.IsNullOrEmpty(isArchiveVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                ADGroupDTO? foundADGroupDTO = null;
                ADGroupDTO changedADGroupDTO = new ADGroupDTO();

                if (!String.IsNullOrEmpty(idVarString))
                {
                    try
                    {
                        idVarGuid = Guid.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        idVarGuid = Guid.Empty;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД группы AD\"). Не удалось получить ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                        rowNumber++;
                        continue;
                    }
                }

                switch (actionVarString.Trim().ToUpper())
                {
                    case "ИЗМЕНИТЬ":
                        {
                            if (idVarGuid == Guid.Empty)
                            {
                                haveErrors = true;
                                idVarGuid = Guid.Empty;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД группы AD\"). В режиме \"Изменить\" должен быть указан ИД группы AD. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            foundADGroupDTO = await _adGroupRepository.Get(idVarGuid);
                            if (foundADGroupDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД группы AD\"). Не найдена Группа AD с ИД группы: " + idVarGuid.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            changedADGroupDTO.Id = idVarGuid;

                            if (String.IsNullOrEmpty(nameVarString))
                                changedADGroupDTO.Name = foundADGroupDTO.Name;
                            else
                                changedADGroupDTO.Name = nameVarString;

                            ADGroupDTO? objectForCheckName = _adGroupRepository.GetByName(changedADGroupDTO.Name).Result;

                            if (objectForCheckName != null)
                            {
                                if (objectForCheckName.Id != foundADGroupDTO.Id)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Наименование\"). Уже есть запись с таким наименованием. ИД записи: " + objectForCheckName.Id.ToString() + ". Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }
                            changedADGroupDTO.Description = descriptionVarString;

                            if (String.IsNullOrEmpty(isArchiveVarString))
                            {
                                changedADGroupDTO.IsArchive = false;
                            }
                            else
                            {
                                changedADGroupDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                            }
                            await _adGroupRepository.Update(changedADGroupDTO, SD.UpdateMode.Update);

                            if (changedADGroupDTO.IsArchive != foundADGroupDTO.IsArchive)
                            {
                                SD.UpdateMode updMode;
                                if (changedADGroupDTO.IsArchive == true)
                                {
                                    updMode = SD.UpdateMode.MoveToArchive;
                                }
                                else
                                {
                                    updMode = SD.UpdateMode.RestoreFromArchive;
                                }
                                await _adGroupRepository.Update(changedADGroupDTO, updMode);
                                await _logEventRepository.ToLog<ADGroupDTO>(oldObject: foundADGroupDTO, newObject: changedADGroupDTO, "Изменение группы AD", "Группа AD: ", _authorizationRepository);
                            }

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. ";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;

                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            break;
                        }
                    case "ДОБАВИТЬ":
                        {
                            if (idVarGuid != Guid.Empty)
                            {
                                ADGroupDTO? objectForCheckId = _adGroupRepository.Get(idVarGuid).Result;
                                if (objectForCheckId != null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД группы AD\"). Уже есть запись с таким ИД группы AD. Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedADGroupDTO.Id = idVarGuid;
                            }


                            if (String.IsNullOrEmpty(nameVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Наименование\"). В режиме добавления Наименование не может быть пустым. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            ADGroupDTO? objectForCheckName = _adGroupRepository.GetByName(nameVarString).Result;

                            if (objectForCheckName != null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Наименование\"). Уже есть запись с таким наименованием. ИД записи: " + objectForCheckName.Id.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                rowNumber++;
                                continue;
                            }
                            changedADGroupDTO.Name = nameVarString;
                            changedADGroupDTO.Description = descriptionVarString;

                            if (String.IsNullOrEmpty(isArchiveVarString))
                                changedADGroupDTO.IsArchive = false;
                            else
                                changedADGroupDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                            ADGroupDTO? newADGroupDTO = await _adGroupRepository.Create(changedADGroupDTO);

                            await _logEventRepository.ToLog<ADGroupDTO>(oldObject: null, newObject: newADGroupDTO, "Добавление группы AD", "Группа AD: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Группа AD добавлена ИД " + newADGroupDTO.Id.ToString();
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "ОК";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 3).Value = newADGroupDTO.Id.ToString();
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);

                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            haveErrors = true;
                            idVarGuid = Guid.Empty;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 2 (\"Действие\"). Не предусмотренное значение действия = " + actionVarString + ". Для Справочника групп AD допустимы действия \"Добавить\" или \"Изменить\". Изменения не применялись.";
                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[1] { 2 }, resultString);
                            rowNumber++;
                            continue;
                        }
                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа " + worksheet.Name + " в справочник Группы AD");
            await loadFromExcelPage.RefreshSate();

            return haveErrors;
        }

        public async Task WriteError(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            int rowNum, int exclamationColumn, int resultColumnNumber, int[] redColumns, string errorMessage)
        {
            worksheet.Cell(rowNum, resultColumnNumber).Value = "!!!";
            worksheet.Cell(rowNum, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
            worksheet.Cell(rowNum, exclamationColumn).Style.Font.SetBold(true);
            worksheet.Cell(rowNum, resultColumnNumber).Value = errorMessage;
            worksheet.Cell(rowNum, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
            worksheet.Cell(rowNum, resultColumnNumber).Style.Font.SetBold(true);

            foreach (var column in redColumns)
            {
                worksheet.Cell(rowNum, column).Style.Font.FontColor = XLColor.Red;
                worksheet.Cell(rowNum, column).Style.Font.SetBold(true);
            }
            loadFromExcelPage.console.Log(errorMessage);
            await loadFromExcelPage.RefreshSate();
        }

        //public int? GetIntValue(object obj, string propertyName)
        //{
        //    var varProperty = obj.GetType().GetProperty(propertyName);
        //    int propertyValue = varProperty.GetValue(obj, null) as int?;
        //    return propertyValue;
        //}

        //public string? GetStringValue(object obj, string propertyName)
        //{
        //    var varProperty = obj.GetType().GetProperty(propertyName);
        //    string? propertyValue = varProperty.GetValue(obj, null) as string;
        //    return propertyValue;
        //}

        //public bool? GetStringBool(object obj, string propertyName)
        //{
        //    var varProperty = obj.GetType().GetProperty(propertyName);
        //    bool? propertyValue = varProperty.GetValue(obj, null) as bool?;
        //    return propertyValue;
        //}


        //public void SetValue<T>(object obj, string propertyName, T? value)
        //{
        //    PropertyInfo varPropertyInfo = obj.GetType().GetProperty(propertyName);
        //    varPropertyInfo.SetValue(obj, Convert.ChangeType(value, varPropertyInfo.PropertyType), null);

        //}

        //public string? SetStringValue(object obj, string propertyName)
        //{
        //    var varProperty = obj.GetType().GetProperty(propertyName);
        //    string? propertyValue = varProperty.GetValue(obj, null) as string;
        //    return propertyValue;
        //}

        //public bool? SetStringBool(object obj, string propertyName)
        //{
        //    var varProperty = obj.GetType().GetProperty(propertyName);
        //    bool? propertyValue = varProperty.GetValue(obj, null) as bool?;
        //    return propertyValue;
        //}

    }
}
