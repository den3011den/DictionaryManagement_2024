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
        private readonly ISapNdoOUTRepository _sapNdoOUTRepository;
        private readonly IMesNdoStocksRepository _mesNdoStocksRepository;
        private readonly IReportEntityRepository _reportEntityRepository;
        private readonly IMesMovementsRepository _mesMovementsRepository;
        private readonly ISapMovementsINRepository _sapMovementsINRepository;
        private readonly ISapMovementsOUTRepository _sapMovementsOUTRepository;
        private readonly IDataSourceRepository _dataSourceRepository;
        private readonly IDataTypeRepository _dataTypeRepository;

        public LoadFromExcelRepository(ISapMaterialRepository sapMaterialRepository, IMesMaterialRepository mesMaterialRepository,
            ISapEquipmentRepository sapEquipmentRepository,
            ILogEventRepository logEventRepository, IMesParamRepository mesParamRepository,
            IMesParamSourceTypeRepository mesParamSourceTypeRepository,
            IMesDepartmentRepository mesDepartmentRepository,
            ISapUnitOfMeasureRepository sapUnitOfMeasureRepository,
            IADGroupRepository adGroupRepository,
            IUserRepository userRepository,
            ISapNdoOUTRepository sapNdoOUTRepository,
            IMesNdoStocksRepository mesNdoStocksRepository,
            IReportEntityRepository reportEntityRepository,
            IMesMovementsRepository mesMovementsRepository,
            ISapMovementsINRepository sapMovementsINRepository,
            ISapMovementsOUTRepository sapMovementsOUTRepository,
            IDataSourceRepository dataSourceRepository,
            IDataTypeRepository dataTypeRepository)
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
            _sapNdoOUTRepository = sapNdoOUTRepository;
            _mesNdoStocksRepository = mesNdoStocksRepository;
            _reportEntityRepository = reportEntityRepository;
            _mesMovementsRepository = mesMovementsRepository;
            _sapMovementsINRepository = sapMovementsINRepository;
            _sapMovementsOUTRepository = sapMovementsOUTRepository;
            _dataSourceRepository = dataSourceRepository;
            _dataTypeRepository = dataTypeRepository;

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

        public async Task<string> SapUnitOfMeasureReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Единиц измерения SAP)";
            await loadFromExcelPage.RefreshSate();

            IEnumerable<SapUnitOfMeasureDTO> sapUnitOfMeasureDTOList = (await _sapUnitOfMeasureRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.ShortName);

            int recordCount = sapUnitOfMeasureDTOList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var sapUnitOfMeasureDTOItem in sapUnitOfMeasureDTOList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    await loadFromExcelPage.RefreshSate();
                }

                excelColNum = 2;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTOItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTOItem.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTOItem.ShortName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapUnitOfMeasureDTOItem.IsArchive == true ? "Да" : "Нет";
                excelRowNum++;
            }
            return "SapUnitOfMeasureDTOItem_Example_with_data_";
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
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
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

                string duplicateNameString = "";

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
                            duplicateNameString = duplicateNameString + " Уже была запись с наименованием " + nameVarString + ". ИД записи: " + objectForCheckName.Id.ToString();
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
                            duplicateNameString = duplicateNameString + " Уже была запись с сокр. наименованием " + shortNameVarString + ". ИД записи: " + objectForCheckShortName.Id.ToString();
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
                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Материал добавлен с кодом " + newMaterialDTO.Id.ToString()
                                + ". " + duplicateNameString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;

                            if (String.IsNullOrEmpty(duplicateNameString))
                            {
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 2).Value = newMaterialDTO.Id.ToString();
                                worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Value = "!!! OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 2).Value = newMaterialDTO.Id.ToString();
                                worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
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

                            if (String.IsNullOrEmpty(duplicateNameString))
                            {
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
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
            int resultColumnNumber = 8;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
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
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
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

            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в витрину SAP НДО-выход");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;
            int resultColumnNumber = 10;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {
                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string actionVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string idVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string addTimeVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string tagNameVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string valueTimeVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string valueVarString = rowVar.Cell(7).CachedValue.ToString().Trim();
                string sapGoneVarString = rowVar.Cell(8).CachedValue.ToString().Trim();
                string sapGoneTimeVarString = rowVar.Cell(9).CachedValue.ToString().Trim();

                string resultString = "";
                Int64 idVarInt64 = 0;


                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(addTimeVarString) && String.IsNullOrEmpty(tagNameVarString)
                        && String.IsNullOrEmpty(valueTimeVarString) && String.IsNullOrEmpty(valueVarString)
                        && String.IsNullOrEmpty(sapGoneVarString) && String.IsNullOrEmpty(sapGoneTimeVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                SapNdoOUTDTO? foundSapNdoOUTDTO = null;
                SapNdoOUTDTO changedSapNdoOUTDTO = new SapNdoOUTDTO();

                if (!String.IsNullOrEmpty(idVarString))
                {
                    try
                    {
                        idVarInt64 = Int64.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        idVarInt64 = 0;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Не удалось получить ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                        rowNumber++;
                        continue;
                    }
                }

                string tagNameWarningString = "";

                switch (actionVarString.Trim().ToUpper())
                {
                    case "ИЗМЕНИТЬ":
                        {
                            if (idVarInt64 <= 0)
                            {
                                haveErrors = true;
                                idVarInt64 = 0;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). ИД записи должен быть положительным числом." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            foundSapNdoOUTDTO = await _sapNdoOUTRepository.GetById(idVarInt64);
                            if (foundSapNdoOUTDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Не найдена запись с ИД: " + idVarInt64.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            changedSapNdoOUTDTO.Id = idVarInt64;

                            DateTime addTimeVarDateTime;

                            if (!String.IsNullOrEmpty(addTimeVarString))
                            {
                                try
                                {
                                    addTimeVarDateTime = DateTime.Parse(addTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Время добавления записи\"). Не удалось получить Время добавления записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.AddTime = addTimeVarDateTime;
                            }
                            else
                                changedSapNdoOUTDTO.AddTime = foundSapNdoOUTDTO.AddTime;

                            changedSapNdoOUTDTO.AddTime = changedSapNdoOUTDTO.AddTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedSapNdoOUTDTO.AddTime;
                            changedSapNdoOUTDTO.AddTime = changedSapNdoOUTDTO.AddTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedSapNdoOUTDTO.AddTime;

                            if (!String.IsNullOrEmpty(tagNameVarString))
                            {
                                MesParamDTO? objectForCheckTagName = await _mesParamRepository.GetByCode(tagNameVarString);
                                if (objectForCheckTagName == null)
                                {
                                    tagNameWarningString = ". Внимание! Код тэга \"" + tagNameVarString + "\" не найден ни у одного тэга СИР.";
                                }
                                changedSapNdoOUTDTO.TagName = tagNameVarString;
                            }
                            else
                            {
                                changedSapNdoOUTDTO.TagName = foundSapNdoOUTDTO.TagName;
                            }

                            DateTime valueTimeVarDateTime;
                            if (!String.IsNullOrEmpty(valueTimeVarString))
                            {
                                try
                                {
                                    valueTimeVarDateTime = DateTime.Parse(valueTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"Время значения\"). Не удалось получить Время значения." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.ValueTime = valueTimeVarDateTime;
                            }
                            else
                                changedSapNdoOUTDTO.ValueTime = foundSapNdoOUTDTO.ValueTime;

                            changedSapNdoOUTDTO.ValueTime = changedSapNdoOUTDTO.ValueTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedSapNdoOUTDTO.ValueTime;
                            changedSapNdoOUTDTO.ValueTime = changedSapNdoOUTDTO.ValueTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedSapNdoOUTDTO.ValueTime;

                            decimal valueDecimal;
                            if (!String.IsNullOrEmpty(valueVarString))
                            {
                                try
                                {
                                    valueDecimal = decimal.Parse(valueVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Значение\"). Не удалось получить Значение." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.Value = valueDecimal;
                            }
                            else
                                changedSapNdoOUTDTO.Value = 0;

                            if (changedSapNdoOUTDTO.Value < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Значение\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (String.IsNullOrEmpty(sapGoneVarString))
                            {
                                changedSapNdoOUTDTO.SapGone = false;
                            }
                            else
                            {
                                changedSapNdoOUTDTO.SapGone = sapGoneVarString.ToUpper().Equals("ДА") ? true : false;
                            }

                            DateTime sapGoneTimeVarDateTime;

                            if (!String.IsNullOrEmpty(sapGoneTimeVarString))
                            {
                                try
                                {
                                    sapGoneTimeVarDateTime = DateTime.Parse(sapGoneTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Когда Sap забрал\"). Не удалось получить Когда Sap забрал." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.SapGoneTime = sapGoneTimeVarDateTime;
                                changedSapNdoOUTDTO.SapGoneTime = changedSapNdoOUTDTO.SapGoneTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedSapNdoOUTDTO.SapGoneTime;
                                changedSapNdoOUTDTO.SapGoneTime = changedSapNdoOUTDTO.SapGoneTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedSapNdoOUTDTO.SapGoneTime;
                            }
                            else
                                changedSapNdoOUTDTO.SapGoneTime = null;

                            if (!((changedSapNdoOUTDTO.SapGone == true && changedSapNdoOUTDTO.SapGoneTime != null)
                                || (changedSapNdoOUTDTO.SapGone != true && changedSapNdoOUTDTO.SapGoneTime == null)))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 8, 9 (\"Sap забрал\", \"Когда Sap забрал\"). Если \"Sap забрал\" равно \"Да\" время \"Когда Sap забрал\" должно быть заполнено."
                                    + " Если \"Sap забрал\" равно \"Нет\" или пусто, то время \"Когда Sap забрал\" должно быть пустым."
                                    + " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 2, 8, 9 }, resultString);
                                rowNumber++;
                                continue;

                            }

                            await _sapNdoOUTRepository.Update(changedSapNdoOUTDTO);
                            await _logEventRepository.ToLog<SapNdoOUTDTO>(oldObject: foundSapNdoOUTDTO, newObject: changedSapNdoOUTDTO, "Изменение записи в витрине SAP НДО-выход SapNdoOUT", "Запись: ", _authorizationRepository);


                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. " + tagNameWarningString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;

                            if (String.IsNullOrEmpty(tagNameWarningString))
                            {
                                worksheet.Cell(rowNumber, 1).Value = "ОК";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 1).Value = "!!! ОК";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 5).Style.Font.FontColor = XLColor.YellowGreen;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "ДОБАВИТЬ":
                        {
                            if (!String.IsNullOrEmpty(idVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Для действия \"Добавить\" ИД записи должно быть пустым. Сформируется автоматически." +
                                        " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            DateTime addTimeVarDateTime;

                            if (!String.IsNullOrEmpty(addTimeVarString))
                            {
                                try
                                {
                                    addTimeVarDateTime = DateTime.Parse(addTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Время добавления записи\"). Не удалось получить Время добавления записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.AddTime = addTimeVarDateTime;
                                changedSapNdoOUTDTO.AddTime = changedSapNdoOUTDTO.AddTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedSapNdoOUTDTO.AddTime;
                                changedSapNdoOUTDTO.AddTime = changedSapNdoOUTDTO.AddTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedSapNdoOUTDTO.AddTime;
                            }


                            if (!String.IsNullOrEmpty(tagNameVarString))
                            {
                                MesParamDTO? objectForCheckTagName = await _mesParamRepository.GetByCode(tagNameVarString);
                                if (objectForCheckTagName == null)
                                {
                                    tagNameWarningString = ". Внимание! Код тэга \"" + tagNameVarString + "\" не найден ни у одного тэга СИР.";
                                }
                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"Код тэга\"). В режиме \"Добавить\" Код тэга не может быть пустым." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            DateTime valueTimeVarDateTime;
                            if (!String.IsNullOrEmpty(valueTimeVarString))
                            {
                                try
                                {
                                    valueTimeVarDateTime = DateTime.Parse(valueTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"Время значения\"). Не удалось получить Время значения." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.ValueTime = valueTimeVarDateTime;
                                changedSapNdoOUTDTO.ValueTime = changedSapNdoOUTDTO.ValueTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedSapNdoOUTDTO.ValueTime;
                                changedSapNdoOUTDTO.ValueTime = changedSapNdoOUTDTO.ValueTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedSapNdoOUTDTO.ValueTime;

                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"Время значения\"). В режиме \"Добавить\" Время значения не может быть пустым." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                rowNumber++;
                                continue;
                            }


                            decimal valueDecimal;
                            if (!String.IsNullOrEmpty(valueVarString))
                            {
                                try
                                {
                                    valueDecimal = decimal.Parse(valueVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Значение\"). Не удалось получить Значение." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.Value = valueDecimal;
                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Значение\"). В режиме \"Добавить\" Значение не может быть пустым.\" +." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (changedSapNdoOUTDTO.Value < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Значение\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                rowNumber++;
                                continue;
                            }


                            if (String.IsNullOrEmpty(sapGoneVarString))
                            {
                                changedSapNdoOUTDTO.SapGone = false;
                            }
                            else
                            {
                                changedSapNdoOUTDTO.SapGone = sapGoneVarString.ToUpper().Equals("ДА") ? true : false;
                            }

                            DateTime sapGoneTimeVarDateTime;

                            if (!String.IsNullOrEmpty(sapGoneTimeVarString))
                            {
                                try
                                {
                                    sapGoneTimeVarDateTime = DateTime.Parse(sapGoneTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Когда Sap забрал\"). Не удалось получить Когда Sap забрал." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedSapNdoOUTDTO.SapGoneTime = sapGoneTimeVarDateTime;
                                changedSapNdoOUTDTO.SapGoneTime = changedSapNdoOUTDTO.SapGoneTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedSapNdoOUTDTO.SapGoneTime;
                                changedSapNdoOUTDTO.SapGoneTime = changedSapNdoOUTDTO.SapGoneTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedSapNdoOUTDTO.SapGoneTime;
                            }
                            else
                                changedSapNdoOUTDTO.SapGoneTime = null;

                            if (!((changedSapNdoOUTDTO.SapGone == true && changedSapNdoOUTDTO.SapGoneTime != null)
                                || (changedSapNdoOUTDTO.SapGone != true && changedSapNdoOUTDTO.SapGoneTime == null)))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 8, 9 (\"Sap забрал\", \"Когда Sap забрал\"). Если \"Sap забрал\" равно \"Да\" время \"Когда Sap забрал\" должно быть заполнено."
                                    + " Если \"Sap забрал\" равно \"Нет\" или пусто, то время \"Когда Sap забрал\" должно быть пустым."
                                    + " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 2, 8, 9 }, resultString);
                                rowNumber++;
                                continue;

                            }

                            SapNdoOUTDTO? newSapNdoOUTDTO = await _sapNdoOUTRepository.Create(changedSapNdoOUTDTO);

                            await _logEventRepository.ToLog<SapNdoOUTDTO>(oldObject: null, newObject: newSapNdoOUTDTO, "Добавление записи в витрину SAP НДО-выход SapNdoOUT", "Запись: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Добавлена запись с ИД " + newSapNdoOUTDTO.Id.ToString()
                                + ". " + tagNameWarningString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            if (String.IsNullOrEmpty(tagNameWarningString))
                            {
                                worksheet.Cell(rowNumber, 1).Value = "OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 3).Value = newSapNdoOUTDTO.Id.ToString();
                                worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Green;
                                worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 1).Value = "!!! OK";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 3).Value = newSapNdoOUTDTO.Id.ToString();
                                worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 5).Style.Font.FontColor = XLColor.YellowGreen;
                                worksheet.Cell(rowNumber, 5).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.YellowGreen;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "УДАЛИТЬ":
                        {
                            if (String.IsNullOrEmpty(idVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Для действия \"Удалить\" ИД записи единственное необходимое поле. Не может быть пустым." +
                                        " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (idVarInt64 <= 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). ИД записи должен быть положительным числом." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            foundSapNdoOUTDTO = await _sapNdoOUTRepository.GetById(idVarInt64);

                            if (foundSapNdoOUTDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Запись с ИД " + idVarInt64.ToString() + " не найдена в витрине Sap НДО-выход." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            Guid deletedUserId = (await _authorizationRepository.GetCurrentUserDTO()).Id;

                            string mesNdoListString = "";
                            var mesNdoStocksList = await _mesNdoStocksRepository.GetBySapNdoOutIdList(idVarInt64);

                            if (mesNdoStocksList != null)
                                if (mesNdoStocksList.Count() > 0)
                                {
                                    foreach (var mesNdoStocksItem in mesNdoStocksList)
                                    {
                                        await _logEventRepository.AddRecord("Удаление записи из Архива данных НДО MesNdoStocks", deletedUserId, "", "", false, "Удаление записи с ИД " + mesNdoStocksItem.Id.ToString() + ": " + mesNdoStocksItem.ToString());
                                        await _mesNdoStocksRepository.Delete(mesNdoStocksItem.Id);
                                        mesNdoListString = mesNdoListString + " " + mesNdoStocksItem.Id.ToString() + ", ";
                                    }
                                }



                            await _logEventRepository.AddRecord("Удаление записи из витрины SAP НДО-выход SapNdoOUT", deletedUserId, "", "", false, "Удаление записи с ИД " + foundSapNdoOUTDTO.Id.ToString() + ": " + foundSapNdoOUTDTO.ToString());
                            await _sapNdoOUTRepository.Delete(foundSapNdoOUTDTO.Id);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Удалена запись с ИД " + foundSapNdoOUTDTO.Id.ToString() + ".";
                            if (!String.IsNullOrEmpty(mesNdoListString))
                                resultString = resultString + ". Также удалены связанные записи в Архиве данных НДО (MesNdoStocks) с ИД: " + mesNdoListString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);

                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            haveErrors = true;
                            idVarInt64 = 0;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 2 (\"Действие\"). Не предусмотренное значение действия = " + actionVarString + ". Для витрины НДО-выход допустимы действия: \"Добавить\", \"Изменить\", \"Удалить\". Изменения не применялись.";
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

        public async Task<bool> MesNdoStocksExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в Архив данных НДО");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;
            int resultColumnNumber = 13;

            bool isEmptyString = false;

            UserDTO currentUserDTO = await _authorizationRepository.GetCurrentUserDTO();

            while (isEmptyString == false)
            {
                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string actionVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string idVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string addTimeVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string addUserIdVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string addUserNameVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string mesParamCodeVarString = rowVar.Cell(7).CachedValue.ToString().Trim();
                string valueTimeVarString = rowVar.Cell(8).CachedValue.ToString().Trim();
                string valueVarString = rowVar.Cell(9).CachedValue.ToString().Trim();
                string valueDifferenceVarString = rowVar.Cell(10).CachedValue.ToString().Trim();
                string reportGuidVarString = rowVar.Cell(11).CachedValue.ToString().Trim();
                string sapNdoOutIdVarString = rowVar.Cell(12).CachedValue.ToString().Trim();

                string resultString = "";
                Int64 idVarInt64 = 0;


                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(addTimeVarString) && String.IsNullOrEmpty(addUserIdVarString)
                        && String.IsNullOrEmpty(addUserNameVarString) && String.IsNullOrEmpty(mesParamCodeVarString) && String.IsNullOrEmpty(valueTimeVarString)
                        && String.IsNullOrEmpty(valueVarString) && String.IsNullOrEmpty(valueDifferenceVarString) && String.IsNullOrEmpty(reportGuidVarString)
                        && String.IsNullOrEmpty(sapNdoOutIdVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                MesNdoStocksDTO? foundMesNdoStocksDTO = null;
                MesNdoStocksDTO changedMesNdoStocksDTO = new MesNdoStocksDTO();

                if (!String.IsNullOrEmpty(idVarString))
                {
                    try
                    {
                        idVarInt64 = Int64.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Не удалось получить ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                        rowNumber++;
                        continue;
                    }
                }

                string tagNameWarningString = "";

                switch (actionVarString.Trim().ToUpper())
                {
                    case "ИЗМЕНИТЬ":
                        {
                            if (idVarInt64 <= 0)
                            {
                                haveErrors = true;
                                idVarInt64 = 0;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). ИД записи должен быть положительным числом." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            foundMesNdoStocksDTO = await _mesNdoStocksRepository.GetById(idVarInt64);
                            if (foundMesNdoStocksDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Не найдена запись с ИД: " + idVarInt64.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            changedMesNdoStocksDTO.Id = idVarInt64;

                            DateTime addTimeVarDateTime;

                            if (!String.IsNullOrEmpty(addTimeVarString))
                            {
                                try
                                {
                                    addTimeVarDateTime = DateTime.Parse(addTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Время добавления записи\"). Не удалось получить Время добавления записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.AddTime = addTimeVarDateTime;
                            }
                            else
                                changedMesNdoStocksDTO.AddTime = foundMesNdoStocksDTO.AddTime;

                            changedMesNdoStocksDTO.AddTime = changedMesNdoStocksDTO.AddTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesNdoStocksDTO.AddTime;
                            changedMesNdoStocksDTO.AddTime = changedMesNdoStocksDTO.AddTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesNdoStocksDTO.AddTime;

                            changedMesNdoStocksDTO.AddUserId = Guid.Empty;
                            changedMesNdoStocksDTO.AddUserDTOFK = new UserDTO();

                            Guid addUserIdVarGuid = Guid.Empty;

                            if (!String.IsNullOrEmpty(addUserIdVarString))
                            {
                                try
                                {
                                    addUserIdVarGuid = Guid.Parse(addUserIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ИД добавившего запись пользователя\"). Не удалось получить ИД добавившего запись пользователя." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                var userObjectByGuid = await _userRepository.Get(addUserIdVarGuid);
                                if (userObjectByGuid == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ИД добавившего запись пользователя\"). Не найден пользователь с ИД " + addUserIdVarGuid.ToString() + " в Справочнике пользователей." +
                                        " Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.AddUserId = userObjectByGuid.Id;
                                changedMesNdoStocksDTO.AddUserDTOFK = userObjectByGuid;
                            }

                            if (changedMesNdoStocksDTO.AddUserId == Guid.Empty)
                                if (!String.IsNullOrEmpty(addUserNameVarString))
                                {
                                    var userObjectList = await _userRepository.GetListByUserName(addUserNameVarString);
                                    if (userObjectList != null)
                                    {
                                        if (userObjectList.Count() > 1)
                                        {
                                            haveErrors = true;
                                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). Найдено больше одного пользователя с ФИО " + addUserNameVarString + "." +
                                                " Изменения не применялись.";
                                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                            rowNumber++;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        haveErrors = true;
                                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). В Справочнике пользователей не найден пользователь с ФИО " + addUserNameVarString + "." +
                                            " Изменения не применялись.";
                                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                        rowNumber++;
                                        continue;
                                    }
                                    changedMesNdoStocksDTO.AddUserId = userObjectList.FirstOrDefault().Id;
                                    changedMesNdoStocksDTO.AddUserDTOFK = userObjectList.FirstOrDefault();
                                }

                            if (changedMesNdoStocksDTO.AddUserId == Guid.Empty)
                            {
                                changedMesNdoStocksDTO.AddUserId = foundMesNdoStocksDTO.AddUserId;
                                changedMesNdoStocksDTO.AddUserDTOFK = foundMesNdoStocksDTO.AddUserDTOFK;
                            }

                            if (!String.IsNullOrEmpty(mesParamCodeVarString))
                            {
                                MesParamDTO? objectForCheckMesParamCode = await _mesParamRepository.GetByCode(mesParamCodeVarString);
                                if (objectForCheckMesParamCode == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Код тэга\"). В Справочнике тэгов СИР не найден тэг с кодом " + objectForCheckMesParamCode + "." +
                                        " Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                    rowNumber++;
                                    continue;

                                }
                                changedMesNdoStocksDTO.MesParamId = objectForCheckMesParamCode.Id;
                                changedMesNdoStocksDTO.MesParamDTOFK = objectForCheckMesParamCode;
                            }
                            else
                            {
                                changedMesNdoStocksDTO.MesParamId = foundMesNdoStocksDTO.MesParamId;
                                changedMesNdoStocksDTO.MesParamDTOFK = foundMesNdoStocksDTO.MesParamDTOFK;
                            }

                            DateTime valueTimeVarDateTime;
                            if (!String.IsNullOrEmpty(valueTimeVarString))
                            {
                                try
                                {
                                    valueTimeVarDateTime = DateTime.Parse(valueTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время значения\"). Не удалось получить Время значения." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.ValueTime = valueTimeVarDateTime;
                            }
                            else
                                changedMesNdoStocksDTO.ValueTime = foundMesNdoStocksDTO.ValueTime;

                            changedMesNdoStocksDTO.ValueTime = changedMesNdoStocksDTO.ValueTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesNdoStocksDTO.ValueTime;
                            changedMesNdoStocksDTO.ValueTime = changedMesNdoStocksDTO.ValueTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesNdoStocksDTO.ValueTime;

                            decimal valueDecimal;
                            if (!String.IsNullOrEmpty(valueVarString))
                            {
                                try
                                {
                                    valueDecimal = decimal.Parse(valueVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Не удалось получить Значение." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.Value = valueDecimal;
                            }
                            else
                                changedMesNdoStocksDTO.Value = 0;

                            if (changedMesNdoStocksDTO.Value < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            decimal valueDifferenceDecimal;
                            if (!String.IsNullOrEmpty(valueDifferenceVarString))
                            {
                                try
                                {
                                    valueDifferenceDecimal = decimal.Parse(valueDifferenceVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 10 (\"Разность\"). Не удалось получить Разность." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 10 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.ValueDifference = valueDifferenceDecimal;
                            }
                            else
                                changedMesNdoStocksDTO.ValueDifference = 0;

                            if (changedMesNdoStocksDTO.ValueDifference < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 10 (\"Разность\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 10 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            Guid reportGuidVarGuid = Guid.Empty;
                            if (!String.IsNullOrEmpty(reportGuidVarString))
                            {
                                try
                                {
                                    reportGuidVarGuid = Guid.Parse(reportGuidVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД Экземпляра отчёта\"). Не удалось получить ИД Экземпляра отчёта." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }

                            if (reportGuidVarGuid == Guid.Empty)
                            {
                                changedMesNdoStocksDTO.ReportGuid = null;
                                changedMesNdoStocksDTO.ReportEntityDTOFK = null;
                            }
                            else
                            {
                                ReportEntityDTO? objectForCheckReportEntityDTO = await _reportEntityRepository.GetById(reportGuidVarGuid);
                                if (objectForCheckReportEntityDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД Экземпляра отчёта\"). Не удалось найти экземпляр отчёта." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesNdoStocksDTO.ReportGuid = objectForCheckReportEntityDTO.Id;
                                    changedMesNdoStocksDTO.ReportEntityDTOFK = objectForCheckReportEntityDTO;
                                }
                            }

                            Int64 sapNdoOutIdVarInt64 = 0;
                            if (!String.IsNullOrEmpty(sapNdoOutIdVarString))
                            {
                                try
                                {
                                    sapNdoOutIdVarInt64 = Int64.Parse(sapNdoOutIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 12 (\"ИД записи в Витрине SAP \"НДО-выход\"\"). Не удалось получить ИД записи в Витрине SAP \"НДО-выход\"." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 12 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.SapNdoOutId = sapNdoOutIdVarInt64;
                            }

                            if (sapNdoOutIdVarInt64 == 0)
                            {
                                changedMesNdoStocksDTO.SapNdoOutId = null;
                                changedMesNdoStocksDTO.SapNdoOUTDTOFK = null;
                            }
                            else
                            {
                                SapNdoOUTDTO? objectForCheckSapNdoOUTDTO = await _sapNdoOUTRepository.GetById(sapNdoOutIdVarInt64);
                                if (objectForCheckSapNdoOUTDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 12 (\"ИД записи в Витрине SAP \"НДО-выход\"\"). Не удалось найти запись с ИД "
                                        + sapNdoOutIdVarInt64.ToString() + " в Витрине SAP \"НДО-выход\". Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 12 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesNdoStocksDTO.SapNdoOutId = objectForCheckSapNdoOUTDTO.Id;
                                    changedMesNdoStocksDTO.SapNdoOUTDTOFK = objectForCheckSapNdoOUTDTO;
                                }
                            }

                            await _mesNdoStocksRepository.Update(changedMesNdoStocksDTO);
                            await _logEventRepository.ToLog<MesNdoStocksDTO>(oldObject: foundMesNdoStocksDTO, newObject: changedMesNdoStocksDTO
                                , "Изменение записи в Архиве данных НДО MesNdoStocks", "Запись: ", _authorizationRepository);


                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;

                            worksheet.Cell(rowNumber, 1).Value = "ОК";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);

                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "ДОБАВИТЬ":
                        {
                            if (!String.IsNullOrEmpty(idVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Для действия \"Добавить\" ИД записи должен быть пустым. Сформируется автоматически." +
                                        " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            DateTime addTimeVarDateTime;

                            if (!String.IsNullOrEmpty(addTimeVarString))
                            {
                                try
                                {
                                    addTimeVarDateTime = DateTime.Parse(addTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Время добавления записи\"). Не удалось получить Время добавления записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.AddTime = addTimeVarDateTime;
                                changedMesNdoStocksDTO.AddTime = changedMesNdoStocksDTO.AddTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesNdoStocksDTO.AddTime;
                                changedMesNdoStocksDTO.AddTime = changedMesNdoStocksDTO.AddTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesNdoStocksDTO.AddTime;
                            }

                            Guid addUserIdVarGuid = Guid.Empty;

                            changedMesNdoStocksDTO.AddUserId = Guid.Empty;
                            changedMesNdoStocksDTO.AddUserDTOFK = new UserDTO();

                            if (!String.IsNullOrEmpty(addUserIdVarString))
                            {
                                try
                                {
                                    addUserIdVarGuid = Guid.Parse(addUserIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ИД добавившего запись пользователя\"). Не удалось получить ИД добавившего запись пользователя." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                if (addUserIdVarGuid != Guid.Empty)
                                    if (!String.IsNullOrEmpty(addUserNameVarString))
                                    {
                                        var userObjectByGuid = await _userRepository.Get(addUserIdVarGuid);

                                        if (userObjectByGuid != null)
                                        {
                                            changedMesNdoStocksDTO.AddUserId = userObjectByGuid.Id;
                                            changedMesNdoStocksDTO.AddUserDTOFK = userObjectByGuid;
                                        }
                                    }
                            }


                            if (changedMesNdoStocksDTO.AddUserId == Guid.Empty)
                                if (!String.IsNullOrEmpty(addUserNameVarString))
                                {
                                    var userObjectList = await _userRepository.GetListByUserName(addUserNameVarString);
                                    if (userObjectList != null)
                                    {
                                        if (userObjectList.Count() > 1)
                                        {
                                            haveErrors = true;
                                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). Найдено больше одного пользователя с ФИО " + addUserNameVarString + "." +
                                                " Изменения не применялись.";
                                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                            rowNumber++;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        haveErrors = true;
                                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). В Справочнике пользователей не найден пользователь с ФИО " + addUserNameVarString + "." +
                                            " Изменения не применялись.";
                                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                        rowNumber++;
                                        continue;
                                    }
                                    if (userObjectList.Count() > 0)
                                    {
                                        changedMesNdoStocksDTO.AddUserId = userObjectList.FirstOrDefault().Id;
                                        changedMesNdoStocksDTO.AddUserDTOFK = userObjectList.FirstOrDefault();
                                    }
                                }

                            if (changedMesNdoStocksDTO.AddUserId == Guid.Empty)
                            {
                                changedMesNdoStocksDTO.AddUserId = currentUserDTO.Id;
                                changedMesNdoStocksDTO.AddUserDTOFK = currentUserDTO;
                            }

                            if (!String.IsNullOrEmpty(mesParamCodeVarString))
                            {
                                MesParamDTO? objectForCheckMesParamCode = await _mesParamRepository.GetByCode(mesParamCodeVarString);
                                if (objectForCheckMesParamCode == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Код тэга\"). В Справочнике тэгов СИР не найден не найден тэг с кодом " + mesParamCodeVarString + "." +
                                        " Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.MesParamId = objectForCheckMesParamCode.Id;
                                changedMesNdoStocksDTO.MesParamDTOFK = objectForCheckMesParamCode;
                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Код тэга\"). В режиме \"Добавить\" Код тэга не может быть пустым." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            DateTime valueTimeVarDateTime;
                            if (!String.IsNullOrEmpty(valueTimeVarString))
                            {
                                try
                                {
                                    valueTimeVarDateTime = DateTime.Parse(valueTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    idVarInt64 = 0;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время значения\"). Не удалось получить Время значения." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.ValueTime = valueTimeVarDateTime;
                                changedMesNdoStocksDTO.ValueTime = changedMesNdoStocksDTO.ValueTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesNdoStocksDTO.ValueTime;
                                changedMesNdoStocksDTO.ValueTime = changedMesNdoStocksDTO.ValueTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesNdoStocksDTO.ValueTime;
                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время значения\"). В режиме \"Добавить\" Время значения не может быть пустым." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                rowNumber++;
                                continue;
                            }


                            decimal valueDecimal;
                            if (!String.IsNullOrEmpty(valueVarString))
                            {
                                try
                                {
                                    valueDecimal = decimal.Parse(valueVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Не удалось получить Значение." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.Value = valueDecimal;
                            }
                            else
                            {
                                changedMesNdoStocksDTO.Value = 0;
                            }

                            if (changedMesNdoStocksDTO.Value < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            decimal valueDifferenceDecimal;
                            if (!String.IsNullOrEmpty(valueDifferenceVarString))
                            {
                                try
                                {
                                    valueDifferenceDecimal = decimal.Parse(valueDifferenceVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 10 (\"Разность\"). Не удалось получить Разность." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 10 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.ValueDifference = valueDifferenceDecimal;
                            }
                            else
                            {
                                changedMesNdoStocksDTO.ValueDifference = 0;
                            }

                            if (changedMesNdoStocksDTO.ValueDifference < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 10 (\"Разность\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 10 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            Guid reportGuidVarGuid = Guid.Empty;
                            if (!String.IsNullOrEmpty(reportGuidVarString))
                            {
                                try
                                {
                                    reportGuidVarGuid = Guid.Parse(reportGuidVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД Экземпляра отчёта\"). Не удалось получить ИД Экземпляра отчёта." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }

                            if (reportGuidVarGuid == Guid.Empty)
                            {
                                changedMesNdoStocksDTO.ReportGuid = null;
                                changedMesNdoStocksDTO.ReportEntityDTOFK = null;
                            }
                            else
                            {
                                ReportEntityDTO? objectForCheckReportEntityDTO = await _reportEntityRepository.GetById(reportGuidVarGuid);
                                if (objectForCheckReportEntityDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД Экземпляра отчёта\"). Не удалось найти экземпляр отчёта." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesNdoStocksDTO.ReportGuid = objectForCheckReportEntityDTO.Id;
                                    changedMesNdoStocksDTO.ReportEntityDTOFK = objectForCheckReportEntityDTO;
                                }
                            }

                            Int64 sapNdoOutIdVarInt64 = 0;
                            if (!String.IsNullOrEmpty(sapNdoOutIdVarString))
                            {
                                try
                                {
                                    sapNdoOutIdVarInt64 = Int64.Parse(sapNdoOutIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 12 (\"ИД записи в Витрине SAP \"НДО-выход\"\"). Не удалось получить ИД записи в Витрине SAP \"НДО-выход\"." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 12 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesNdoStocksDTO.ReportGuid = reportGuidVarGuid;
                            }

                            if (sapNdoOutIdVarInt64 == 0)
                            {
                                changedMesNdoStocksDTO.SapNdoOutId = null;
                                changedMesNdoStocksDTO.SapNdoOUTDTOFK = null;
                            }
                            else
                            {
                                SapNdoOUTDTO? objectForCheckSapNdoOUTDTO = await _sapNdoOUTRepository.GetById(sapNdoOutIdVarInt64);
                                if (objectForCheckSapNdoOUTDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 12 (\"ИД записи в Витрине SAP \"НДО-выход\"\"). Не удалось найти запись с ИД "
                                        + sapNdoOutIdVarInt64.ToString() + " в Витрине SAP \"НДО-выход\". Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 12 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesNdoStocksDTO.SapNdoOutId = objectForCheckSapNdoOUTDTO.Id;
                                    changedMesNdoStocksDTO.SapNdoOUTDTOFK = objectForCheckSapNdoOUTDTO;
                                }

                            }
                            MesNdoStocksDTO? newMesNdoStocksDTO = await _mesNdoStocksRepository.Create(changedMesNdoStocksDTO);
                            await _logEventRepository.ToLog<MesNdoStocksDTO>(oldObject: null, newObject: newMesNdoStocksDTO, "Добавление записи в Архив данных НДО MesNdoStocks", "Запись: ", _authorizationRepository);
                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Добавлена запись с ИД " + newMesNdoStocksDTO.Id.ToString();
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 3).Value = newMesNdoStocksDTO.Id.ToString();
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "УДАЛИТЬ":
                        {
                            if (String.IsNullOrEmpty(idVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Для действия \"Удалить\" ИД записи единственное необходимое поле. Не может быть пустым." +
                                        " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (idVarInt64 <= 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). ИД записи должен быть положительным числом." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }
                            foundMesNdoStocksDTO = await _mesNdoStocksRepository.GetById(idVarInt64);
                            if (foundMesNdoStocksDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Запись с ИД " + idVarInt64.ToString() + " не найдена в Архиве данных НДО." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            await _logEventRepository.AddRecord("Удаление записи из Архива данных НДО MesNdoStocks", currentUserDTO.Id, "", "", false, "Удаление записи с ИД " + foundMesNdoStocksDTO.Id.ToString() + ": " + foundMesNdoStocksDTO.ToString());
                            await _mesNdoStocksRepository.Delete(foundMesNdoStocksDTO.Id);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Удалена запись с ИД " + foundMesNdoStocksDTO.Id.ToString() + ".";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);

                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            haveErrors = true;
                            idVarInt64 = 0;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 2 (\"Действие\"). Не предусмотренное значение действия = " + actionVarString + ". Для витрины НДО-выход допустимы действия: \"Добавить\", \"Изменить\", \"Удалить\". Изменения не применялись.";
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
            int resultColumnNumber = 11;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
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
            int resultColumnNumber = 7;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
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


        public async Task<bool> MesMovementsExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в Архив данных");
            await loadFromExcelPage.RefreshSate();

            int rowNumber = 9;
            int resultColumnNumber = 19;

            bool isEmptyString = false;

            UserDTO currentUserDTO = await _authorizationRepository.GetCurrentUserDTO();

            while (isEmptyString == false)
            {

                worksheet.Cell(rowNumber, 1).Value = "";
                worksheet.Cell(rowNumber, 2).Value = "";
                worksheet.Cell(rowNumber, resultColumnNumber).Value = "";
                worksheet.Row(rowNumber).Style.Font.SetBold(false);
                worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Black;

                var rowVar = worksheet.Row(rowNumber);

                string actionVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string idVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string addTimeVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string addUserIdVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string addUserNameVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string mesParamCodeVarString = rowVar.Cell(7).CachedValue.ToString().Trim();
                string valueTimeVarString = rowVar.Cell(8).CachedValue.ToString().Trim();
                string valueVarString = rowVar.Cell(9).CachedValue.ToString().Trim();
                string sapMovementOutIdVarString = rowVar.Cell(10).CachedValue.ToString().Trim();
                string sapMovementInIdVarString = rowVar.Cell(11).CachedValue.ToString().Trim();
                string dataSourceNameVarString = rowVar.Cell(12).CachedValue.ToString().Trim();
                string dataTypeNameVarString = rowVar.Cell(13).CachedValue.ToString().Trim();
                string reportGuidVarString = rowVar.Cell(14).CachedValue.ToString().Trim();
                string mesGoneVarString = rowVar.Cell(15).CachedValue.ToString().Trim();
                string mesGoneTimeVarString = rowVar.Cell(16).CachedValue.ToString().Trim();
                string needWriteToSapVarString = rowVar.Cell(17).CachedValue.ToString().Trim();
                string previousRecordIdVarString = rowVar.Cell(18).CachedValue.ToString().Trim();

                string resultString = "";
                Guid idVarGuid = Guid.Empty;

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(addTimeVarString) && String.IsNullOrEmpty(addUserIdVarString)
                        && String.IsNullOrEmpty(addUserNameVarString) && String.IsNullOrEmpty(mesParamCodeVarString) && String.IsNullOrEmpty(valueTimeVarString)
                        && String.IsNullOrEmpty(valueVarString) && String.IsNullOrEmpty(sapMovementOutIdVarString) && String.IsNullOrEmpty(sapMovementInIdVarString)
                        && String.IsNullOrEmpty(dataSourceNameVarString) && String.IsNullOrEmpty(dataTypeNameVarString) && String.IsNullOrEmpty(reportGuidVarString)
                        && String.IsNullOrEmpty(mesGoneVarString) && String.IsNullOrEmpty(mesGoneTimeVarString) && String.IsNullOrEmpty(needWriteToSapVarString)
                        && String.IsNullOrEmpty(previousRecordIdVarString))
                {
                    isEmptyString = true;
                    continue;
                }

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

                MesMovementsDTO? foundMesMovementsDTO = null;
                MesMovementsDTO changedMesMovementsDTO = new MesMovementsDTO();

                if (!String.IsNullOrEmpty(idVarString))
                {
                    try
                    {
                        idVarGuid = Guid.Parse(idVarString);
                    }
                    catch (Exception ex)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Не удалось получить ИД записи." +
                            " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                        rowNumber++;
                        continue;
                    }
                }

                string mesMovementsIdWarningString = "";

                switch (actionVarString.Trim().ToUpper())
                {
                    case "ИЗМЕНИТЬ":
                        {

                            foundMesMovementsDTO = await _mesMovementsRepository.GetById(idVarGuid);
                            if (foundMesMovementsDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Не найдена запись с ИД: " + idVarGuid.ToString() + ". Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            changedMesMovementsDTO.Id = idVarGuid;

                            DateTime addTimeVarDateTime;

                            if (!String.IsNullOrEmpty(addTimeVarString))
                            {
                                try
                                {
                                    addTimeVarDateTime = DateTime.Parse(addTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Время добавления записи\"). Не удалось получить Время добавления записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.AddTime = addTimeVarDateTime;
                            }
                            else
                                changedMesMovementsDTO.AddTime = foundMesMovementsDTO.AddTime;

                            changedMesMovementsDTO.AddTime = changedMesMovementsDTO.AddTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesMovementsDTO.AddTime;
                            changedMesMovementsDTO.AddTime = changedMesMovementsDTO.AddTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesMovementsDTO.AddTime;

                            changedMesMovementsDTO.AddUserId = Guid.Empty;
                            changedMesMovementsDTO.AddUserDTOFK = new UserDTO();

                            Guid addUserIdVarGuid = Guid.Empty;

                            if (!String.IsNullOrEmpty(addUserIdVarString))
                            {
                                try
                                {
                                    addUserIdVarGuid = Guid.Parse(addUserIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ИД добавившего запись пользователя\"). Не удалось получить ИД добавившего запись пользователя." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                var userObjectByGuid = await _userRepository.Get(addUserIdVarGuid);
                                if (userObjectByGuid == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ИД добавившего запись пользователя\"). Не найден пользователь с ИД " + addUserIdVarGuid.ToString() + " в Справочнике пользователей." +
                                        " Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.AddUserId = userObjectByGuid.Id;
                                changedMesMovementsDTO.AddUserDTOFK = userObjectByGuid;
                            }

                            if (changedMesMovementsDTO.AddUserId == Guid.Empty)
                                if (!String.IsNullOrEmpty(addUserNameVarString))
                                {
                                    var userObjectList = await _userRepository.GetListByUserName(addUserNameVarString);
                                    if (userObjectList != null)
                                    {
                                        if (userObjectList.Count() > 1)
                                        {
                                            haveErrors = true;
                                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). Найдено больше одного пользователя с ФИО " + addUserNameVarString + "." +
                                                " Изменения не применялись.";
                                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                            rowNumber++;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        haveErrors = true;
                                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). В Справочнике пользователей не найден пользователь с ФИО " + addUserNameVarString + "." +
                                            " Изменения не применялись.";
                                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                        rowNumber++;
                                        continue;
                                    }
                                    changedMesMovementsDTO.AddUserId = userObjectList.FirstOrDefault().Id;
                                    changedMesMovementsDTO.AddUserDTOFK = userObjectList.FirstOrDefault();
                                }

                            if (changedMesMovementsDTO.AddUserId == Guid.Empty)
                            {
                                changedMesMovementsDTO.AddUserId = foundMesMovementsDTO.AddUserId;
                                changedMesMovementsDTO.AddUserDTOFK = foundMesMovementsDTO.AddUserDTOFK;
                            }

                            if (!String.IsNullOrEmpty(mesParamCodeVarString))
                            {
                                MesParamDTO? objectForCheckMesParamCode = await _mesParamRepository.GetByCode(mesParamCodeVarString);
                                if (objectForCheckMesParamCode == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Код тэга\"). В Справочнике тэгов СИР не найден тэг с кодом " + objectForCheckMesParamCode + "." +
                                        " Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                    rowNumber++;
                                    continue;

                                }
                                changedMesMovementsDTO.MesParamId = objectForCheckMesParamCode.Id;
                                changedMesMovementsDTO.MesParamDTOFK = objectForCheckMesParamCode;
                            }
                            else
                            {
                                changedMesMovementsDTO.MesParamId = foundMesMovementsDTO.MesParamId;
                                changedMesMovementsDTO.MesParamDTOFK = foundMesMovementsDTO.MesParamDTOFK;
                            }

                            DateTime valueTimeVarDateTime;
                            if (!String.IsNullOrEmpty(valueTimeVarString))
                            {
                                try
                                {
                                    valueTimeVarDateTime = DateTime.Parse(valueTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время значения\"). Не удалось получить Время значения." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.ValueTime = valueTimeVarDateTime;
                            }
                            else
                                changedMesMovementsDTO.ValueTime = foundMesMovementsDTO.ValueTime;

                            changedMesMovementsDTO.ValueTime = changedMesMovementsDTO.ValueTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesMovementsDTO.ValueTime;
                            changedMesMovementsDTO.ValueTime = changedMesMovementsDTO.ValueTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesMovementsDTO.ValueTime;

                            decimal valueDecimal;
                            if (!String.IsNullOrEmpty(valueVarString))
                            {
                                try
                                {
                                    valueDecimal = decimal.Parse(valueVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Не удалось получить Значение." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.Value = valueDecimal;
                            }
                            else
                                changedMesMovementsDTO.Value = 0;

                            if (changedMesMovementsDTO.Value < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (String.IsNullOrEmpty(sapMovementInIdVarString))
                            {
                                changedMesMovementsDTO.SapMovementInId = null;
                                changedMesMovementsDTO.SapMovementsINDTOFK = null;
                            }
                            else
                            {
                                SapMovementsINDTO? objectForCheckSapMovementsINDTO = await _sapMovementsINRepository.GetById(sapMovementInIdVarString);
                                if (objectForCheckSapMovementsINDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 10 (\"ИД записи во входной витрине SapMovementsIN\"). Не удалось найти запись в витрине SAP Движения-вход." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 10 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.SapMovementInId = objectForCheckSapMovementsINDTO.ErpId;
                                    changedMesMovementsDTO.SapMovementsINDTOFK = objectForCheckSapMovementsINDTO;
                                }
                            }

                            if (String.IsNullOrEmpty(sapMovementOutIdVarString))
                            {
                                changedMesMovementsDTO.SapMovementOutId = null;
                                changedMesMovementsDTO.SapMovementsOUTDTOFK = null;
                            }
                            else
                            {
                                Guid sapMovementOutIdVarGuid = Guid.Empty;
                                try
                                {
                                    sapMovementOutIdVarGuid = Guid.Parse(reportGuidVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД записи в выходной витрине SapMovementsOUT\"). Не удалось получить ИД записи в выходной витрине SapMovementsOUT." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                SapMovementsOUTDTO? objectForCheckSapMovementsOUTDTO = await _sapMovementsOUTRepository.GetById(sapMovementOutIdVarGuid);
                                if (objectForCheckSapMovementsOUTDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД записи в выходной витрине SapMovementsOUT\"). Не удалось найти запись в витрине SAP Движения-выход." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.SapMovementOutId = objectForCheckSapMovementsOUTDTO.Id;
                                    changedMesMovementsDTO.SapMovementsOUTDTOFK = objectForCheckSapMovementsOUTDTO;
                                }
                            }

                            if (String.IsNullOrEmpty(dataSourceNameVarString))
                            {
                                changedMesMovementsDTO.DataSourceId = null;
                                changedMesMovementsDTO.DataSourceDTOFK = null;
                            }
                            else
                            {
                                DataSourceDTO? objectForCheckDataSourceDTO = await _dataSourceRepository.GetByName(dataSourceNameVarString);
                                if (objectForCheckDataSourceDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 13 (\"Источник данных\"). Не удалось найти Источник данных с таким наименованием." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 13 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.DataSourceId = objectForCheckDataSourceDTO.Id;
                                    changedMesMovementsDTO.DataSourceDTOFK = objectForCheckDataSourceDTO;
                                }
                            }

                            if (String.IsNullOrEmpty(dataTypeNameVarString))
                            {
                                changedMesMovementsDTO.DataTypeId = null;
                                changedMesMovementsDTO.DataTypeDTOFK = null;
                            }
                            else
                            {
                                DataTypeDTO? objectForCheckDataTypeDTO = await _dataTypeRepository.GetByName(dataTypeNameVarString);
                                if (objectForCheckDataTypeDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 14 (\"Тип данных\"). Не удалось найти Тип данных с таким наименованием." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 14 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.DataTypeId = objectForCheckDataTypeDTO.Id;
                                    changedMesMovementsDTO.DataTypeDTOFK = objectForCheckDataTypeDTO;
                                }
                            }

                            Guid reportGuidVarGuid = Guid.Empty;
                            if (!String.IsNullOrEmpty(reportGuidVarString))
                            {
                                try
                                {
                                    reportGuidVarGuid = Guid.Parse(reportGuidVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 15 (\"ИД Экземпляра отчёта\"). Не удалось получить ИД Экземпляра отчёта." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 15 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }

                            if (reportGuidVarGuid == Guid.Empty)
                            {
                                changedMesMovementsDTO.ReportGuid = null;
                                changedMesMovementsDTO.ReportEntityDTOFK = null;
                            }
                            else
                            {
                                ReportEntityDTO? objectForCheckReportEntityDTO = await _reportEntityRepository.GetById(reportGuidVarGuid);
                                if (objectForCheckReportEntityDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 15 (\"ИД Экземпляра отчёта\"). Не удалось найти экземпляр отчёта." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 15 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.ReportGuid = objectForCheckReportEntityDTO.Id;
                                    changedMesMovementsDTO.ReportEntityDTOFK = objectForCheckReportEntityDTO;
                                }
                            }

                            if (String.IsNullOrEmpty(mesGoneVarString))
                                changedMesMovementsDTO.MesGone = false;
                            else
                                changedMesMovementsDTO.MesGone = mesGoneVarString.ToUpper().Equals("ДА") ? true : false;

                            DateTime mesGoneTimeVarDateTime;

                            if (!String.IsNullOrEmpty(mesGoneTimeVarString))
                            {
                                try
                                {
                                    mesGoneTimeVarDateTime = DateTime.Parse(mesGoneTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 16 (\"Время MES забрал\"). Не удалось получить Время MES забрал." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 16 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.MesGoneTime = mesGoneTimeVarDateTime;
                                changedMesMovementsDTO.MesGoneTime = changedMesMovementsDTO.MesGoneTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesMovementsDTO.MesGoneTime;
                                changedMesMovementsDTO.MesGoneTime = changedMesMovementsDTO.MesGoneTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesMovementsDTO.MesGoneTime;
                            }
                            else
                                changedMesMovementsDTO.MesGoneTime = null;

                            if (!((changedMesMovementsDTO.MesGone == true && changedMesMovementsDTO.MesGoneTime != null)
                                || (changedMesMovementsDTO.MesGone != true && changedMesMovementsDTO.MesGoneTime == null)))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 15, 16 (\"Mes забрал\", \"Время MES забрал\"). Если \"Mes забрал\" равно \"Да\", то \"Время MES забрал\" должно быть заполнено."
                                    + " Если \"Mes забрал\" равно \"Нет\" или пусто, то \"Время MES забрал\" должно быть пустым."
                                    + " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 2, 15, 16 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (String.IsNullOrEmpty(needWriteToSapVarString))
                                changedMesMovementsDTO.NeedWriteToSap = false;
                            else
                                changedMesMovementsDTO.NeedWriteToSap = needWriteToSapVarString.ToUpper().Equals("ДА") ? true : false;

                            Guid previousRecordIdVarGuid = Guid.Empty;
                            if (!String.IsNullOrEmpty(previousRecordIdVarString))
                            {
                                try
                                {
                                    previousRecordIdVarGuid = Guid.Parse(previousRecordIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 18 (\"ИД предыдущей записи\"). Не удалось получить ИД предыдущей записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 18 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                            }

                            if (previousRecordIdVarGuid == Guid.Empty)
                            {
                                changedMesMovementsDTO.PreviousRecordId = null;
                            }
                            else
                            {
                                MesMovementsDTO? objectForCheckMesMovementsPreviousRecordDTO = await _mesMovementsRepository.GetById(previousRecordIdVarGuid);
                                if (objectForCheckMesMovementsPreviousRecordDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 18 (\"ИД предыдущей записи\"). Не удалось найти запись в Архиве данных MesMovements." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 18 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.PreviousRecordId = objectForCheckMesMovementsPreviousRecordDTO.Id;
                                }
                            }

                            await _mesMovementsRepository.Update(changedMesMovementsDTO);
                            await _logEventRepository.ToLog<MesMovementsDTO>(oldObject: foundMesMovementsDTO, newObject: changedMesMovementsDTO
                                , "Изменение записи в Архиве данных MesMovements", "Запись: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;

                            worksheet.Cell(rowNumber, 1).Value = "ОК";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);

                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "ДОБАВИТЬ":
                        {
                            if (!String.IsNullOrEmpty(idVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Для действия \"Добавить\" ИД записи должен быть пустым. Сформируется автоматически." +
                                        " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            DateTime addTimeVarDateTime;

                            if (!String.IsNullOrEmpty(addTimeVarString))
                            {
                                try
                                {
                                    addTimeVarDateTime = DateTime.Parse(addTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 4 (\"Время добавления записи\"). Не удалось получить Время добавления записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 4 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.AddTime = addTimeVarDateTime;
                                changedMesMovementsDTO.AddTime = changedMesMovementsDTO.AddTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesMovementsDTO.AddTime;
                                changedMesMovementsDTO.AddTime = changedMesMovementsDTO.AddTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesMovementsDTO.AddTime;
                            }

                            Guid addUserIdVarGuid = Guid.Empty;

                            changedMesMovementsDTO.AddUserId = Guid.Empty;
                            changedMesMovementsDTO.AddUserDTOFK = new UserDTO();

                            if (!String.IsNullOrEmpty(addUserIdVarString))
                            {
                                try
                                {
                                    addUserIdVarGuid = Guid.Parse(addUserIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5 (\"ИД добавившего запись пользователя\"). Не удалось получить ИД добавившего запись пользователя." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 5 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                if (addUserIdVarGuid != Guid.Empty)
                                    if (!String.IsNullOrEmpty(addUserNameVarString))
                                    {
                                        var userObjectByGuid = await _userRepository.Get(addUserIdVarGuid);

                                        if (userObjectByGuid != null)
                                        {
                                            changedMesMovementsDTO.AddUserId = userObjectByGuid.Id;
                                            changedMesMovementsDTO.AddUserDTOFK = userObjectByGuid;
                                        }
                                    }
                            }


                            if (changedMesMovementsDTO.AddUserId == Guid.Empty)
                                if (!String.IsNullOrEmpty(addUserNameVarString))
                                {
                                    var userObjectList = await _userRepository.GetListByUserName(addUserNameVarString);
                                    if (userObjectList != null)
                                    {
                                        if (userObjectList.Count() > 1)
                                        {
                                            haveErrors = true;
                                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). Найдено больше одного пользователя с ФИО " + addUserNameVarString + "." +
                                                " Изменения не применялись.";
                                            await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                            rowNumber++;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        haveErrors = true;
                                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 6 (\"ФИО добавившего пользователя\"). В Справочнике пользователей не найден пользователь с ФИО " + addUserNameVarString + "." +
                                            " Изменения не применялись.";
                                        await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 6 }, resultString);
                                        rowNumber++;
                                        continue;
                                    }
                                    if (userObjectList.Count() > 0)
                                    {
                                        changedMesMovementsDTO.AddUserId = userObjectList.FirstOrDefault().Id;
                                        changedMesMovementsDTO.AddUserDTOFK = userObjectList.FirstOrDefault();
                                    }
                                }

                            if (changedMesMovementsDTO.AddUserId == Guid.Empty)
                            {
                                changedMesMovementsDTO.AddUserId = currentUserDTO.Id;
                                changedMesMovementsDTO.AddUserDTOFK = currentUserDTO;
                            }

                            if (!String.IsNullOrEmpty(mesParamCodeVarString))
                            {
                                MesParamDTO? objectForCheckMesParamCode = await _mesParamRepository.GetByCode(mesParamCodeVarString);
                                if (objectForCheckMesParamCode == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Код тэга\"). В Справочнике тэгов СИР не найден не найден тэг с кодом " + mesParamCodeVarString + "." +
                                        " Изменения не применялись.";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.MesParamId = objectForCheckMesParamCode.Id;
                                changedMesMovementsDTO.MesParamDTOFK = objectForCheckMesParamCode;
                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 7 (\"Код тэга\"). В режиме \"Добавить\" Код тэга не может быть пустым." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 7 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            DateTime valueTimeVarDateTime;
                            if (!String.IsNullOrEmpty(valueTimeVarString))
                            {
                                try
                                {
                                    valueTimeVarDateTime = DateTime.Parse(valueTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время значения\"). Не удалось получить Время значения." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.ValueTime = valueTimeVarDateTime;
                                changedMesMovementsDTO.ValueTime = changedMesMovementsDTO.ValueTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesMovementsDTO.ValueTime;
                                changedMesMovementsDTO.ValueTime = changedMesMovementsDTO.ValueTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesMovementsDTO.ValueTime;
                            }
                            else
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 8 (\"Время значения\"). В режиме \"Добавить\" Время значения не может быть пустым." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 8 }, resultString);
                                rowNumber++;
                                continue;
                            }


                            decimal valueDecimal;
                            if (!String.IsNullOrEmpty(valueVarString))
                            {
                                try
                                {
                                    valueDecimal = decimal.Parse(valueVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Не удалось получить Значение." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.Value = valueDecimal;
                            }
                            else
                            {
                                changedMesMovementsDTO.Value = 0;
                            }

                            if (changedMesMovementsDTO.Value < 0)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 9 (\"Значение\"). Значение не может быть отрицательным. Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 9 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (!String.IsNullOrEmpty(sapMovementInIdVarString))
                            {
                                SapMovementsINDTO? objectForCheckSapMovementsINDTO = await _sapMovementsINRepository.GetById(sapMovementInIdVarString);
                                if (objectForCheckSapMovementsINDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 10 (\"ИД записи во входной витрине SapMovementsIN\"). Не удалось найти запись с ИД "
                                        + sapMovementInIdVarString + " в Витрине SAP \"НДО-вход\". Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 10 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.SapMovementInId = objectForCheckSapMovementsINDTO.ErpId;
                                    changedMesMovementsDTO.SapMovementsINDTOFK = objectForCheckSapMovementsINDTO;
                                }
                            }
                            else
                            {
                                changedMesMovementsDTO.SapMovementInId = null;
                                changedMesMovementsDTO.SapMovementsINDTOFK = null;
                            }

                            Guid sapMovementOutIdVarGuid = Guid.Empty;
                            if (!String.IsNullOrEmpty(sapMovementOutIdVarString))
                            {
                                try
                                {
                                    sapMovementOutIdVarGuid = Guid.Parse(sapMovementOutIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД записи в выходной витрине SapMovementsOUT\"). Не удалось получить ИД записи в выходной витрине SapMovementsOUT." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 11 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                SapMovementsOUTDTO? objectForCheckSapMovementsOUTDTO = await _sapMovementsOUTRepository.GetById(sapMovementOutIdVarGuid);
                                if (objectForCheckSapMovementsOUTDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 11 (\"ИД записи в выходной витрине SapMovementsOUT\"). Не удалось найти запись с ИД "
                                        + sapMovementOutIdVarGuid.ToString() + " в Витрине SAP \"Движения-выход\". Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 12 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.SapMovementOutId = objectForCheckSapMovementsOUTDTO.Id;
                                    changedMesMovementsDTO.SapMovementsOUTDTOFK = objectForCheckSapMovementsOUTDTO;
                                }
                            }
                            else
                            {
                                changedMesMovementsDTO.SapMovementOutId = null;
                                changedMesMovementsDTO.SapMovementsOUTDTOFK = null;
                            }

                            if (String.IsNullOrEmpty(dataSourceNameVarString))
                            {
                                changedMesMovementsDTO.DataSourceId = null;
                                changedMesMovementsDTO.DataSourceDTOFK = null;
                            }
                            else
                            {
                                DataSourceDTO? objectForCheckDataSourceDTO = await _dataSourceRepository.GetByName(dataSourceNameVarString);
                                if (objectForCheckDataSourceDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 12 (\"Источник данных\"). Не удалось источник данных с наименованием "
                                        + dataSourceNameVarString + "." + " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 12 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.DataSourceId = objectForCheckDataSourceDTO.Id;
                                    changedMesMovementsDTO.DataSourceDTOFK = objectForCheckDataSourceDTO;
                                }
                            }

                            if (String.IsNullOrEmpty(dataTypeNameVarString))
                            {
                                changedMesMovementsDTO.DataTypeId = null;
                                changedMesMovementsDTO.DataTypeDTOFK = null;
                            }
                            else
                            {
                                DataTypeDTO? objectForCheckDataTypeDTO = await _dataTypeRepository.GetByName(dataTypeNameVarString);
                                if (objectForCheckDataTypeDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 13 (\"Тип данных\"). Не удалось Тип данных с наименованием "
                                        + dataTypeNameVarString + "." + " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 13 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.DataTypeId = objectForCheckDataTypeDTO.Id;
                                    changedMesMovementsDTO.DataTypeDTOFK = objectForCheckDataTypeDTO;
                                }
                            }


                            if (String.IsNullOrEmpty(reportGuidVarString))
                            {
                                changedMesMovementsDTO.ReportGuid = null;
                                changedMesMovementsDTO.ReportEntityDTOFK = null;
                            }
                            else
                            {
                                Guid reportGuidVarGuid = Guid.Empty;
                                try
                                {
                                    reportGuidVarGuid = Guid.Parse(reportGuidVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 14 (\"ИД Экземпляра отчёта\"). Не удалось получить ИД Экземпляра отчёта." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 14 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                ReportEntityDTO? objectForCheckReportEntityDTO = await _reportEntityRepository.GetById(reportGuidVarGuid);
                                if (objectForCheckReportEntityDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 14 (\"ИД Экземпляра отчёта\"). Не удалось найти экземпляр отчёта." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 14 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMovementsDTO.ReportGuid = objectForCheckReportEntityDTO.Id;
                                    changedMesMovementsDTO.ReportEntityDTOFK = objectForCheckReportEntityDTO;
                                }
                            }

                            if (String.IsNullOrEmpty(mesGoneVarString))
                                changedMesMovementsDTO.MesGone = false;
                            else
                                changedMesMovementsDTO.MesGone = mesGoneVarString.ToUpper().Equals("ДА") ? true : false;

                            DateTime mesGoneTimeVarDateTime;
                            if (!String.IsNullOrEmpty(mesGoneTimeVarString))
                            {
                                try
                                {
                                    mesGoneTimeVarDateTime = DateTime.Parse(mesGoneTimeVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 16 (\"Время MES забрал\"). Не удалось получить Время MES забрал." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 16 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.MesGoneTime = mesGoneTimeVarDateTime;
                                changedMesMovementsDTO.MesGoneTime = changedMesMovementsDTO.MesGoneTime < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue : changedMesMovementsDTO.MesGoneTime;
                                changedMesMovementsDTO.MesGoneTime = changedMesMovementsDTO.MesGoneTime > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue ? (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue : changedMesMovementsDTO.MesGoneTime;
                            }
                            else
                                changedMesMovementsDTO.MesGoneTime = null;

                            if (!((changedMesMovementsDTO.MesGone == true && changedMesMovementsDTO.MesGoneTime != null)
                                || (changedMesMovementsDTO.MesGone != true && changedMesMovementsDTO.MesGoneTime == null)))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 15, 16 (\"Mes забрал\", \"Время MES забрал\"). Если \"Mes забрал\" равно \"Да\", то \"Время MES забрал\" должно быть заполнено."
                                    + " Если \"Mes забрал\" равно \"Нет\" или пусто, то \"Время MES забрал\" должно быть пустым."
                                    + " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[3] { 2, 15, 16 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            if (String.IsNullOrEmpty(needWriteToSapVarString))
                                changedMesMovementsDTO.NeedWriteToSap = false;
                            else
                                changedMesMovementsDTO.NeedWriteToSap = needWriteToSapVarString.ToUpper().Equals("ДА") ? true : false;

                            if (String.IsNullOrEmpty(previousRecordIdVarString))
                            {
                                changedMesMovementsDTO.PreviousRecordId = null;
                            }
                            else
                            {
                                Guid previousRecordIdVarGuid = Guid.Empty;
                                try
                                {
                                    previousRecordIdVarGuid = Guid.Parse(previousRecordIdVarString);
                                }
                                catch (Exception ex)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 18 (\"ИД предыдущей записи\"). Не удалось получить ИД предыдущей записи." +
                                        " Изменения не применялись. Сообщение ошибки: " + ex.Message;
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 18 }, resultString);
                                    rowNumber++;
                                    continue;
                                }

                                MesMovementsDTO? objectForCheckMesMovementsPreviousRecordDTO = await _mesMovementsRepository.GetById(previousRecordIdVarGuid);
                                if (objectForCheckMesMovementsPreviousRecordDTO == null)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 18 (\"ИД предыдущей записи\"). Не удалось найти запись в Архиве данных MesMovements." +
                                        " Изменения не применялись";
                                    await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 18 }, resultString);
                                    rowNumber++;
                                    continue;
                                }
                                changedMesMovementsDTO.PreviousRecordId = objectForCheckMesMovementsPreviousRecordDTO.Id;
                            }

                            MesMovementsDTO? newMesMovementsDTO = await _mesMovementsRepository.Create(changedMesMovementsDTO);
                            await _logEventRepository.ToLog<MesMovementsDTO>(oldObject: null, newObject: newMesMovementsDTO, "Добавление записи в Архив данных MesMovements", "Запись: ", _authorizationRepository);
                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Добавлена запись с ИД " + newMesMovementsDTO.Id.ToString();
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 3).Value = newMesMovementsDTO.Id.ToString();
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    case "УДАЛИТЬ":
                        {
                            if (String.IsNullOrEmpty(idVarString))
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Для действия \"Удалить\" ИД записи единственное необходимое поле. Не может быть пустым." +
                                        " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            foundMesMovementsDTO = await _mesMovementsRepository.GetById(idVarGuid);
                            if (foundMesMovementsDTO == null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 3 (\"ИД записи\"). Запись с ИД " + idVarGuid.ToString() + " не найдена в Архиве данных." +
                                    " Изменения не применялись.";
                                await WriteError(loadFromExcelPage, worksheet, rowNumber, 1, resultColumnNumber, new int[2] { 2, 3 }, resultString);
                                rowNumber++;
                                continue;
                            }

                            var SapMovementsINListToClean = await _sapMovementsINRepository.GetListByMesMovementId(foundMesMovementsDTO.Id);
                            if (SapMovementsINListToClean != null)
                                if (SapMovementsINListToClean.Count() > 0)
                                {
                                    foreach (var SapMovementsINItemToClean in SapMovementsINListToClean)
                                    {
                                        await _logEventRepository.AddRecord("Изменение записи в витрине SAP Движения вход SapMovementsIN", currentUserDTO.Id,
                                            SapMovementsINItemToClean.MesMovementId.ToString(), "<Пусто>", false, "Запись:  " + SapMovementsINItemToClean.ErpId + ": Поле: Запись в архиве данных СИР");
                                        await _sapMovementsINRepository.CleanMesMovementId(SapMovementsINItemToClean);
                                    }
                                }

                            var SapMovementsOUTListToClean = await _sapMovementsOUTRepository.GetListByMesMovementId(foundMesMovementsDTO.Id);
                            if (SapMovementsOUTListToClean != null)
                                if (SapMovementsOUTListToClean.Count() > 0)
                                {
                                    foreach (var SapMovementsOUTItemToClean in SapMovementsOUTListToClean)
                                    {
                                        await _logEventRepository.AddRecord("Изменение записи в витрине SAP Движения выход SapMovementsOUT", currentUserDTO.Id,
                                            SapMovementsOUTItemToClean.MesMovementId.ToString(), "<Пусто>", false, "Запись:  " + SapMovementsOUTItemToClean.Id.ToString() + ": Поле: Запись в архиве данных СИР");
                                        await _sapMovementsOUTRepository.CleanMesMovementId(SapMovementsOUTItemToClean);
                                    }
                                }

                            var MesMovementsPreviousRecordsListToClean = await _mesMovementsRepository.GetListByPreviousRecordId(foundMesMovementsDTO.Id);
                            if (MesMovementsPreviousRecordsListToClean != null)
                                if (MesMovementsPreviousRecordsListToClean.Count() > 0)
                                {
                                    foreach (var MesMovementsPreviousRecordItemToClean in MesMovementsPreviousRecordsListToClean)
                                    {
                                        await _logEventRepository.AddRecord("Изменение записи в Архиве данных MesMovements", currentUserDTO.Id,
                                            MesMovementsPreviousRecordItemToClean.PreviousRecordId.ToString(), "<Пусто>", false, "Запись:  " + MesMovementsPreviousRecordItemToClean.Id.ToString() + ": Поле: ИД предыдущей записи");
                                        await _mesMovementsRepository.CleanPreviousRecordId(MesMovementsPreviousRecordItemToClean);
                                    }
                                }

                            await _logEventRepository.AddRecord("Удаление записи из Архива данных MesMovements", currentUserDTO.Id, "", "", false, "Удаление записи с ИД : " + foundMesMovementsDTO.Id.ToString());
                            await _mesMovementsRepository.Delete(foundMesMovementsDTO.Id);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Удалена запись с ИД " + foundMesMovementsDTO.Id.ToString() + ".";
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, 1).Value = "OK";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);

                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 2 (\"Действие\"). Не предусмотренное значение действия = " + actionVarString + ". Для витрины НДО-выход допустимы действия: \"Добавить\", \"Изменить\", \"Удалить\". Изменения не применялись.";
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


    }
}
