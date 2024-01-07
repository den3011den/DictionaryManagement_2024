using ClosedXML.Excel;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_Models.IntDBModels;
using DictionaryManagement_Server.Extensions.Repository.IRepository;

namespace DictionaryManagement_Server.Extensions.Repository
{
    public class LoadFromExcelRepository : ILoadFromExcelRepository
    {
        private readonly ISapMaterialRepository _sapMaterialRepository;
        private readonly IMesMaterialRepository _mesMaterialRepository;
        private readonly ISapEquipmentRepository _sapEquipmentRepository;

        public LoadFromExcelRepository(ISapMaterialRepository sapMaterialRepository, IMesMaterialRepository mesMaterialRepository,
            ISapEquipmentRepository sapEquipmentRepository)
        {
            _sapMaterialRepository = sapMaterialRepository;
            _mesMaterialRepository = mesMaterialRepository;
            _sapEquipmentRepository = sapEquipmentRepository;
        }

        public async Task<string> SapMaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Материалов SAP)";
            loadFromExcelPage.RefreshSate();

            IEnumerable<SapMaterialDTO> sapMaterialDTOList = (await _sapMaterialRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.Id);

            int recordCount = sapMaterialDTOList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var sapMaterialDTOItem in sapMaterialDTOList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    loadFromExcelPage.RefreshSate();
                }

                excelColNum = 2;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMaterialDTOItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMaterialDTOItem.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMaterialDTOItem.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMaterialDTOItem.ShortName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = sapMaterialDTOItem.IsArchive == true ? "Да" : "Нет";
                excelRowNum++;
            }
            return "SapMaterial_Example_with_data_";
        }

        public async Task<string> MesMaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Материалов MES)";
            loadFromExcelPage.RefreshSate();

            IEnumerable<MesMaterialDTO> mesMaterialDTOList = (await _mesMaterialRepository.GetAll(SD.SelectDictionaryScope.All)).OrderBy(u => u.Id);

            int recordCount = mesMaterialDTOList.Count();
            int recordOrder = 0;

            int excelRowNum = 9;
            int excelColNum;
            foreach (var mesMaterialDTOItem in mesMaterialDTOList)
            {
                recordOrder++;
                if ((recordOrder == 1) || (recordOrder % 50) == 0)
                {
                    loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (обрабатывается запись " + recordOrder.ToString() + " из " + recordCount.ToString() + ")";
                    loadFromExcelPage.RefreshSate();
                }

                excelColNum = 2;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMaterialDTOItem.Id.ToString();
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMaterialDTOItem.Code;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMaterialDTOItem.Name;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMaterialDTOItem.ShortName;
                excelColNum++;
                worksheet.Cell(excelRowNum, excelColNum).Value = mesMaterialDTOItem.IsArchive == true ? "Да" : "Нет";
                excelRowNum++;
            }

            return "MesMaterial_Example_with_data_";
        }
        public async Task<string> SapEquipmentReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Оборудования SAP)";
            loadFromExcelPage.RefreshSate();

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
                    loadFromExcelPage.RefreshSate();
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
            return "MesParam_Example_with_data_";
        }
    }
}
