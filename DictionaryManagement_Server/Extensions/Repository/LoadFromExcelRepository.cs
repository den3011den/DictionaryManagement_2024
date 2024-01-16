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

        public LoadFromExcelRepository(ISapMaterialRepository sapMaterialRepository, IMesMaterialRepository mesMaterialRepository,
            ISapEquipmentRepository sapEquipmentRepository,
            ILogEventRepository logEventRepository, IMesParamRepository mesParamRepository,
            IMesParamSourceTypeRepository mesParamSourceTypeRepository,
            IMesDepartmentRepository mesDepartmentRepository,
            ISapUnitOfMeasureRepository sapUnitOfMeasureRepository)
        {
            _sapMaterialRepository = sapMaterialRepository;
            _mesMaterialRepository = mesMaterialRepository;
            _sapEquipmentRepository = sapEquipmentRepository;
            _logEventRepository = logEventRepository;
            _mesParamRepository = mesParamRepository;
            _mesParamSourceTypeRepository = mesParamSourceTypeRepository;
            _mesDepartmentRepository = mesDepartmentRepository;
            _sapUnitOfMeasureRepository = sapUnitOfMeasureRepository;
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

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

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

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(codeVarString))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3. И ИД записи, и Код материала пустые. Изменения не применялись.";
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 4).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 4).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                if (String.IsNullOrEmpty(changedMaterialDTO.ShortName))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Сокращённое наименование материала не может быть пустым. Изменения не применялись.";
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 5).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 5).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 4).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 4).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 5).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 5).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

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

                if (String.IsNullOrEmpty(idVarString) && (String.IsNullOrEmpty(erpPlantIdVarString) || String.IsNullOrEmpty(erpIdVarString)))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3, 4. Пустые: ИД записи, и одно из полей \"Код завода SAP\" или . \"Код ресурса/склада SAP\". Изменения не применялись.";
                    worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 4).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 4).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);


                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (!((String.IsNullOrEmpty(erpPlantIdVarString) && String.IsNullOrEmpty(erpIdVarString)) ||
                        (!String.IsNullOrEmpty(erpPlantIdVarString) && !String.IsNullOrEmpty(erpIdVarString))))
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3, 4. Если указан ИД записи (режим редактирования записи), то \"Код завода SAP\" и \"Код ресурса/склада SAP\" должны быть" +
                            " или оба пусты (тогда останутся без изменений), или оба заполнены. Изменения не применялись.";
                        worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 4).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 4).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }


                    foundSapEquipmentDTO = await _sapEquipmentRepository.Get(idVarInt);

                    if (foundSapEquipmentDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не найдена запись в справочнике с ИД записи равным " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 5).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 5).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 4).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 4).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Green;
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                await loadFromExcelPage.RefreshSate();

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

                if (String.IsNullOrEmpty(idVarString) && String.IsNullOrEmpty(codeVarString))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3. Пустые одновременно поля \"ИД записи\" и \"Код тэга СИР\". Изменения не применялись.";
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    foundMesParamDTO = await _mesParamRepository.GetById(idVarInt);

                    if (foundMesParamDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не найдена запись в справочнике с ИД записи равным " + idVarInt.ToString() + ". Изменения не применялись.";
                        worksheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 2).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                    worksheet.Cell(rowNumber, 6).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 6).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 3).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 3).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 8).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 8).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                        foundMesDepartment = _mesDepartmentRepository.GetById(departmentIdVarInt).GetAwaiter().GetResult();

                        if (foundMesDepartment == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 8. Не удалось найти производство с \"ИД производства\" равным " + departmentIdVarInt.ToString() + ". Изменения не применялись.";
                            worksheet.Cell(rowNumber, 8).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 8).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                                    worksheet.Cell(rowNumber, 9).Style.Font.FontColor = XLColor.Red;
                                    worksheet.Cell(rowNumber, 9).Style.Font.SetBold(true);
                                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                    await loadFromExcelPage.RefreshSate();
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
                                        worksheet.Cell(rowNumber, 9).Style.Font.FontColor = XLColor.Red;
                                        worksheet.Cell(rowNumber, 9).Style.Font.SetBold(true);
                                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                        await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 8).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 8).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 9).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 9).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 10).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 10).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                        foundSapEquipmentSourceDTO = _sapEquipmentRepository.Get(sapEquipmentIdSourceVarInt).GetAwaiter().GetResult();

                        if (foundSapEquipmentSourceDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 10. Не найден Ресурс-источник SAP c \"ИД ресурса-источника SAP\" равным " + sapEquipmentIdSourceVarInt.ToString() + ". Изменения не применялись.";
                            worksheet.Cell(rowNumber, 10).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 10).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 11).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 11).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 12).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 12).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 11).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 11).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 12).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 12).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 13).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 13).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapEquipmentSourceDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 10, 11, 12, 13. Ресурс-источник SAP не найден в справочнике Ресурсы SAP. Изменения не применялись.";
                        worksheet.Cell(rowNumber, 10).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 10).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 11).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 11).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 12).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 12).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 13).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 13).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 14).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 14).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                        foundSapEquipmentDestDTO = _sapEquipmentRepository.Get(sapEquipmentIdDestVarInt).GetAwaiter().GetResult();

                        if (foundSapEquipmentDestDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 14. Не найден Ресурс-приёмник SAP c \"ИД ресурса-приёмника SAP\" равным " + sapEquipmentIdDestVarInt.ToString() + ". Изменения не применялись.";
                            worksheet.Cell(rowNumber, 14).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 14).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 15).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 15).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 16).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 16).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 15).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 15).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 16).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 16).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 17).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 17).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    }

                    if (foundSapEquipmentDestDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 14, 15, 16, 17. Ресурс-приёмник SAP не найден в справочнике Ресурсы SAP. Изменения не применялись.";
                        worksheet.Cell(rowNumber, 14).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 14).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 15).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 15).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 16).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 16).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 17).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 17).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 18).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 18).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                        foundSapMaterialDTO = _sapMaterialRepository.Get(sapMaterialIdVarInt).GetAwaiter().GetResult();

                        if (foundSapMaterialDTO == null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 18. Не найден Материал SAP c \"ИД материала SAP\" равным " + sapMaterialIdVarInt.ToString() + ". Изменения не применялись.";
                            worksheet.Cell(rowNumber, 18).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 18).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 19).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 19).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 20).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 20).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
                                rowNumber++;
                                continue;
                            }
                        }
                    }
                    if (foundSapMaterialDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 18, 19, 20. Материал SAP не найден в справочнике Материалы SAP. Изменения не применялись.";
                        worksheet.Cell(rowNumber, 18).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 18).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 19).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 19).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 20).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 20).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                                worksheet.Cell(rowNumber, 10).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 11).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 12).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 13).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 14).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 15).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 16).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 17).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 18).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 19).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 20).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 10).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 11).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 12).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 13).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 14).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 15).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 16).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 17).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 18).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 19).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 20).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, 1).Value = "!!!";
                                worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                                worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                                worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                await loadFromExcelPage.RefreshSate();
                                rowNumber++;
                                continue;
                            }
                        }
                    }
                    else
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ". Не полный мэппинг Тэга СИР с SAP по связке параметров \"Источник SAP + Приёмник SAP + Материал SAP\". Должны быть проставлены или все 3 параметра, или ни одного. Изменения не применялись.";
                        worksheet.Cell(rowNumber, 10).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 11).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 12).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 13).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 14).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 15).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 16).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 17).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 18).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 19).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 20).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 10).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 11).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 12).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 13).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 14).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 15).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 16).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 17).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 18).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 19).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 20).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 21).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 21).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 22).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 22).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 27).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 27).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, 1).Value = "!!!";
                        worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                        worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        await loadFromExcelPage.RefreshSate();
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
                    worksheet.Cell(rowNumber, 29).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 31).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 32).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 29).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 31).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 32).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedReadFromMes && (bool)changedMesParamDTO.NeedWriteToMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Читать из MES\", \"Передавать в MES\". Тэг СИР не может одновременно иметь включенными признаки \"Читать из MES\" и \"Передавать в MES\"." +
                        " Изменения не применялись";
                    worksheet.Cell(rowNumber, 30).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 31).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 30).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 31).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedReadFromSap && (bool)changedMesParamDTO.NeedWriteToSap))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Передавать в SAP\", \"Читать из SAP\". Тэг СИР не может одновременно иметь включенными признаки \"Передавать в SAP\" и \"Читать из SAP\"." +
                        " Изменения не применялись";
                    worksheet.Cell(rowNumber, 28).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 29).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 28).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 29).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedWriteToSap && (bool)changedMesParamDTO.NeedWriteToMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Передавать в SAP\", \"Передавать в MES\". Тэг СИР не может одновременно иметь включенными признаки \"Передавать в SAP\" и \"Передавать в MES\"." +
                        " Изменения не применялись";
                    worksheet.Cell(rowNumber, 28).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 31).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 28).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 31).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                if (((bool)changedMesParamDTO.NeedReadFromSap && (bool)changedMesParamDTO.NeedReadFromMes))
                {
                    haveErrors = true;
                    resultString = "! Строка " + rowNumber.ToString() + ", поля \"Читать из SAP\", \"Читать из MES\". Тэг СИР не может одновременно иметь включенными признаки \"Читать из SAP\" и \"Читать из MES\"." +
                        " Изменения не применялись";
                    worksheet.Cell(rowNumber, 29).Style.Font.FontColor = XLColor.Red; worksheet.Cell(rowNumber, 30).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 29).Style.Font.SetBold(true); worksheet.Cell(rowNumber, 30).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, 1).Value = "!!!";
                    worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                    worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    await loadFromExcelPage.RefreshSate();
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
                            worksheet.Row(rowNumber).Style.Font.FontColor = XLColor.Red;
                            worksheet.Row(rowNumber).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, 1).Value = "!!!";
                            worksheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 1).Style.Font.SetBold(true);
                            worksheet.Cell(rowNumber, resultColumnNumber).Value = resultString;
                            worksheet.Cell(rowNumber, resultColumnNumber).Style.Font.FontColor = XLColor.Red;
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            await loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа SapEquipment в справочник Ресурсов SAP");
            await loadFromExcelPage.RefreshSate();

            return haveErrors;

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
