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

        public LoadFromExcelRepository(ISapMaterialRepository sapMaterialRepository, IMesMaterialRepository mesMaterialRepository,
            ISapEquipmentRepository sapEquipmentRepository,
            ILogEventRepository logEventRepository)
        {
            _sapMaterialRepository = sapMaterialRepository;
            _mesMaterialRepository = mesMaterialRepository;
            _sapEquipmentRepository = sapEquipmentRepository;
            _logEventRepository = logEventRepository;
        }

        public async Task<string> MaterialReportTemplateDownloadFileWithData(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet)
        {
            loadFromExcelPage.reportTemplateDownloadFileWithDataBusyText = "Выполняется ... (получение списка Материалов " + worksheet.Name.Substring(0, 3);
            loadFromExcelPage.RefreshSate();

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
                    loadFromExcelPage.RefreshSate();
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

        public async Task<bool> MaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист " + worksheet.Name + " загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа " + worksheet.Name + " в справочник Материалов " + worksheet.Name.Substring(0, 3));
            loadFromExcelPage.RefreshSate();

            int rowNumber = 9;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                loadFromExcelPage.RefreshSate();

                var rowVar = worksheet.Row(rowNumber);

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
                    worksheet.Cell(rowNumber, 7).Value = resultString;
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
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

                        changedMaterialDTO.Code = foundMaterialDTO.Code;
                        changedMaterialDTO.Name = nameVarString;
                        changedMaterialDTO.ShortName = shortNameVarString;
                    }
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
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
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

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Материал добавлен с кодом " + newMaterialDTO.Id.ToString();
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            loadFromExcelPage.RefreshSate();
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

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            resultString = "!!! Для строки " + rowNumber.ToString() + " определен не предусмотренный режим обработки = " + isUpdateOrAddMode + ". Изменения не производились.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа " + worksheet.Name + " в справочник Материалов " + worksheet.Name.Substring(0, 3));
            loadFromExcelPage.RefreshSate();
            return haveErrors;
        }

        public async Task<bool> SapEquipmentExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
        IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист SapEquipment загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа SapEquipment в справочник Ресурсов SAP");
            loadFromExcelPage.RefreshSate();

            int rowNumber = 9;

            bool isEmptyString = false;

            while (isEmptyString == false)
            {

                loadFromExcelPage.console.Log($"Обработка строки " + rowNumber.ToString());
                loadFromExcelPage.RefreshSate();

                var rowVar = worksheet.Row(rowNumber);

                string idVarString = rowVar.Cell(2).CachedValue.ToString().Trim();
                string erpPlantIdVarString = rowVar.Cell(3).CachedValue.ToString().Trim();
                string erpIdVarString = rowVar.Cell(4).CachedValue.ToString().Trim();
                string nameVarString = rowVar.Cell(5).CachedValue.ToString().Trim();
                string isWarehouseVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string isArchiveVarString = rowVar.Cell(6).CachedValue.ToString().Trim();
                string resultString = "";
                int idVarInt = 0;

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
                    worksheet.Cell(rowNumber, 7).Value = resultString;
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    loadFromExcelPage.RefreshSate();
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
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (idVarInt <= 0)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Неверное значение ИД записи равное " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }

                    if (!((String.IsNullOrEmpty(erpPlantIdVarString) && String.IsNullOrEmpty(erpIdVarString)) ||
                        (!String.IsNullOrEmpty(erpPlantIdVarString) && !String.IsNullOrEmpty(erpIdVarString))))
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 3, 4. Если указан ИД записи (режим редактирования записи), то \"Код завода SAP\" и \"Код ресурса/склада SAP\" должны быть" +
                            " или оба пусты (тогда останутся без изменений), или оба заполнены. Изменения не применялись.";
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }


                    foundSapEquipmentDTO = await _sapEquipmentRepository.Get(idVarInt);

                    if (foundSapEquipmentDTO == null)
                    {
                        haveErrors = true;
                        resultString = "! Строка " + rowNumber.ToString() + ", столбец 2. Не найдена запись в справочнике с ИД записи равным " + idVarInt.ToString() + ".Изменения не применялись.";
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                        loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                        loadFromExcelPage.RefreshSate();
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

                    if (String.IsNullOrEmpty(isWarehouseVarString))
                    {
                        changedSapEquipmentDTO.IsWarehouse = false;
                    }
                    else
                    {
                        changedSapEquipmentDTO.IsWarehouse = isWarehouseVarString.ToUpper() == "ДА" ? true : false;
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
                        }
                        else  // редактирование
                        {
                            isUpdateOrAddMode = "UPDATE";
                            needCheckErpPlantIdPlusErpId = false;
                            needCheckName = true;

                            changedSapEquipmentDTO.ErpPlantId = foundSapEquipmentDTO.ErpPlantId;
                            changedSapEquipmentDTO.ErpId = foundSapEquipmentDTO.ErpId;
                            changedSapEquipmentDTO.Name = nameVarString;
                        }

                        if (String.IsNullOrEmpty(isWarehouseVarString))
                        {
                            changedSapEquipmentDTO.IsWarehouse = false;
                        }
                        else
                        {
                            changedSapEquipmentDTO.IsWarehouse = isWarehouseVarString.ToUpper() == "Да" ? true : false;
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
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
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

                            SapEquipmentDTO? newSapEquipmentDTO = await _sapEquipmentRepository.Create(changedSapEquipmentDTO);

                            await _logEventRepository.ToLog<SapEquipmentDTO>(oldObject: null, newObject: newSapEquipmentDTO, "Добавление ресурса SAP", "Ресурс SAP: ", _authorizationRepository);

                            resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Ресурс Sap добавлен с кодом " + newSapEquipmentDTO.Id.ToString()
                                + ". " + duplicateNameString;
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            if (String.IsNullOrEmpty(duplicateNameString))
                            {
                                worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.PowderBlue;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            loadFromExcelPage.RefreshSate();
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
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            if (String.IsNullOrEmpty(duplicateNameString))
                            {
                                worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                                loadFromExcelPage.console.Log(resultString);
                            }
                            else
                            {
                                worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.PowderBlue;
                                loadFromExcelPage.console.Log(resultString, AlertStyle.Light);
                            }
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                    default:
                        {
                            resultString = "!!! Для строки " + rowNumber.ToString() + " определен не предусмотренный режим обработки = " + isUpdateOrAddMode + ". Изменения не производились.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                }
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа SapEquipment в справочник Ресурсов SAP");
            loadFromExcelPage.RefreshSate();

            return haveErrors;
        }


        public async Task<bool> MesParamExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист MesParam загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа MesParam в справочник Тэгов СИР");
            loadFromExcelPage.RefreshSate();

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа MesParam в справочник Тэгов СИР");
            loadFromExcelPage.RefreshSate();

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
