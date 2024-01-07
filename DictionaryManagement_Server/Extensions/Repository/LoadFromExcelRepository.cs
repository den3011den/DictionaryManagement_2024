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

        public async Task<bool> SapMaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист SapMaterial загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа SapMaterial в справочник Материалов SAP");
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

                if (String.IsNullOrEmpty(codeVarString) && String.IsNullOrEmpty(codeVarString))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3. И ИД записи, и код пустые. Изменения не применялись.";
                    worksheet.Cell(rowNumber, 7).Value = resultString;
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                SapMaterialDTO? foundSapMaterialDTO = null;
                if (!String.IsNullOrEmpty(idVarString))  // если указан id элемента, то рассматриваем только вариант редактирования уже существующей записи
                {
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

                    foundSapMaterialDTO = await _sapMaterialRepository.Get(idVarInt);
                    if (foundSapMaterialDTO == null)
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

                    SapMaterialDTO changedSapMaterialDTO = new SapMaterialDTO();
                    changedSapMaterialDTO.Id = foundSapMaterialDTO.Id;

                    if (String.IsNullOrEmpty(codeVarString))
                    {
                        changedSapMaterialDTO.Code = foundSapMaterialDTO.Code;
                    }
                    else
                    {
                        if (foundSapMaterialDTO.Code.Equals(codeVarString))
                        {
                            changedSapMaterialDTO.Code = foundSapMaterialDTO.Code;
                        }
                        else
                        {
                            var objectForCheckCode = _sapMaterialRepository.GetByCode(codeVarString).Result;
                            if (objectForCheckCode != null)
                            {
                                if (objectForCheckCode.Id != foundSapMaterialDTO.Id)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 3. Уже есть запись с кодом материала " + codeVarString
                                        + ". ИД записи: " + objectForCheckCode.Id.ToString() +
                                        ". Наименование: " + objectForCheckCode.ShortName + ". Изменения не применялись.";
                                    worksheet.Cell(rowNumber, 7).Value = resultString;
                                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                    loadFromExcelPage.RefreshSate();
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedSapMaterialDTO.Code = codeVarString;
                                }
                            }
                            else
                            {
                                changedSapMaterialDTO.Code = codeVarString;
                            }
                        }
                    }


                    if (String.IsNullOrEmpty(nameVarString))
                    {
                        changedSapMaterialDTO.Name = foundSapMaterialDTO.Name;
                    }
                    else
                    {
                        if (foundSapMaterialDTO.Name.Equals(nameVarString))
                        {
                            changedSapMaterialDTO.Name = foundSapMaterialDTO.Name;
                        }
                        else
                        {
                            var objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                            if (objectForCheckName != null)
                            {
                                if (objectForCheckName.Id != foundSapMaterialDTO.Id)
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
                                else
                                {
                                    changedSapMaterialDTO.Name = nameVarString;
                                }
                            }
                            else
                            {
                                changedSapMaterialDTO.Name = nameVarString;
                            }
                        }
                    }

                    if (String.IsNullOrEmpty(shortNameVarString))
                    {
                        changedSapMaterialDTO.ShortName = foundSapMaterialDTO.ShortName;
                    }
                    else
                    {
                        if (foundSapMaterialDTO.ShortName.Equals(shortNameVarString))
                        {
                            changedSapMaterialDTO.ShortName = foundSapMaterialDTO.ShortName;
                        }
                        else
                        {
                            var objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                            if (objectForCheckShortName != null)
                            {
                                if (objectForCheckShortName.Id != foundSapMaterialDTO.Id)
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
                                else
                                {
                                    changedSapMaterialDTO.ShortName = shortNameVarString;
                                }
                            }
                            else
                            {
                                changedSapMaterialDTO.ShortName = shortNameVarString;
                            }
                        }
                    }

                    if (String.IsNullOrEmpty(isArchiveVarString))
                    {
                        changedSapMaterialDTO.IsArchive = false;
                    }
                    else
                    {
                        changedSapMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                    }

                    await _sapMaterialRepository.Update(changedSapMaterialDTO, SD.UpdateMode.Update);

                    if (changedSapMaterialDTO.IsArchive != foundSapMaterialDTO.IsArchive)
                        if (changedSapMaterialDTO.IsArchive == true)
                        {
                            await _sapMaterialRepository.Update(changedSapMaterialDTO, SD.UpdateMode.MoveToArchive);
                        }
                        else
                        {
                            await _sapMaterialRepository.Update(changedSapMaterialDTO, SD.UpdateMode.RestoreFromArchive);
                        }

                    await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: foundSapMaterialDTO, newObject: changedSapMaterialDTO, "Изменение материала SAP", "Материал SAP: ", _authorizationRepository);

                    resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                    worksheet.Cell(rowNumber, 7).Value = resultString;
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                    loadFromExcelPage.console.Log(resultString);
                    loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                //-------------------------------------

                if (!String.IsNullOrEmpty(codeVarString))  // если указан Code элемента, то может быть как обновление записи, так и добавление
                {
                    foundSapMaterialDTO = await _sapMaterialRepository.GetByCode(codeVarString);
                    if (foundSapMaterialDTO == null)  // добавление
                    {

                        SapMaterialDTO forAddSapMaterialDTO = new SapMaterialDTO();

                        var objectForCheckCode = _sapMaterialRepository.GetByCode(codeVarString).Result;
                        if (objectForCheckCode != null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 3. Уже есть запись с кодом " + codeVarString +
                                ". ИД записи: " + objectForCheckCode.Id.ToString() + " Наименование: " + objectForCheckCode.ShortName + ".Изменения не применялись.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                        forAddSapMaterialDTO.Code = codeVarString;

                        if (String.IsNullOrEmpty(nameVarString))
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 4. Для добавляемой записи не может быть пустое наименование. Изменения не применялись.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                        else
                        {
                            var objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                            if (objectForCheckName != null)
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
                            forAddSapMaterialDTO.Name = nameVarString;
                        }

                        if (String.IsNullOrEmpty(shortNameVarString))
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Для добавляемой записи не может быть пустое сокращённое наименование. Изменения не применялись.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                        else
                        {
                            var objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                            if (objectForCheckShortName != null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Уже есть запись с сокращённым наименованием " + shortNameVarString
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
                            forAddSapMaterialDTO.ShortName = shortNameVarString;
                        }

                        forAddSapMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                        var newSapMaterialDTO = await _sapMaterialRepository.Create(forAddSapMaterialDTO);
                        await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: null, newObject: newSapMaterialDTO, "Добавление материала SAP", "Материал SAP: ", _authorizationRepository);

                        resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Материал добавлен с кодом " + newSapMaterialDTO.Id.ToString();
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                        loadFromExcelPage.console.Log(resultString);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }
                    else  // изменение
                    {
                        SapMaterialDTO forChangeSapMaterialDTO = new SapMaterialDTO();
                        forChangeSapMaterialDTO.Id = foundSapMaterialDTO.Id;
                        forChangeSapMaterialDTO.Code = foundSapMaterialDTO.Code;
                        if (String.IsNullOrEmpty(nameVarString))
                        {
                            forChangeSapMaterialDTO.Name = foundSapMaterialDTO.Name;
                        }
                        else
                        {
                            if (forChangeSapMaterialDTO.Name.Equals(nameVarString))
                            {
                                forChangeSapMaterialDTO.Name = foundSapMaterialDTO.Name;
                            }
                            else
                            {
                                var objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                                if (objectForCheckName != null)
                                {
                                    if (objectForCheckName.Id != foundSapMaterialDTO.Id)
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
                                    else
                                    {
                                        forChangeSapMaterialDTO.Name = nameVarString;
                                    }
                                }
                                else
                                {
                                    forChangeSapMaterialDTO.Name = nameVarString;
                                }
                            }
                        }


                        if (String.IsNullOrEmpty(shortNameVarString))
                        {
                            forChangeSapMaterialDTO.ShortName = foundSapMaterialDTO.ShortName;
                        }
                        else
                        {
                            if (foundSapMaterialDTO.ShortName.Equals(shortNameVarString))
                            {
                                forChangeSapMaterialDTO.ShortName = foundSapMaterialDTO.ShortName;
                            }
                            else
                            {
                                var objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                                if (objectForCheckShortName != null)
                                {
                                    if (objectForCheckShortName.Id != foundSapMaterialDTO.Id)
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
                                    else
                                    {
                                        forChangeSapMaterialDTO.ShortName = shortNameVarString;
                                    }
                                }
                                else
                                {
                                    forChangeSapMaterialDTO.ShortName = shortNameVarString;
                                }
                            }
                        }

                        if (String.IsNullOrEmpty(isArchiveVarString))
                        {
                            forChangeSapMaterialDTO.IsArchive = false;
                        }
                        else
                        {
                            forChangeSapMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                        }

                        await _sapMaterialRepository.Update(forChangeSapMaterialDTO, SD.UpdateMode.Update);

                        if (forChangeSapMaterialDTO.IsArchive != foundSapMaterialDTO.IsArchive)
                            if (forChangeSapMaterialDTO.IsArchive == true)
                            {
                                await _sapMaterialRepository.Update(forChangeSapMaterialDTO, SD.UpdateMode.MoveToArchive);
                            }
                            else
                            {
                                await _sapMaterialRepository.Update(forChangeSapMaterialDTO, SD.UpdateMode.RestoreFromArchive);
                            }

                        await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: foundSapMaterialDTO, newObject: forChangeSapMaterialDTO, "Изменение материала SAP", "Материал SAP: ", _authorizationRepository);

                        resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                        loadFromExcelPage.console.Log(resultString);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }
                }
                rowNumber++;
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа SapMaterial в справочник Материалов SAP");
            loadFromExcelPage.RefreshSate();
            return haveErrors;
        }


        public async Task<bool> MesMaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
                IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист MesMaterial загружен в память");
            loadFromExcelPage.console.Log($"Начало загрузки данных листа MesMaterial в справочник Материалов Mes");
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

                if (String.IsNullOrEmpty(codeVarString) && String.IsNullOrEmpty(codeVarString))
                {
                    haveErrors = true;
                    idVarInt = 0;
                    resultString = "! Строка " + rowNumber.ToString() + ", столбцы 2, 3. И ИД записи, и код пустые. Изменения не применялись.";
                    worksheet.Cell(rowNumber, 7).Value = resultString;
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                    loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }


                MesMaterialDTO? foundMesMaterialDTO = null;
                if (!String.IsNullOrEmpty(idVarString))  // если указан id элемента, то рассматриваем только вариант редактирования уже существующей записи
                {
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

                    foundMesMaterialDTO = await _mesMaterialRepository.Get(idVarInt);
                    if (foundMesMaterialDTO == null)
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

                    MesMaterialDTO changedMesMaterialDTO = new MesMaterialDTO();
                    changedMesMaterialDTO.Id = foundMesMaterialDTO.Id;

                    if (String.IsNullOrEmpty(codeVarString))
                    {
                        changedMesMaterialDTO.Code = foundMesMaterialDTO.Code;
                    }
                    else
                    {
                        if (foundMesMaterialDTO.Code.Equals(codeVarString))
                        {
                            changedMesMaterialDTO.Code = foundMesMaterialDTO.Code;
                        }
                        else
                        {
                            var objectForCheckCode = _mesMaterialRepository.GetByCode(codeVarString).Result;
                            if (objectForCheckCode != null)
                            {
                                if (objectForCheckCode.Id != foundMesMaterialDTO.Id)
                                {
                                    haveErrors = true;
                                    resultString = "! Строка " + rowNumber.ToString() + ", столбец 3. Уже есть запись с кодом материала " + codeVarString
                                        + ". ИД записи: " + objectForCheckCode.Id.ToString() +
                                        ". Наименование: " + objectForCheckCode.ShortName + ". Изменения не применялись.";
                                    worksheet.Cell(rowNumber, 7).Value = resultString;
                                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                                    worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                                    loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                                    loadFromExcelPage.RefreshSate();
                                    rowNumber++;
                                    continue;
                                }
                                else
                                {
                                    changedMesMaterialDTO.Code = codeVarString;
                                }
                            }
                            else
                            {
                                changedMesMaterialDTO.Code = codeVarString;
                            }
                        }
                    }


                    if (String.IsNullOrEmpty(nameVarString))
                    {
                        changedMesMaterialDTO.Name = foundMesMaterialDTO.Name;
                    }
                    else
                    {
                        if (foundMesMaterialDTO.Name.Equals(nameVarString))
                        {
                            changedMesMaterialDTO.Name = foundMesMaterialDTO.Name;
                        }
                        else
                        {
                            var objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;
                            if (objectForCheckName != null)
                            {
                                if (objectForCheckName.Id != foundMesMaterialDTO.Id)
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
                                else
                                {
                                    changedMesMaterialDTO.Name = nameVarString;
                                }
                            }
                            else
                            {
                                changedMesMaterialDTO.Name = nameVarString;
                            }
                        }
                    }

                    if (String.IsNullOrEmpty(shortNameVarString))
                    {
                        changedMesMaterialDTO.ShortName = foundMesMaterialDTO.ShortName;
                    }
                    else
                    {
                        if (foundMesMaterialDTO.ShortName.Equals(shortNameVarString))
                        {
                            changedMesMaterialDTO.ShortName = foundMesMaterialDTO.ShortName;
                        }
                        else
                        {
                            var objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;
                            if (objectForCheckShortName != null)
                            {
                                if (objectForCheckShortName.Id != foundMesMaterialDTO.Id)
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
                                else
                                {
                                    changedMesMaterialDTO.ShortName = shortNameVarString;
                                }
                            }
                            else
                            {
                                changedMesMaterialDTO.ShortName = shortNameVarString;
                            }
                        }
                    }

                    if (String.IsNullOrEmpty(isArchiveVarString))
                    {
                        changedMesMaterialDTO.IsArchive = false;
                    }
                    else
                    {
                        changedMesMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                    }

                    await _mesMaterialRepository.Update(changedMesMaterialDTO, SD.UpdateMode.Update);

                    if (changedMesMaterialDTO.IsArchive != foundMesMaterialDTO.IsArchive)
                        if (changedMesMaterialDTO.IsArchive == true)
                        {
                            await _mesMaterialRepository.Update(changedMesMaterialDTO, SD.UpdateMode.MoveToArchive);
                        }
                        else
                        {
                            await _mesMaterialRepository.Update(changedMesMaterialDTO, SD.UpdateMode.RestoreFromArchive);
                        }

                    await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: foundMesMaterialDTO, newObject: changedMesMaterialDTO, "Изменение материала MES", "Материал MES: ", _authorizationRepository);

                    resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                    worksheet.Cell(rowNumber, 7).Value = resultString;
                    worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                    loadFromExcelPage.console.Log(resultString);
                    loadFromExcelPage.RefreshSate();
                    rowNumber++;
                    continue;
                }

                //-------------------------------------

                if (!String.IsNullOrEmpty(codeVarString))  // если указан Code элемента, то может быть как обновление записи, так и добавление
                {
                    foundMesMaterialDTO = await _mesMaterialRepository.GetByCode(codeVarString);
                    if (foundMesMaterialDTO == null)  // добавление
                    {

                        MesMaterialDTO forAddMesMaterialDTO = new MesMaterialDTO();

                        var objectForCheckCode = _mesMaterialRepository.GetByCode(codeVarString).Result;
                        if (objectForCheckCode != null)
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 3. Уже есть запись с кодом " + codeVarString +
                                ". ИД записи: " + objectForCheckCode.Id.ToString() + " Наименование: " + objectForCheckCode.ShortName + ".Изменения не применялись.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }

                        forAddMesMaterialDTO.Code = codeVarString;

                        if (String.IsNullOrEmpty(nameVarString))
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 4. Для добавляемой записи не может быть пустое наименование. Изменения не применялись.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                        else
                        {
                            var objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;
                            if (objectForCheckName != null)
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
                            forAddMesMaterialDTO.Name = nameVarString;
                        }

                        if (String.IsNullOrEmpty(shortNameVarString))
                        {
                            haveErrors = true;
                            resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Для добавляемой записи не может быть пустое сокращённое наименование. Изменения не применялись.";
                            worksheet.Cell(rowNumber, 7).Value = resultString;
                            worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Red;
                            worksheet.Cell(rowNumber, 7).Style.Font.SetBold(true);
                            loadFromExcelPage.console.Log(resultString, AlertStyle.Danger);
                            loadFromExcelPage.RefreshSate();
                            rowNumber++;
                            continue;
                        }
                        else
                        {
                            var objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;
                            if (objectForCheckShortName != null)
                            {
                                haveErrors = true;
                                resultString = "! Строка " + rowNumber.ToString() + ", столбец 5. Уже есть запись с сокращённым наименованием " + shortNameVarString
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
                            forAddMesMaterialDTO.ShortName = shortNameVarString;
                        }

                        forAddMesMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                        var newMesMaterialDTO = await _mesMaterialRepository.Create(forAddMesMaterialDTO);
                        await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: null, newObject: newMesMaterialDTO, "Добавление материала MES", "Материал MES: ", _authorizationRepository);

                        resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана. Материал добавлен с кодом " + newMesMaterialDTO.Id.ToString();
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                        loadFromExcelPage.console.Log(resultString);
                        loadFromExcelPage.RefreshSate();
                        rowNumber++;
                        continue;
                    }
                    else  // изменение
                    {
                        MesMaterialDTO forChangeMesMaterialDTO = new MesMaterialDTO();
                        forChangeMesMaterialDTO.Id = foundMesMaterialDTO.Id;
                        forChangeMesMaterialDTO.Code = foundMesMaterialDTO.Code;
                        if (String.IsNullOrEmpty(nameVarString))
                        {
                            forChangeMesMaterialDTO.Name = foundMesMaterialDTO.Name;
                        }
                        else
                        {
                            if (forChangeMesMaterialDTO.Name.Equals(nameVarString))
                            {
                                forChangeMesMaterialDTO.Name = foundMesMaterialDTO.Name;
                            }
                            else
                            {
                                var objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;
                                if (objectForCheckName != null)
                                {
                                    if (objectForCheckName.Id != foundMesMaterialDTO.Id)
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
                                    else
                                    {
                                        forChangeMesMaterialDTO.Name = nameVarString;
                                    }
                                }
                                else
                                {
                                    forChangeMesMaterialDTO.Name = nameVarString;
                                }
                            }
                        }


                        if (String.IsNullOrEmpty(shortNameVarString))
                        {
                            forChangeMesMaterialDTO.ShortName = foundMesMaterialDTO.ShortName;
                        }
                        else
                        {
                            if (foundMesMaterialDTO.ShortName.Equals(shortNameVarString))
                            {
                                forChangeMesMaterialDTO.ShortName = foundMesMaterialDTO.ShortName;
                            }
                            else
                            {
                                var objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;
                                if (objectForCheckShortName != null)
                                {
                                    if (objectForCheckShortName.Id != foundMesMaterialDTO.Id)
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
                                    else
                                    {
                                        forChangeMesMaterialDTO.ShortName = shortNameVarString;
                                    }
                                }
                                else
                                {
                                    forChangeMesMaterialDTO.ShortName = shortNameVarString;
                                }
                            }
                        }

                        if (String.IsNullOrEmpty(isArchiveVarString))
                        {
                            forChangeMesMaterialDTO.IsArchive = false;
                        }
                        else
                        {
                            forChangeMesMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                        }

                        await _mesMaterialRepository.Update(forChangeMesMaterialDTO, SD.UpdateMode.Update);

                        if (forChangeMesMaterialDTO.IsArchive != foundMesMaterialDTO.IsArchive)
                            if (forChangeMesMaterialDTO.IsArchive == true)
                            {
                                await _mesMaterialRepository.Update(forChangeMesMaterialDTO, SD.UpdateMode.MoveToArchive);
                            }
                            else
                            {
                                await _mesMaterialRepository.Update(forChangeMesMaterialDTO, SD.UpdateMode.RestoreFromArchive);
                            }

                        await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: foundMesMaterialDTO, newObject: forChangeMesMaterialDTO, "Изменение материала MES", "Материал MES: ", _authorizationRepository);

                        resultString = "OK. Строка  " + rowNumber.ToString() + " успешно обработана.";
                        worksheet.Cell(rowNumber, 7).Value = resultString;
                        worksheet.Cell(rowNumber, 7).Style.Font.FontColor = XLColor.Green;
                        loadFromExcelPage.console.Log(resultString);
                        rowNumber++;
                        continue;
                    }
                }
                loadFromExcelPage.RefreshSate();
                rowNumber++;
            }

            loadFromExcelPage.console.Log($"Окончание загрузки данных листа MesMaterial в справочник Материалов Mes");
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

    }
}
