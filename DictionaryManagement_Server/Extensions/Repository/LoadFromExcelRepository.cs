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

        public async Task<bool> MaterialExcelFileLoad(Shared.LoadFromExcel? loadFromExcelPage, IXLWorksheet worksheet,
            IAuthorizationRepository _authorizationRepository)
        {
            bool haveErrors = false;

            loadFromExcelPage.console.Log($"Лист SapMaterial загружен в память");
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

                MaterialDTO? foundMaterialDTO = null;
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

                    MaterialDTO changedMaterialDTO = new MaterialDTO();
                    changedMaterialDTO.Id = foundMaterialDTO.Id;

                    if (String.IsNullOrEmpty(codeVarString))
                    {
                        changedMaterialDTO.Code = foundMaterialDTO.Code;
                    }
                    else
                    {
                        if (foundMaterialDTO.Code.Equals(codeVarString))
                        {
                            changedMaterialDTO.Code = foundMaterialDTO.Code;
                        }
                        else
                        {
                            MaterialDTO? objectForCheckCode;
                            if (worksheet.Name.Equals("SapMaterial"))
                                objectForCheckCode = _sapMaterialRepository.GetByCode(codeVarString).Result;
                            else
                                objectForCheckCode = _mesMaterialRepository.GetByCode(codeVarString).Result;

                            if (objectForCheckCode != null)
                            {
                                if (objectForCheckCode.Id != foundMaterialDTO.Id)
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
                                    changedMaterialDTO.Code = codeVarString;
                                }
                            }
                            else
                            {
                                changedMaterialDTO.Code = codeVarString;
                            }
                        }
                    }


                    if (String.IsNullOrEmpty(nameVarString))
                    {
                        changedMaterialDTO.Name = foundMaterialDTO.Name;
                    }
                    else
                    {
                        if (foundMaterialDTO.Name.Equals(nameVarString))
                        {
                            changedMaterialDTO.Name = foundMaterialDTO.Name;
                        }
                        else
                        {
                            MaterialDTO? objectForCheckName;
                            if (worksheet.Name.Equals("SapMaterial"))
                                objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                            else
                                objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;

                            if (objectForCheckName != null)
                            {
                                if (objectForCheckName.Id != foundMaterialDTO.Id)
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
                                    changedMaterialDTO.Name = nameVarString;
                                }
                            }
                            else
                            {
                                changedMaterialDTO.Name = nameVarString;
                            }
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
                            MaterialDTO? objectForCheckShortName;
                            if (worksheet.Name.Equals("SapMaterial"))
                                objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                            else
                                objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;
                            if (objectForCheckShortName != null)
                            {
                                if (objectForCheckShortName.Id != foundMaterialDTO.Id)
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
                                    changedMaterialDTO.ShortName = shortNameVarString;
                                }
                            }
                            else
                            {
                                changedMaterialDTO.ShortName = shortNameVarString;
                            }
                        }
                    }

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
                        if (changedMaterialDTO.IsArchive == true)
                        {
                            if (worksheet.Name.Equals("SapMaterial"))
                                await _sapMaterialRepository.Update(new SapMaterialDTO(changedMaterialDTO), SD.UpdateMode.MoveToArchive);
                            else
                                await _mesMaterialRepository.Update(new MesMaterialDTO(changedMaterialDTO), SD.UpdateMode.MoveToArchive);
                        }
                        else
                        {
                            if (worksheet.Name.Equals("SapMaterial"))
                                await _sapMaterialRepository.Update(new SapMaterialDTO(changedMaterialDTO), SD.UpdateMode.RestoreFromArchive);
                            else
                                await _mesMaterialRepository.Update(new MesMaterialDTO(changedMaterialDTO), SD.UpdateMode.RestoreFromArchive);
                        }

                    if (worksheet.Name.Equals("SapMaterial"))
                        await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: new SapMaterialDTO(foundMaterialDTO), newObject: new SapMaterialDTO(changedMaterialDTO), "Изменение материала SAP", "Материал SAP: ", _authorizationRepository);
                    else
                        await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: new MesMaterialDTO(foundMaterialDTO), newObject: new MesMaterialDTO(changedMaterialDTO), "Изменение материала MES", "Материал MES: ", _authorizationRepository);

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
                    if (worksheet.Name.Equals("SapMaterial"))
                        foundMaterialDTO = await _sapMaterialRepository.GetByCode(codeVarString);
                    else
                        foundMaterialDTO = await _mesMaterialRepository.GetByCode(codeVarString);
                    if (foundMaterialDTO == null)  // добавление
                    {

                        MaterialDTO forAddMaterialDTO = new MaterialDTO();

                        MaterialDTO? objectForCheckCode;
                        if (worksheet.Name.Equals("SapMaterial"))
                            objectForCheckCode = _sapMaterialRepository.GetByCode(codeVarString).Result;
                        else
                            objectForCheckCode = _mesMaterialRepository.GetByCode(codeVarString).Result;

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

                        forAddMaterialDTO.Code = codeVarString;

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
                            MaterialDTO? objectForCheckName;
                            if (worksheet.Name.Equals("SapMaterial"))
                                objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                            else
                                objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;

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
                            forAddMaterialDTO.Name = nameVarString;
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
                            MaterialDTO? objectForCheckShortName;
                            if (worksheet.Name.Equals("SapMaterial"))
                                objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                            else
                                objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;

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
                            forAddMaterialDTO.ShortName = shortNameVarString;
                        }

                        forAddMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;

                        MaterialDTO? newMaterialDTO;
                        if (worksheet.Name.Equals("SapMaterial"))
                        {
                            newMaterialDTO = await _sapMaterialRepository.Create(new SapMaterialDTO(forAddMaterialDTO));
                            await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: null, newObject: new SapMaterialDTO(newMaterialDTO), "Добавление материала SAP", "Материал SAP: ", _authorizationRepository);
                        }
                        else
                        {
                            newMaterialDTO = await _mesMaterialRepository.Create(new MesMaterialDTO(forAddMaterialDTO));
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
                    else  // изменение
                    {
                        MaterialDTO forChangeMaterialDTO = new MaterialDTO();
                        forChangeMaterialDTO.Id = foundMaterialDTO.Id;
                        forChangeMaterialDTO.Code = foundMaterialDTO.Code;
                        if (String.IsNullOrEmpty(nameVarString))
                        {
                            forChangeMaterialDTO.Name = foundMaterialDTO.Name;
                        }
                        else
                        {
                            if (forChangeMaterialDTO.Name.Equals(nameVarString))
                            {
                                forChangeMaterialDTO.Name = foundMaterialDTO.Name;
                            }
                            else
                            {
                                MaterialDTO? objectForCheckName;
                                if (worksheet.Name.Equals("SapMaterial"))
                                    objectForCheckName = _sapMaterialRepository.GetByName(nameVarString).Result;
                                else
                                    objectForCheckName = _mesMaterialRepository.GetByName(nameVarString).Result;

                                if (objectForCheckName != null)
                                {
                                    if (objectForCheckName.Id != foundMaterialDTO.Id)
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
                                        forChangeMaterialDTO.Name = nameVarString;
                                    }
                                }
                                else
                                {
                                    forChangeMaterialDTO.Name = nameVarString;
                                }
                            }
                        }


                        if (String.IsNullOrEmpty(shortNameVarString))
                        {
                            forChangeMaterialDTO.ShortName = foundMaterialDTO.ShortName;
                        }
                        else
                        {
                            if (foundMaterialDTO.ShortName.Equals(shortNameVarString))
                            {
                                forChangeMaterialDTO.ShortName = foundMaterialDTO.ShortName;
                            }
                            else
                            {
                                MaterialDTO? objectForCheckShortName;
                                if (worksheet.Name.Equals("SapMaterial"))
                                    objectForCheckShortName = _sapMaterialRepository.GetByShortName(shortNameVarString).Result;
                                else
                                    objectForCheckShortName = _mesMaterialRepository.GetByShortName(shortNameVarString).Result;

                                if (objectForCheckShortName != null)
                                {
                                    if (objectForCheckShortName.Id != foundMaterialDTO.Id)
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
                                        forChangeMaterialDTO.ShortName = shortNameVarString;
                                    }
                                }
                                else
                                {
                                    forChangeMaterialDTO.ShortName = shortNameVarString;
                                }
                            }
                        }

                        if (String.IsNullOrEmpty(isArchiveVarString))
                        {
                            forChangeMaterialDTO.IsArchive = false;
                        }
                        else
                        {
                            forChangeMaterialDTO.IsArchive = isArchiveVarString.ToUpper().Equals("ДА") ? true : false;
                        }

                        if (worksheet.Name.Equals("SapMaterial"))
                            await _sapMaterialRepository.Update(new SapMaterialDTO(forChangeMaterialDTO), SD.UpdateMode.Update);
                        else
                            await _mesMaterialRepository.Update(new MesMaterialDTO(forChangeMaterialDTO), SD.UpdateMode.Update);

                        if (forChangeMaterialDTO.IsArchive != foundMaterialDTO.IsArchive)
                        {
                            SD.UpdateMode updMode;
                            if (forChangeMaterialDTO.IsArchive == true)
                            {
                                updMode = SD.UpdateMode.MoveToArchive;
                            }
                            else
                            {
                                updMode = SD.UpdateMode.RestoreFromArchive;
                            }
                            if (worksheet.Name.Equals("SapMaterial"))
                            {
                                await _sapMaterialRepository.Update(new SapMaterialDTO(forChangeMaterialDTO), updMode);
                                await _logEventRepository.ToLog<SapMaterialDTO>(oldObject: new SapMaterialDTO(foundMaterialDTO), newObject: new SapMaterialDTO(forChangeMaterialDTO), "Изменение материала SAP", "Материал SAP: ", _authorizationRepository);
                            }
                            else
                            {
                                await _mesMaterialRepository.Update(new MesMaterialDTO(forChangeMaterialDTO), updMode);
                                await _logEventRepository.ToLog<MesMaterialDTO>(oldObject: new MesMaterialDTO(foundMaterialDTO), newObject: new MesMaterialDTO(forChangeMaterialDTO), "Изменение материала MES", "Материал MES: ", _authorizationRepository);
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
                }
                rowNumber++;
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
