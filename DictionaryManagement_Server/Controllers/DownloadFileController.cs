using DictionaryManagement_Business.Repository;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_Models.IntDBModels;
using Microsoft.AspNetCore.Mvc;

namespace DictionaryManagement_Server.Controllers
{
    public class DownloadFileController : Controller
    {

        private readonly ISettingsRepository _settingsRepository;
        private readonly IReportTemplateRepository _reportTemplateRepository;
        private readonly IReportEntityRepository _reportEntityRepository;
        private readonly IReportTemplateTypeRepository _reportTemplateTypeRepository;
        private readonly IAuthorizationControllersRepository _authorizationControllersRepository;

        public DownloadFileController(ISettingsRepository settingsRepository, IReportTemplateRepository reportTemplateRepository,
            IReportEntityRepository reportEntityRepository, IReportTemplateTypeRepository reportTemplateTypeRepository,
            IAuthorizationControllersRepository authorizationControllersRepository)
        {
            _settingsRepository = settingsRepository;
            _reportTemplateRepository = reportTemplateRepository;
            _reportEntityRepository = reportEntityRepository;
            _reportTemplateTypeRepository = reportTemplateTypeRepository;
            _authorizationControllersRepository = authorizationControllersRepository;
        }

        [DisableRequestSizeLimit]
        [HttpGet("DownloadFileController/DownloadReportTemplateFile/{reportTemplateId}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> DownloadReportTemplateFile(Guid reportTemplateId)
        {

            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }


            string pathVar = (await _settingsRepository.GetByName("ReportTemplatePath")).Value;
            ReportTemplateDTO? foundTemplate = await _reportTemplateRepository.GetById(reportTemplateId);
            if (foundTemplate == null)
            {
                return StatusCode(500, "Запись по шаблону отчёта " + reportTemplateId.ToString() + " не найдена в БД");
            }


            string fileName = foundTemplate.TemplateFileName;
            string file = System.IO.Path.Combine(pathVar, fileName);
            var extension = Path.GetExtension(fileName);
            if (System.IO.File.Exists(file))
            {
                try
                {
                    var forFileName = SD.RemoveInvalidCharsFromFilename("Template_" + foundTemplate.ReportTemplateTypeDTOFK.Name + "_"
                        + foundTemplate.MesDepartmentDTOFK.ShortName + "_", 190) + SD.RemoveInvalidCharsFromFilename(fileName);
                    return File(new FileStream(file, FileMode.Open), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", forFileName /*+ extension*/);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, ex.Message);
                }
            }
            return StatusCode(500, "Файл " + file + " не найден");
        }

        [DisableRequestSizeLimit]
        [HttpGet("DownloadFileController/DownloadReportEntityDownloadFile/{reportEntityId}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> DownloadReportEntityDownloadFile(Guid reportEntityId)
        {

            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
                //else
                //{
                //    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                //    {
                //        return StatusCode(401, "Вы не входите в группу " + SD.AdminRoleName + ". Доступ запрещён");
                //    }
                //}
            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }


            string pathVar = (await _settingsRepository.GetByName("ReportDownloadPath")).Value;
            ReportEntityDTO? foundEntity = await _reportEntityRepository.GetById(reportEntityId);

            if (foundEntity == null)
            {
                return StatusCode(500, "Запись об экземпляре отчёта " + reportEntityId.ToString() + " не найдена в БД");
            }
            string fileName = foundEntity.DownloadReportFileName;
            string file = System.IO.Path.Combine(pathVar, fileName);
            var extension = Path.GetExtension(fileName);
            if (System.IO.File.Exists(file))
            {
                try
                {
                    var repTmplTypeDTO = await _reportTemplateTypeRepository.Get(foundEntity.ReportTemplateDTOFK.ReportTemplateTypeId);
                    var forFileName = (SD.RemoveInvalidCharsFromFilename("Download_" + repTmplTypeDTO.Name + "_"
                        + foundEntity.DownloadUserDTOFK.UserName
                        + "_" + foundEntity.DownloadTime.ToString() + "_", 190)
                        + SD.RemoveInvalidCharsFromFilename(fileName));
                    return File(new FileStream(file, FileMode.Open), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", forFileName /*+ extension*/);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, ex.Message);
                }
            }
            return StatusCode(500, "Файл " + file + " не найден");
        }

        [DisableRequestSizeLimit]
        [HttpGet("DownloadFileController/DownloadReportEntityUploadFile/{reportEntityId}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> DownloadReportEntityUploadFile(Guid reportEntityId)
        {

            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
                //else
                //{
                //    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                //    {
                //        return StatusCode(401, "Вы не входите в группу " + SD.AdminRoleName + ". Доступ запрещён");
                //    }
                //}

            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }
            string pathVar = (await _settingsRepository.GetByName("ReportUploadPath")).Value;
            ReportEntityDTO? foundEntity = await _reportEntityRepository.GetById(reportEntityId);
            //string fileName = foundEntity.UploadReportFileName;
            // в шестёрке решили в UploadReportFileName сохранять имя загружаемого пользователем файла
            // теперь приходится брать реально храняшееся имя файла из DownloadReportFileName
            if (foundEntity == null)
            {
                return StatusCode(500, "Запись об экземпляре отчёта " + reportEntityId.ToString() + " не найдена в БД");
            }
            string fileName = foundEntity.Id.ToString() + ".xlsx";
            string file = System.IO.Path.Combine(pathVar, fileName);
            var extension = Path.GetExtension(fileName);
            if (System.IO.File.Exists(file))
            {
                try
                {
                    var reportTemptateTypeDTO = await _reportTemplateTypeRepository.Get(foundEntity.ReportTemplateDTOFK.ReportTemplateTypeId);
                    var forFileName = (SD.RemoveInvalidCharsFromFilename("Upload_" + reportTemptateTypeDTO.Name + "_"
                        + foundEntity.UploadUserDTOFK.UserName
                        + "_" + foundEntity.UploadTime.ToString() + "_", 190)
                        + SD.RemoveInvalidCharsFromFilename(fileName))
                        .Replace(":", "_").Replace(",", "_").Replace("\"", "_").Replace("\'", "_");
                    return File(new FileStream(file, FileMode.Open), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", forFileName/* + extension*/);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, ex.Message);
                }
            }
            return StatusCode(500, "Файл " + file + " не найден");
        }

        [DisableRequestSizeLimit]
        [HttpGet("DownloadFileController/SimpleExcelExport/{filename}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> SimpleExcelExport(string filename)
        {

            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
                //else
                //{
                //    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                //    {
                //        return StatusCode(401, "Вы не входите в группу " + SD.AdminRoleName + ". Доступ запрещён");
                //    }
                //}
            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            string file = System.IO.Path.Combine(pathVar, filename);
            if (System.IO.File.Exists(file))
            {
                try
                {
                    var forFileName = SD.RemoveInvalidCharsFromFilename(filename);
                    return File(new FileStream(file, FileMode.Open), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", forFileName /*+ extension*/);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, ex.Message);
                }
            }
            return StatusCode(500, "Файл " + file + " не найден");
        }


        [DisableRequestSizeLimit]
        [HttpGet("DownloadFileController/DownloadReportTemplateFileByFileName/{reportTemplateFileName}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> DownloadReportTemplateFileByFileName(string reportTemplateFileName)
        {

            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
                if (String.IsNullOrEmpty(reportTemplateFileName))
                {
                    return StatusCode(500, "Пустое имя файла");
                }
            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }
            string pathVar = (await _settingsRepository.GetByName("ReportTemplatePath")).Value;
            string file = System.IO.Path.Combine(pathVar, reportTemplateFileName);
            if (System.IO.File.Exists(file))
            {
                try
                {
                    return File(new FileStream(file, FileMode.Open), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", reportTemplateFileName);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, ex.Message);
                }
            }
            return StatusCode(500, "Файл " + file + " не найден");
        }

        [DisableRequestSizeLimit]
        [HttpGet("DownloadFileController/DownloadAnyFile/{filePath}/{fileName}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> DownloadAnyFile(string filePath, string fileName)
        {

            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }
            string file = System.IO.Path.Combine(filePath.UnhideSlash(), fileName);
            if (System.IO.File.Exists(file))
            {
                try
                {
                    return File(new FileStream(file, FileMode.Open), "application/octet-stream", fileName);
                }
                catch (Exception ex)
                {
                    return StatusCode(500, ex.Message);
                }
            }
            return StatusCode(500, "Файл " + file + " не найден");
        }

    }
}
