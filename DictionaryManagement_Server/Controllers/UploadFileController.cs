using DictionaryManagement_Business.Repository;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using Microsoft.AspNetCore.Mvc;

namespace DictionaryManagement_Server.Controllers
{
    public class UploadFileController : Controller
    {

        private readonly ISettingsRepository _settingsRepository;
        private readonly IAuthorizationControllersRepository _authorizationControllersRepository;

        public UploadFileController(ISettingsRepository settingsRepository,
            IAuthorizationControllersRepository authorizationControllersRepository)
        {
            _settingsRepository = settingsRepository;
            _authorizationControllersRepository = authorizationControllersRepository;
        }

        [DisableRequestSizeLimit]
        [HttpPost("UploadFileController/UploadReportTemplateFile/{reportTemplateId}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> UploadReportTemplateFile(IFormFile file, Guid reportTemplateId)
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

            string pathVar = (await _settingsRepository.GetByName("TempFilePath")).Value;
            try
            {

                Exception? exReturn = await UploadTemplateFile(file, reportTemplateId, pathVar);
                if (exReturn == null)
                {
                    return StatusCode(200);
                }
                else
                {
                    return StatusCode(500, exReturn.Message);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        public async Task<Exception?> UploadTemplateFile(IFormFile file, Guid reportTemplateGuid, string reportTemplatePath)
        {
            try
            {
                if (file != null && file.Length > 0)
                {
                    var extension = Path.GetExtension(file.FileName);
                    var fullPath = Path.Combine(reportTemplatePath, reportTemplateGuid.ToString() + extension);
                    using (FileStream fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.ReadWrite/*, FileShare.ReadWrite, 800000000*/))
                    {
                        file.CopyTo(fileStream);
                    }
                }
            }
            catch (Exception ex)
            {
                return ex;
            }
            return null;
        }

        [DisableRequestSizeLimit]
        [HttpPost("UploadFileController/UploadReportEntityFile/{reportEntityId}")]
        //[RequestSizeLimit(2147483648)]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> UploadReportEntityFile(IFormFile file, Guid reportEntityId)
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


            string pathVar = (await _settingsRepository.GetByName(SD.ReportUploadPathSettingName)).Value;
            try
            {

                await UploadEntityFile(file, reportEntityId, pathVar);
                return StatusCode(200);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        public async Task UploadEntityFile(IFormFile file, Guid reportEntityGuid, string reportEntityPath)
        {
            if (file != null && file.Length > 0)
            {
                var extension = Path.GetExtension(file.FileName);
                var fullPath = Path.Combine(reportEntityPath, reportEntityGuid.ToString() + extension);
                using (FileStream fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.ReadWrite/*, FileShare.ReadWrite, 800000000*/))
                {
                    file.CopyTo(fileStream);
                    //fileStream.Close();
                    if (fileStream != null)
                        await fileStream.DisposeAsync();
                }
            }
        }



        [DisableRequestSizeLimit]
        [HttpPost("UploadFileController/UploadAnyFile/{filePath}/{fileName}")]
        //[RequestSizeLimit(60000000)]
        public async Task<IActionResult> UploadAnyFile(IFormFile file, string filePath, string fileName)
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

            try
            {
                if (file != null && file.Length > 0)
                {
                    using (FileStream fileStream = new FileStream(filePath.UnhideSlash() + fileName, FileMode.Create, FileAccess.ReadWrite))
                    {
                        file.CopyTo(fileStream);
                    }
                    return StatusCode(200);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
            return StatusCode(200);
        }
    }
}