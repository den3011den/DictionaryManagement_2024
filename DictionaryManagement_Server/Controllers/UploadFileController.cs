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

        [HttpPost("UploadFileController/UploadReportTemplateFile/{reportTemplateId}")]
        [RequestSizeLimit(60000000)]
        public async Task<IActionResult> UploadReportTemplateFile(IFormFile file, Guid reportTemplateId)
        {
            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
                else
                {
                    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                    {
                        return StatusCode(401, "Вы не входите в группу " + SD.AdminRoleName + ". Доступ запрещён");
                    }
                }
            }
            catch
            {
                return StatusCode(401, "Не удалось проверить авторизацию. Вы не авторизованы. Доступ запрещён. Возможно авторизация отключена.");
            }


            string pathVar = (await _settingsRepository.GetByName("ReportTemplatePath")).Value;
            try
            {

                await UploadTemplateFile(file, reportTemplateId, pathVar);
                return StatusCode(200);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        public async Task UploadTemplateFile(IFormFile file, Guid reportTemplateGuid, string reportTemplatePath)
        {
            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return;
                }
                else
                {
                    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                    {
                        return;
                    }
                }
            }
            catch
            {
                return;
            }



            if (file != null && file.Length > 0)
            {
                var extension = Path.GetExtension(file.FileName);
                var fullPath = Path.Combine(reportTemplatePath, reportTemplateGuid.ToString() + extension);
                using (FileStream fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.ReadWrite/*, FileShare.ReadWrite, 800000000*/))
                {
                    await file.CopyToAsync(fileStream);

                }
            }
        }


        [HttpPost("UploadFileController/UploadReportEntityFile/{reportEntityId}")]
        [RequestSizeLimit(60000000)]
        public async Task<IActionResult> UploadReportEntityFile(IFormFile file, Guid reportEntityId)
        {
            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return StatusCode(401, "Вы не авторизованы. Доступ запрещён");
                }
                else
                {
                    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                    {
                        return StatusCode(401, "Вы не входите в группу " + SD.AdminRoleName + ". Доступ запрещён");
                    }
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
            try
            {
                if (!User.Identity.IsAuthenticated)
                {
                    return;
                }
                else
                {
                    if (!(await _authorizationControllersRepository.CurrentUserIsInAdminRoleByLogin(User.Identity.Name, SD.MessageBoxMode.Off)))
                    {
                        return;
                    }
                }
            }
            catch
            {
                return;
            }


            if (file != null && file.Length > 0)
            {
                var extension = Path.GetExtension(file.FileName);
                var fullPath = Path.Combine(reportEntityPath, reportEntityGuid.ToString() + extension);
                using (FileStream fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.ReadWrite/*, FileShare.ReadWrite, 800000000*/))
                {
                    await file.CopyToAsync(fileStream);
                    //fileStream.Close();
                    if (fileStream != null)
                        await fileStream.DisposeAsync();
                }
            }
        }

    }
}
