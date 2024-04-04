using Microsoft.JSInterop;
using System.Text.Json;

namespace DictionaryManagement_Common
{
    public static class SD
    {
        public enum CheckReportTemplateTagsType
        {
            IsDuplicate,
            IsNotInBase,
            IsInArchive,
            IsInOtherNotArchiveReportTemplatesBySheetName
        }
        public enum SelectDictionaryScope
        {
            All,
            ArchiveOnly,
            NotArchiveOnly
        }

        public enum UpdateMode
        {
            Update,
            MoveToArchive,
            RestoreFromArchive
        }

        public enum FactoryMode
        {
            NKNH,
            KOS
        }

        public enum MessageBoxMode
        {
            On,
            Off
        }

        public enum LoginReturnMode
        {
            LoginOnly,
            NameOnly,
            LoginAndName,
            LoginAndNameAndAccessMode
        }

        public enum AdminMode
        {
            IsAdmin,
            IsAdminReadOnly,
            None
        }

        public static string DictionaryManagementUserName = "DictionaryManagement";
        public static string AppVersion = "";
        public static string? AppFactoryMode = "КOC";
        public static string? AppFactoryModeShort = "КOC";
        public static int? ShowBackgroundText = 0;
        public static string AdminRoleName = "Администратор";
        public static string SyncUserADGroupsIntervalInMinutesSettingName = "SyncUserADGroupsIntervalInMinutes";

        public static int MaxAllowedExcelRows = 1048576;

        public enum EditMode
        {
            CreateBasedOnRow = 1,
            Create = 2,
            Edit = 3
        }

        public static string ReportDownloadPathSettingName = "ReportDownloadPath";
        public static string ReportUploadPathSettingName = "ReportUploadPath";
        public static string ReportInputSheetSettingName = "ReportInputSheet";
        public static string ReportOutputSheetSettingName = "ReportOutputSheet";
        public static string ReportStartEndDateSheetSettingName = "ReportStartEndDateSheet";
        public static string ReportReasonLiabrarySheetSettingName = "ReportReasonLiabrarySheet";
        public static string ReportTagLibrarySheetSettingName = "ReportTagLibrary";
        public static string ReportTemplatePathSettingName = "ReportTemplatePath";
        public static string ExcelWorkBookProtectionPasswordName = "ExcelWorkBookProtectionPassword";
        public static string TempFilePathSettingName = "TempFilePath";

        public static string SapMaterialLoadFromExcelReportTemplateTypeNameSettingName = "SapMaterialLoadFromExcelReportTemplateTypeName";
        public static string MesMaterialLoadFromExcelReportTemplateTypeNameSettingName = "MesMaterialLoadFromExcelReportTemplateTypeName";
        public static string SapEquipmentLoadFromExcelReportTemplateTypeNameSettingName = "SapEquipmentLoadFromExcelReportTemplateTypeName";
        public static string MesParamLoadFromExcelReportTemplateTypeNameSettingName = "MesParamLoadFromExcelReportTemplateTypeName";
        public static string MesMovementsLoadFromExcelReportTemplateTypeNameSettingName = "MesMovementsLoadFromExcelReportTemplateTypeName";
        public static string MesNdoStocksLoadFromExcelReportTemplateTypeNameSettingName = "MesNdoStocksLoadFromExcelReportTemplateTypeName";
        public static string SapMovementsINLoadFromExcelReportTemplateTypeNameSettingName = "SapMovementsINLoadFromExcelReportTemplateTypeName";
        public static string SapMovementsOUTLoadFromExcelReportTemplateTypeNameSettingName = "SapMovementsOUTLoadFromExcelReportTemplateTypeName";
        public static string SapNdoOUTLoadFromExcelReportTemplateTypeNameSettingName = "SapNdoOUTLoadFromExcelReportTemplateTypeName";
        public static string UsersLoadFromExcelReportTemplateTypeNameSettingName = "UsersLoadFromExcelReportTemplateTypeName";
        public static string ADGroupsLoadFromExcelReportTemplateTypeNameSettingName = "ADGroupsLoadFromExcelReportTemplateTypeName";
        public static string ControlMesParamSourceTypeImmutableSettingName = "ControlMesParamSourceTypeImmutable";
        public static string ControlDataSourceImmutableSettingName = "ControlDataSourceImmutable";
        public static string ControlDataTypeCantChangeNameSettingName = "ControlDataTypeCantChangeName";
        public static string ControlReportTemplateTypeCantChangeNameSettingName = "ControlReportTemplateTypeCantChangeName";
        public static string ShowFillReportTemplateToMesParamTableForAllReportTemplatesButtonSettingName = "ShowFillReportTemplateToMesParamTableForAllReportTemplatesButton";

        public static int SqlCommandConnectionTimeout = 180;

        public static string EmbReportTemplateTypeName = "Отчёт ЭМБ";
        public static string CorrectionReportTemplateTypeName = "Корректировки";
        public static string NdoReportTemplateTypeName = "Ручной ввод НДО";
        public static string TebReportTemplateTypeName = "Форма по Энергетике";
        public static string NeedCheckReportTemplateBeforeSaving = "NeedCheckReportTemplateBeforeSaving";

        public record SheetHeader
        {
            public string SheetHeaderColumnName { get; set; }
            public int SheetHeaderColumnNumber { get; set; }
        }

        public class SheetTemplate
        {
            public string SheetName { get; set; }
            public List<SheetHeader> SheetHeaderList { get; set; }
            public bool NeedCheckIsDuplicateTags { get; set; }
            public bool NeedCheckIsNotInBaseTags { get; set; }
            public bool NeedCheckIsInArchiveTags { get; set; }
            public bool NeedCheckSheetHeaders { get; set; }
            public bool NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName { get; set; }
        }

        public static List<SheetTemplate> EmbSheetsList = new List<SheetTemplate>
        {
            new SheetTemplate {SheetName = "StartEndDate", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "StartDate", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "EndDate", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "InputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Measured", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "MeasuredUpdated", SheetHeaderColumnNumber = 4},
                new SheetHeader { SheetHeaderColumnName = "Reconciled", SheetHeaderColumnNumber = 5},
                new SheetHeader { SheetHeaderColumnName = "GoneToSAPValue", SheetHeaderColumnNumber = 6},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 7},
                new SheetHeader { SheetHeaderColumnName = "CorrectionComment", SheetHeaderColumnNumber = 8},
                new SheetHeader { SheetHeaderColumnName = "AddUserNAME", SheetHeaderColumnNumber = 9}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "OutputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Value", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 4},
                new SheetHeader { SheetHeaderColumnName = "CorrectionComment", SheetHeaderColumnNumber = 5}
            },
            NeedCheckIsDuplicateTags = true,
            NeedCheckIsNotInBaseTags = true,
            NeedCheckIsInArchiveTags = true,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = true
            },
            new SheetTemplate {SheetName = "ReasonLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonName", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "Отчет матрица", SheetHeaderList = new List<SheetHeader>(),
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = false,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "TagLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1}
            },
            NeedCheckIsDuplicateTags = true,
            NeedCheckIsNotInBaseTags = true,
            NeedCheckIsInArchiveTags = true,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            }
        };

        public static List<SheetTemplate> TebSheetsList = new List<SheetTemplate>
        {
            new SheetTemplate {SheetName = "StartEndDate", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "StartDate", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "EndDate", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "InputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Measured", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "MeasuredUpdated", SheetHeaderColumnNumber = 4},
                new SheetHeader { SheetHeaderColumnName = "Reconciled", SheetHeaderColumnNumber = 5},
                new SheetHeader { SheetHeaderColumnName = "GoneToSAPValue", SheetHeaderColumnNumber = 6},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 7},
                new SheetHeader { SheetHeaderColumnName = "CorrectionComment", SheetHeaderColumnNumber = 8},
                new SheetHeader { SheetHeaderColumnName = "AddUserNAME", SheetHeaderColumnNumber = 9}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "OutputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Value", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 4},
                new SheetHeader { SheetHeaderColumnName = "CorrectionComment", SheetHeaderColumnNumber = 5}
            },
            NeedCheckIsDuplicateTags = true,
            NeedCheckIsNotInBaseTags = true,
            NeedCheckIsInArchiveTags = true,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = true
            },
            new SheetTemplate {SheetName = "ReasonLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonName", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "Отчет матрица", SheetHeaderList = new List<SheetHeader>(),
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = false,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "TagLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1}
            },
            NeedCheckIsDuplicateTags = true,
            NeedCheckIsNotInBaseTags = true,
            NeedCheckIsInArchiveTags = true,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            }
        };

        public static List<SheetTemplate> CorrectionSheetsList = new List<SheetTemplate>
        {
            new SheetTemplate {SheetName = "StartEndDate", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "StartDate", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "EndDate", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "InputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Measured", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "MeasuredUpdated", SheetHeaderColumnNumber = 4},
                new SheetHeader { SheetHeaderColumnName = "Reconciled", SheetHeaderColumnNumber = 5},
                new SheetHeader { SheetHeaderColumnName = "GoneToSAPValue", SheetHeaderColumnNumber = 6},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 7},
                new SheetHeader { SheetHeaderColumnName = "CorrectionComment", SheetHeaderColumnNumber = 8},
                new SheetHeader { SheetHeaderColumnName = "AddUserNAME", SheetHeaderColumnNumber = 9}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "OutputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Value", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 4},
                new SheetHeader { SheetHeaderColumnName = "CorrectionComment", SheetHeaderColumnNumber = 5}
            },
            NeedCheckIsDuplicateTags = true,
            NeedCheckIsNotInBaseTags = true,
            NeedCheckIsInArchiveTags = true,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = true
            },
            new SheetTemplate {SheetName = "ReasonLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonId", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "CorrectionReasonName", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "Форма корректировок", SheetHeaderList = new List<SheetHeader>(),
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = false,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "TagLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1}
            },
            NeedCheckIsDuplicateTags = true,
            NeedCheckIsNotInBaseTags = true,
            NeedCheckIsInArchiveTags = true,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            }
        };


        public static List<SheetTemplate> NdoSheetsList = new List<SheetTemplate>
        {
            new SheetTemplate {SheetName = "StartEndDate", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "StartDate", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "EndDate", SheetHeaderColumnNumber = 2}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "InputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "Value", SheetHeaderColumnNumber = 3},
                new SheetHeader { SheetHeaderColumnName = "ValueDifference", SheetHeaderColumnNumber = 4}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "OutputData", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "MesParamCode", SheetHeaderColumnNumber = 1},
                new SheetHeader { SheetHeaderColumnName = "ValueTime", SheetHeaderColumnNumber = 2},
                new SheetHeader { SheetHeaderColumnName = "ValueDifference", SheetHeaderColumnNumber = 3}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = true,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            },
            new SheetTemplate {SheetName = "TagLibrary", SheetHeaderList = new List<SheetHeader>()
            {
                new SheetHeader { SheetHeaderColumnName = "СписокТеговСИР", SheetHeaderColumnNumber = 1}
            },
            NeedCheckIsDuplicateTags = false,
            NeedCheckIsNotInBaseTags = false,
            NeedCheckIsInArchiveTags = false,
            NeedCheckSheetHeaders = false,
            NeedCheckExistenceTagsInOtherNotArchiveReportTemplatesBySheetName = false
            }
        };

        public static String RemoveInvalidCharsFromFilename(this String file_name, int? maxFileNameLength = 200)
        {
            foreach (Char invalid_char in Path.GetInvalidFileNameChars())
            {
                file_name = file_name.Replace(oldValue: invalid_char.ToString(), newValue: "_");
            }

            if (maxFileNameLength != null)
            {
                if ((file_name.Length - (int)maxFileNameLength) > 0)
                    file_name = file_name.Substring(file_name.Length - (int)maxFileNameLength);
            }
            return file_name;
        }

        public async static Task<bool> CheckPageSettingsVersion(string settingName, IJSRuntime iJSRuntime)
        {

            bool needDeleteSettings = false;

            string resultVersion = (await iJSRuntime.InvokeAsync<string>("window.localStorage.getItem", settingName + "Version"));
            if (resultVersion == null)
            {
                needDeleteSettings = true;
            }
            else
            {
                string? currVersion = JsonSerializer.Deserialize<string>(resultVersion);
                if (currVersion == null)
                {
                    needDeleteSettings = true;
                }
                else
                {
                    if (currVersion.Trim().ToUpper() != SD.AppVersion.Trim().ToUpper())
                    {
                        needDeleteSettings = true;
                    }
                }

                if (needDeleteSettings)
                {
                    await iJSRuntime.InvokeAsync<string>("window.localStorage.removeItem", settingName);
                }
            }
            return !needDeleteSettings;
        }
        public async static Task SetPageSettingsVersion(string settingName, IJSRuntime iJSRuntime)
        {
            await iJSRuntime.InvokeVoidAsync("eval", $@"window.localStorage.setItem('{settingName}Version',
                '{JsonSerializer.Serialize<string>(SD.AppVersion)}')");
        }

    }
}
