using AutoMapper;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;

namespace DictionaryManagement_Business.Mapper
{
    public class MappingProfile : Profile
    {
        public MappingProfile()
        {
            CreateMap<SapEquipment, SapEquipmentDTO>().ReverseMap();
            CreateMap<SapMaterial, SapMaterialDTO>().ReverseMap();
            CreateMap<MesMaterial, MesMaterialDTO>().ReverseMap();
            CreateMap<MesUnitOfMeasure, MesUnitOfMeasureDTO>().ReverseMap();
            CreateMap<SapUnitOfMeasure, SapUnitOfMeasureDTO>().ReverseMap();
            CreateMap<CorrectionReason, CorrectionReasonDTO>().ReverseMap();

            CreateMap<MesParamSourceType, MesParamSourceTypeDTO>().ReverseMap();
            CreateMap<DataType, DataTypeDTO>()
                .ForMember(dest => dest.IsAutoCalcDestDataType, opt => opt.MapFrom(src => src.IsAutoCalcDestDataType == null ? false : src.IsAutoCalcDestDataType));
            CreateMap<DataTypeDTO, DataType>()
                .ForMember(dest => dest.IsAutoCalcDestDataType, opt => opt.MapFrom(src => src.IsAutoCalcDestDataType == null ? false : src.IsAutoCalcDestDataType));

            CreateMap<DataSource, DataSourceDTO>().ReverseMap();

            CreateMap<ReportTemplateType, ReportTemplateTypeDTO>()
                .ForMember(dest => dest.NeedAutoCalc, opt => opt.MapFrom(src => src.NeedAutoCalc == null ? false : src.NeedAutoCalc))
                .ForMember(dest => dest.CanAutoCalc, opt => opt.MapFrom(src => src.CanAutoCalc == null ? false : src.CanAutoCalc));
            CreateMap<ReportTemplateTypeDTO, ReportTemplateType>()
                .ForMember(dest => dest.NeedAutoCalc, opt => opt.MapFrom(src => src.NeedAutoCalc == null ? false : src.NeedAutoCalc))
                .ForMember(dest => dest.CanAutoCalc, opt => opt.MapFrom(src => src.CanAutoCalc == null ? false : src.CanAutoCalc));

            CreateMap<LogEventType, LogEventTypeDTO>().ReverseMap();
            CreateMap<Settings, SettingsDTO>().ReverseMap();
            CreateMap<User, UserDTO>().ReverseMap();


            //CreateMap<Role, RoleDTO>().ReverseMap();
            CreateMap<Role, RoleDTO>()
                .ForMember(dest => dest.IsAdmin, opt => opt.MapFrom(src => src.IsAdmin == null ? false : src.IsAdmin))
                .ForMember(dest => dest.IsAdminReadOnly, opt => opt.MapFrom(src => src.IsAdminReadOnly == null ? false : src.IsAdminReadOnly));
            CreateMap<RoleDTO, Role>()
                .ForMember(dest => dest.IsAdmin, opt => opt.MapFrom(src => src.IsAdmin == null ? false : src.IsAdmin))
                .ForMember(dest => dest.IsAdminReadOnly, opt => opt.MapFrom(src => src.IsAdminReadOnly == null ? false : src.IsAdminReadOnly));


            CreateMap<UnitOfMeasureSapToMesMapping, UnitOfMeasureSapToMesMappingDTO>().ReverseMap();

            CreateMap<UnitOfMeasureSapToMesMapping, UnitOfMeasureSapToMesMappingDTO>()
                    .ForMember(dest => dest.SapUnitOfMeasureDTO, opt => opt.MapFrom(src => src.SapUnitOfMeasure))
                    .ForMember(dest => dest.MesUnitOfMeasureDTO, opt => opt.MapFrom(src => src.MesUnitOfMeasure));

            CreateMap<UnitOfMeasureSapToMesMappingDTO, UnitOfMeasureSapToMesMapping>()
                    .ForMember(dest => dest.SapUnitOfMeasure, opt => opt.MapFrom(src => src.SapUnitOfMeasureDTO))
                    .ForMember(dest => dest.MesUnitOfMeasure, opt => opt.MapFrom(src => src.MesUnitOfMeasureDTO));


            CreateMap<SapToMesMaterialMapping, SapToMesMaterialMappingDTO>().ReverseMap();

            CreateMap<SapToMesMaterialMapping, SapToMesMaterialMappingDTO>()
                    .ForMember(dest => dest.SapMaterialDTO, opt => opt.MapFrom(src => src.SapMaterial))
                    .ForMember(dest => dest.MesMaterialDTO, opt => opt.MapFrom(src => src.MesMaterial));

            CreateMap<SapToMesMaterialMappingDTO, SapToMesMaterialMapping>()
                    .ForMember(dest => dest.SapMaterial, opt => opt.MapFrom(src => src.SapMaterialDTO))
                    .ForMember(dest => dest.MesMaterial, opt => opt.MapFrom(src => src.MesMaterialDTO));

            CreateMap<MesDepartmentDTO, MesDepartment>()
                .ForMember(dest => dest.DepartmentParent, opt => opt.MapFrom(src => src.DepartmentParentDTO));
            CreateMap<MesDepartment, MesDepartmentDTO>()
                .ForMember(dest => dest.DepartmentParentDTO, opt => opt.MapFrom(src => src.DepartmentParent));


            CreateMap<MesParamDTO, MesParam>()
                .ForMember(dest => dest.MesParamSourceTypeFK, opt => opt.MapFrom(src => src.MesParamSourceTypeDTOFK))
                .ForMember(dest => dest.MesDepartmentFK, opt => opt.MapFrom(src => src.MesDepartmentDTOFK))
                .ForMember(dest => dest.SapEquipmentSourceFK, opt => opt.MapFrom(src => src.SapEquipmentSourceDTOFK))
                .ForMember(dest => dest.SapEquipmentDestFK, opt => opt.MapFrom(src => src.SapEquipmentDestDTOFK))
                .ForMember(dest => dest.MesMaterialFK, opt => opt.MapFrom(src => src.MesMaterialDTOFK))
                .ForMember(dest => dest.SapMaterialFK, opt => opt.MapFrom(src => src.SapMaterialDTOFK))
                .ForMember(dest => dest.MesUnitOfMeasureFK, opt => opt.MapFrom(src => src.MesUnitOfMeasureDTOFK))
                .ForMember(dest => dest.SapUnitOfMeasureFK, opt => opt.MapFrom(src => src.SapUnitOfMeasureDTOFK));
            //.ForMember(dest => dest.MesParamSourceLink, opt => opt.MapFrom(src => src.MesParamSourceLink == "" ? null : src.MesParamSourceLink));

            CreateMap<MesParam, MesParamDTO>()
                .ForMember(dest => dest.MesParamSourceTypeDTOFK, opt => opt.MapFrom(src => src.MesParamSourceTypeFK))
                .ForMember(dest => dest.MesDepartmentDTOFK, opt => opt.MapFrom(src => src.MesDepartmentFK))
                .ForMember(dest => dest.SapEquipmentSourceDTOFK, opt => opt.MapFrom(src => src.SapEquipmentSourceFK))
                .ForMember(dest => dest.SapEquipmentDestDTOFK, opt => opt.MapFrom(src => src.SapEquipmentDestFK))
                .ForMember(dest => dest.MesMaterialDTOFK, opt => opt.MapFrom(src => src.MesMaterialFK))
                .ForMember(dest => dest.SapMaterialDTOFK, opt => opt.MapFrom(src => src.SapMaterialFK))
                .ForMember(dest => dest.MesUnitOfMeasureDTOFK, opt => opt.MapFrom(src => src.MesUnitOfMeasureFK))
                .ForMember(dest => dest.SapUnitOfMeasureDTOFK, opt => opt.MapFrom(src => src.SapUnitOfMeasureFK));
            //.ForMember(dest => dest.MesParamSourceLink, opt => opt.MapFrom(src => (src.MesParamSourceLink == null || src.MesParamSourceLink.ToUpper() == "NULL") ? "" : src.MesParamSourceLink));

            CreateMap<ReportTemplate, ReportTemplateDTO>()
                .ForMember(dest => dest.AddUserDTOFK, opt => opt.MapFrom(src => src.AddUserFK))
                .ForMember(dest => dest.ReportTemplateTypeDTOFK, opt => opt.MapFrom(src => src.ReportTemplateTypeFK))
                .ForMember(dest => dest.DestDataTypeDTOFK, opt => opt.MapFrom(src => src.DestDataTypeFK))
                .ForMember(dest => dest.MesDepartmentDTOFK, opt => opt.MapFrom(src => src.MesDepartmentFK))
                .ForMember(dest => dest.NeedAutoCalc, opt => opt.MapFrom(src => src.NeedAutoCalc == null ? false : src.NeedAutoCalc))
                .ForMember(dest => dest.AutoCalcOrder, opt => opt.MapFrom(src => src.AutoCalcOrder == null ? 1 : src.AutoCalcOrder))
                .ForMember(dest => dest.AutoCalcNumber, opt => opt.MapFrom(src => src.AutoCalcNumber == null ? 1 : src.AutoCalcNumber));

            CreateMap<ReportTemplateDTO, ReportTemplate>()
                    .ForMember(dest => dest.AddUserFK, opt => opt.MapFrom(src => src.AddUserDTOFK))
                    .ForMember(dest => dest.ReportTemplateTypeFK, opt => opt.MapFrom(src => src.ReportTemplateTypeDTOFK))
                    .ForMember(dest => dest.DestDataTypeFK, opt => opt.MapFrom(src => src.DestDataTypeDTOFK))
                    .ForMember(dest => dest.MesDepartmentFK, opt => opt.MapFrom(src => src.MesDepartmentDTOFK))
                    .ForMember(dest => dest.NeedAutoCalc, opt => opt.MapFrom(src => src.NeedAutoCalc == null ? false : src.NeedAutoCalc))
                    .ForMember(dest => dest.NeedAutoCalc, opt => opt.MapFrom(src => src.NeedAutoCalc == null ? false : src.NeedAutoCalc))
                    .ForMember(dest => dest.AutoCalcOrder, opt => opt.MapFrom(src => src.AutoCalcOrder == null ? 1 : src.AutoCalcOrder))
                    .ForMember(dest => dest.AutoCalcNumber, opt => opt.MapFrom(src => src.AutoCalcNumber == null ? 1 : src.AutoCalcNumber));


            CreateMap<ReportTemplateTypeTоRole, ReportTemplateTypeTоRoleDTO>()
                    .ForMember(dest => dest.ReportTemplateTypeDTOFK, opt => opt.MapFrom(src => src.ReportTemplateTypeFK))
                    .ForMember(dest => dest.RoleDTOFK, opt => opt.MapFrom(src => src.RoleFK));

            CreateMap<ReportTemplateTypeTоRoleDTO, ReportTemplateTypeTоRole>()
                    .ForMember(dest => dest.ReportTemplateTypeFK, opt => opt.MapFrom(src => src.ReportTemplateTypeDTOFK))
                    .ForMember(dest => dest.RoleFK, opt => opt.MapFrom(src => src.RoleDTOFK));

            CreateMap<RoleToDepartment, RoleToDepartmentDTO>()
                    .ForMember(dest => dest.RoleDTOFK, opt => opt.MapFrom(src => src.RoleFK))
                    .ForMember(dest => dest.DepartmentDTOFK, opt => opt.MapFrom(src => src.DepartmentFK));

            CreateMap<RoleToDepartmentDTO, RoleToDepartment>()
                    .ForMember(dest => dest.RoleFK, opt => opt.MapFrom(src => src.RoleDTOFK))
                    .ForMember(dest => dest.DepartmentFK, opt => opt.MapFrom(src => src.DepartmentDTOFK));

            CreateMap<ReportEntity, ReportEntityDTO>()
                    .ForMember(dest => dest.ReportTemplateDTOFK, opt => opt.MapFrom(src => src.ReportTemplateFK))
                    .ForMember(dest => dest.ReportDepartmentDTOFK, opt => opt.MapFrom(src => src.ReportDepartmentFK))
                    .ForMember(dest => dest.UploadUserDTOFK, opt => opt.MapFrom(src => src.UploadUserFK))
                    .ForMember(dest => dest.DownloadUserDTOFK, opt => opt.MapFrom(src => src.DownloadUserFK))
                    .ForMember(dest => dest.ReportEntityResendDatesListDTO, opt => opt.MapFrom(src => src.ReportEntityResendDatesList));

            CreateMap<ReportEntityDTO, ReportEntity>()
                    .ForMember(dest => dest.ReportTemplateFK, opt => opt.MapFrom(src => src.ReportTemplateDTOFK))
                    .ForMember(dest => dest.ReportDepartmentFK, opt => opt.MapFrom(src => src.ReportDepartmentDTOFK))
                    .ForMember(dest => dest.UploadUserFK, opt => opt.MapFrom(src => src.UploadUserDTOFK))
                    .ForMember(dest => dest.DownloadUserFK, opt => opt.MapFrom(src => src.DownloadUserDTOFK))
                    .ForMember(dest => dest.ReportEntityResendDatesList, opt => opt.MapFrom(src => src.ReportEntityResendDatesListDTO));

            CreateMap<ReportEntityLog, ReportEntityLogDTO>()
                    .ForMember(dest => dest.ReportEntityDTOFK, opt => opt.MapFrom(src => src.ReportEntityFK));

            CreateMap<ReportEntityLogDTO, ReportEntityLog>()
                    .ForMember(dest => dest.ReportEntityFK, opt => opt.MapFrom(src => src.ReportEntityDTOFK));

            CreateMap<Smena, SmenaDTO>()
                    .ForMember(dest => dest.DepartmentDTOFK, opt => opt.MapFrom(src => src.DepartmentFK));

            CreateMap<SmenaDTO, Smena>()
                    .ForMember(dest => dest.DepartmentFK, opt => opt.MapFrom(src => src.DepartmentDTOFK));

            CreateMap<RoleVMDTO, Role>().ReverseMap();
            CreateMap<ADGroupDTO, ADGroup>().ReverseMap();

            CreateMap<RoleVMDTO, RoleDTO>().ReverseMap();

            CreateMap<RoleToADGroup, RoleToADGroupDTO>()
                    .ForMember(dest => dest.RoleDTOFK, opt => opt.MapFrom(src => src.RoleFK))
                    .ForMember(dest => dest.ADGroupDTOFK, opt => opt.MapFrom(src => src.ADGroupFK));

            CreateMap<RoleToADGroupDTO, RoleToADGroup>()
                    .ForMember(dest => dest.RoleFK, opt => opt.MapFrom(src => src.RoleDTOFK))
                    .ForMember(dest => dest.ADGroupFK, opt => opt.MapFrom(src => src.ADGroupDTOFK));

            CreateMap<UserToRole, UserToRoleDTO>()
                    .ForMember(dest => dest.UserDTOFK, opt => opt.MapFrom(src => src.UserFK))
                    .ForMember(dest => dest.RoleDTOFK, opt => opt.MapFrom(src => src.RoleFK));

            CreateMap<UserToRoleDTO, UserToRole>()
                    .ForMember(dest => dest.UserFK, opt => opt.MapFrom(src => src.UserDTOFK))
                    .ForMember(dest => dest.RoleFK, opt => opt.MapFrom(src => src.RoleDTOFK));

            CreateMap<MesDepartmentDTO, MesDepartmentVMDTO>()
                .ForMember(dest => dest.ChildrenDTO, opt => opt.AllowNull())
                .ForMember(dest => dest.DepartmentParentVMDTO, opt => opt.MapFrom(src => src.DepartmentParentDTO))
                .ForMember(dest => dest.DepLevel, opt => opt.MapFrom(src => src.DepLevel));
            CreateMap<MesDepartmentVMDTO, MesDepartmentDTO>()
                .ForMember(dest => dest.DepartmentParentDTO, opt => opt.MapFrom(src => src.DepartmentParentVMDTO))
                .ForMember(dest => dest.DepLevel, opt => opt.MapFrom(src => src.DepLevel));

            CreateMap<VersionDTO, DictionaryManagement_DataAccess.Data.IntDB.Version>().ReverseMap();
            CreateMap<Scheduler, SchedulerDTO>().ReverseMap();

            CreateMap<MesNdoStocksDTO, MesNdoStocks>()
                .ForMember(dest => dest.AddUserFK, opt => opt.MapFrom(src => src.AddUserDTOFK))
                .ForMember(dest => dest.MesParamFK, opt => opt.MapFrom(src => src.MesParamDTOFK))
                .ForMember(dest => dest.ReportEntityFK, opt => opt.MapFrom(src => src.ReportEntityDTOFK))
                .ForMember(dest => dest.SapNdoOUTFK, opt => opt.MapFrom(src => src.SapNdoOUTDTOFK));

            CreateMap<MesNdoStocks, MesNdoStocksDTO>()
                .ForMember(dest => dest.AddUserDTOFK, opt => opt.MapFrom(src => src.AddUserFK))
                .ForMember(dest => dest.MesParamDTOFK, opt => opt.MapFrom(src => src.MesParamFK))
                .ForMember(dest => dest.ReportEntityDTOFK, opt => opt.MapFrom(src => src.ReportEntityFK))
                .ForMember(dest => dest.SapNdoOUTDTOFK, opt => opt.MapFrom(src => src.SapNdoOUTFK));

            CreateMap<SapNdoOUTDTO, SapNdoOUT>().ReverseMap();

            CreateMap<MesMovements, MesMovementsDTO>()
                .ForMember(dest => dest.AddUserDTOFK, opt => opt.MapFrom(src => src.AddUserFK))
                .ForMember(dest => dest.MesParamDTOFK, opt => opt.MapFrom(src => src.MesParamFK))
                .ForMember(dest => dest.SapMovementsOUTDTOFK, opt => opt.MapFrom(src => src.SapMovementsOUTFK))
                .ForMember(dest => dest.SapMovementsINDTOFK, opt => opt.MapFrom(src => src.SapMovementsINFK))
                .ForMember(dest => dest.DataSourceDTOFK, opt => opt.MapFrom(src => src.DataSourceFK))
                .ForMember(dest => dest.DataTypeDTOFK, opt => opt.MapFrom(src => src.DataTypeFK))
                .ForMember(dest => dest.ReportEntityDTOFK, opt => opt.MapFrom(src => src.ReportEntityFK))
                .ForMember(dest => dest.MesMovementsDTOFK, opt => opt.MapFrom(src => src.MesMovementsFK))
                .ForMember(dest => dest.MesMovementsCommentListDTO, opt => opt.MapFrom(src => src.MesMovementsCommentList));

            CreateMap<MesMovementsDTO, MesMovements>()
                .ForMember(dest => dest.AddUserFK, opt => opt.MapFrom(src => src.AddUserDTOFK))
                .ForMember(dest => dest.MesParamFK, opt => opt.MapFrom(src => src.MesParamDTOFK))
                .ForMember(dest => dest.SapMovementsOUTFK, opt => opt.MapFrom(src => src.SapMovementsOUTDTOFK))
                .ForMember(dest => dest.SapMovementsINFK, opt => opt.MapFrom(src => src.SapMovementsINDTOFK))
                .ForMember(dest => dest.DataSourceFK, opt => opt.MapFrom(src => src.DataSourceDTOFK))
                .ForMember(dest => dest.DataTypeFK, opt => opt.MapFrom(src => src.DataTypeDTOFK))
                .ForMember(dest => dest.ReportEntityFK, opt => opt.MapFrom(src => src.ReportEntityDTOFK))
                .ForMember(dest => dest.MesMovementsFK, opt => opt.MapFrom(src => src.MesMovementsDTOFK))
                .ForMember(dest => dest.MesMovementsCommentList, opt => opt.MapFrom(src => src.MesMovementsCommentListDTO));

            CreateMap<SapMovementsOUT, SapMovementsOUTDTO>()
                .ForMember(dest => dest.MesMovementsDTOFK, opt => opt.MapFrom(src => src.MesMovementsFK))
                .ForMember(dest => dest.MesParamDTOFK, opt => opt.MapFrom(src => src.MesParamFK))
                .ForMember(dest => dest.PreviousRecordDTOFK, opt => opt.MapFrom(src => src.PreviousRecordFK))
                .ForMember(dest => dest.SapMaterialDTOFK, opt => opt.MapFrom(src => src.SapMaterialFK))
                .ForMember(dest => dest.SapEquipmentSourceDTOFK, opt => opt.MapFrom(src => src.SapEquipmentSourceFK))
                .ForMember(dest => dest.SapEquipmentDestDTOFK, opt => opt.MapFrom(src => src.SapEquipmentDestFK));

            CreateMap<SapMovementsOUTDTO, SapMovementsOUT>()
                .ForMember(dest => dest.MesMovementsFK, opt => opt.MapFrom(src => src.MesMovementsDTOFK))
                .ForMember(dest => dest.MesParamFK, opt => opt.MapFrom(src => src.MesParamDTOFK))
                .ForMember(dest => dest.PreviousRecordFK, opt => opt.MapFrom(src => src.PreviousRecordDTOFK))
                .ForMember(dest => dest.SapMaterialFK, opt => opt.MapFrom(src => src.SapMaterialDTOFK))
                .ForMember(dest => dest.SapEquipmentSourceFK, opt => opt.MapFrom(src => src.SapEquipmentSourceDTOFK))
                .ForMember(dest => dest.SapEquipmentDestFK, opt => opt.MapFrom(src => src.SapEquipmentDestDTOFK));


            CreateMap<SapMovementsIN, SapMovementsINDTO>()
                .ForMember(dest => dest.PreviousRecordDTOFK, opt => opt.MapFrom(src => src.PreviousRecordFK))
                .ForMember(dest => dest.MesMovementDTOFK, opt => opt.MapFrom(src => src.MesMovementFK))
                .ForMember(dest => dest.SapMaterialDTOFK, opt => opt.MapFrom(src => src.SapMaterialFK))
                .ForMember(dest => dest.SapEquipmentSourceDTOFK, opt => opt.MapFrom(src => src.SapEquipmentSourceFK))
                .ForMember(dest => dest.SapEquipmentDestDTOFK, opt => opt.MapFrom(src => src.SapEquipmentDestFK));

            CreateMap<SapMovementsINDTO, SapMovementsIN>()
                .ForMember(dest => dest.PreviousRecordFK, opt => opt.MapFrom(src => src.PreviousRecordDTOFK))
                .ForMember(dest => dest.MesMovementFK, opt => opt.MapFrom(src => src.MesMovementDTOFK))
                .ForMember(dest => dest.SapMaterialFK, opt => opt.MapFrom(src => src.SapMaterialDTOFK))
                .ForMember(dest => dest.SapEquipmentSourceFK, opt => opt.MapFrom(src => src.SapEquipmentSourceDTOFK))
                .ForMember(dest => dest.SapEquipmentDestFK, opt => opt.MapFrom(src => src.SapEquipmentDestDTOFK));

            CreateMap<LogEvent, LogEventDTO>()
                .ForMember(dest => dest.LogEventTypeDTOFK, opt => opt.MapFrom(src => src.LogEventTypeFK))
                .ForMember(dest => dest.UserDTOFK, opt => opt.MapFrom(src => src.UserFK));

            CreateMap<LogEventDTO, LogEvent>()
                .ForMember(dest => dest.LogEventTypeFK, opt => opt.MapFrom(src => src.LogEventTypeDTOFK))
                .ForMember(dest => dest.UserFK, opt => opt.MapFrom(src => src.UserDTOFK));

            CreateMap<MesMovementsComment, MesMovementsCommentDTO>()
                .ForMember(dest => dest.MesMovementsDTOFK, opt => opt.MapFrom(src => src.MesMovementsFK))
                .ForMember(dest => dest.CorrectionReasonDTOFK, opt => opt.MapFrom(src => src.CorrectionReasonFK));

            CreateMap<MesMovementsCommentDTO, MesMovementsComment>()
                .ForMember(dest => dest.MesMovementsFK, opt => opt.MapFrom(src => src.MesMovementsDTOFK))
                .ForMember(dest => dest.CorrectionReasonFK, opt => opt.MapFrom(src => src.CorrectionReasonDTOFK));

            CreateMap<ReportEntityResendDates, ReportEntityResendDatesDTO>()
                .ForMember(dest => dest.ReportEntityDTOFK, opt => opt.MapFrom(src => src.ReportEntityFK));

            CreateMap<ReportEntityResendDatesDTO, ReportEntityResendDates>()
                .ForMember(dest => dest.ReportEntityFK, opt => opt.MapFrom(src => src.ReportEntityDTOFK));
        }
    }
}
