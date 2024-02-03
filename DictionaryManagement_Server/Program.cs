using DictionaryManagement_Business.Repository;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_Common;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Server.Extensions.Repository;
using DictionaryManagement_Server.Extensions.Repository.IRepository;
using Microsoft.AspNetCore.Authentication.Negotiate;
using Microsoft.EntityFrameworkCore;
using Radzen;

var builder = WebApplication.CreateBuilder(args);


builder.WebHost.UseIISIntegration();

builder.Services.AddAuthentication(NegotiateDefaults.AuthenticationScheme)
   .AddNegotiate();

builder.Services.AddAuthorization(options =>
{
    options.FallbackPolicy = options.DefaultPolicy;
});


// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
/* AddHubOptions(options => options.MaximumReceiveMessageSize = 60000 * 1024);*/
builder.Services.AddMvc();


builder.WebHost.UseUrls("http://+:5555");
//builder.Services.AddHttpsRedirection(options => options.HttpsPort = 7777);


//SD.AppFactoryMode = SD.FactoryMode.KOS;
//if (builder.Configuration.GetValue<string>("FactoryMode") == "NKNH")
//    SD.AppFactoryMode = SD.FactoryMode.NKNH;

SD.AppFactoryMode = builder.Configuration.GetValue<string>("FactoryMode");
if (SD.AppFactoryMode != null)
{
    if (SD.AppFactoryMode.ToUpper().Contains(" ¿«¿Õ‹") || SD.AppFactoryMode.ToUpper().Equals(" Œ—"))
    {
        SD.AppFactoryModeShort = " Œ—";
    }
    else
    {
        if (SD.AppFactoryMode.ToUpper().Contains("Õ»∆Õ≈ ¿Ã— ") || SD.AppFactoryMode.ToUpper().Equals("Õ Õ’"))
        {
            SD.AppFactoryModeShort = "Õ Õ’";
        }
        else
            SD.AppFactoryModeShort = "";
    }
}
else
{
    SD.AppFactoryMode = "";
    SD.AppFactoryModeShort = "";
}

SD.ShowBackgroundText = builder.Configuration.GetValue<int>("ShowBackgroundText");

builder.Services.AddDbContext<IntDBApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("IntDBConnection"),
    u => u.CommandTimeout(SD.SqlCommandConnectionTimeout))
);


builder.Services.AddScoped<ISapEquipmentRepository, SapEquipmentRepository>();
builder.Services.AddScoped<ISapMaterialRepository, SapMaterialRepository>();
builder.Services.AddScoped<IMesMaterialRepository, MesMaterialRepository>();
builder.Services.AddScoped<IMesUnitOfMeasureRepository, MesUnitOfMeasureRepository>();
builder.Services.AddScoped<ISapUnitOfMeasureRepository, SapUnitOfMeasureRepository>();
builder.Services.AddScoped<ICorrectionReasonRepository, CorrectionReasonRepository>();

builder.Services.AddScoped<IMesParamSourceTypeRepository, MesParamSourceTypeRepository>();
builder.Services.AddScoped<IDataTypeRepository, DataTypeRepository>();
builder.Services.AddScoped<IDataSourceRepository, DataSourceRepository>();
builder.Services.AddScoped<IReportTemplateTypeRepository, ReportTemplateTypeRepository>();
builder.Services.AddScoped<ILogEventTypeRepository, LogEventTypeRepository>();
builder.Services.AddScoped<ISettingsRepository, SettingsRepository>();
builder.Services.AddScoped<IUnitOfMeasureSapToMesMappingRepository, UnitOfMeasureSapToMesMappingRepository>();
builder.Services.AddScoped<ISapToMesMaterialMappingRepository, SapToMesMaterialMappingRepository>();
builder.Services.AddScoped<IMesDepartmentRepository, MesDepartmentRepository>();
builder.Services.AddScoped<IMesParamRepository, MesParamRepository>();

builder.Services.AddScoped<IUserRepository, UserRepository>();
builder.Services.AddScoped<IRoleRepository, RoleRepository>();
builder.Services.AddScoped<IUserToRoleRepository, UserToRoleRepository>();
builder.Services.AddScoped<IRoleToDepartmentRepository, RoleToDepartmentRepository>();

builder.Services.AddScoped<IReportTemplateRepository, ReportTemplateRepository>();
builder.Services.AddScoped<IReportTemplateTypeToRoleRepository, ReportTemplateTypeToRoleRepository>();
builder.Services.AddScoped<IReportEntityRepository, ReportEntityRepository>();
builder.Services.AddScoped<IReportEntityLogRepository, ReportEntityLogRepository>();
builder.Services.AddScoped<ISmenaRepository, SmenaRepository>();
builder.Services.AddScoped<IRoleVMRepository, RoleVMRepository>();
builder.Services.AddScoped<IADGroupRepository, ADGroupRepository>();
builder.Services.AddScoped<IRoleToADGroupRepository, RoleToADGroupRepository>();
builder.Services.AddScoped<IAuthorizationRepository, AuthorizationRepository>();
builder.Services.AddScoped<IAuthorizationControllersRepository, AuthorizationControllersRepository>();
builder.Services.AddScoped<ISimpleExcelExportRepository, SimpleExcelExportRepository>();
builder.Services.AddScoped<IVersionRepository, VersionRepository>();
builder.Services.AddScoped<ISchedulerRepository, SchedulerRepository>();
builder.Services.AddScoped<IMesNdoStocksRepository, MesNdoStocksRepository>();
builder.Services.AddScoped<ISapNdoOUTRepository, SapNdoOUTRepository>();
builder.Services.AddScoped<IMesMovementsRepository, MesMovementsRepository>();
builder.Services.AddScoped<ISapMovementsINRepository, SapMovementsINRepository>();
builder.Services.AddScoped<ISapMovementsOUTRepository, SapMovementsOUTRepository>();
builder.Services.AddScoped<ILogEventRepository, LogEventRepository>();
builder.Services.AddScoped<ILoadFromExcelRepository, LoadFromExcelRepository>();
builder.Services.AddScoped<IReportEntityResendDatesRepository, ReportEntityResendDatesRepository>();


builder.Services.AddAutoMapper(AppDomain.CurrentDomain.GetAssemblies());
builder.Services.AddScoped<DialogService>();
builder.Services.AddScoped<NotificationService>();
builder.Services.AddScoped<TooltipService>();
builder.Services.AddScoped<ContextMenuService>();

System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
SD.AppVersion = fvi.FileVersion;

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    //app.UseForwardedHeaders();
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    //app.UseHsts();
}
else
    // app.UseForwardedHeaders();

    app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.UseEndpoints(endpoints =>
{
    endpoints.MapRazorPages();
    endpoints.MapControllerRoute(
        name: "default",
        pattern: "{controller=Home}/{action=Index}/{id?}");
});

app.Run();
